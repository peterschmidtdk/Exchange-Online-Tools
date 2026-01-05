<#
.SYNOPSIS
  Mailbox SOA Manager (GUI) for Exchange Online - Cloud-managed Exchange attributes (SOA) toggle.

.DESCRIPTION
  GUI tool to view and change mailbox Exchange attribute SOA state via IsExchangeCloudManaged:
    - Enable cloud management     : IsExchangeCloudManaged = $true
    - Revert to on-prem management: IsExchangeCloudManaged = $false

  Browse/Search:
    - Loads all mailboxes into local cache for fast browsing/search
    - Paging (Prev/Next + Page size + indicator)
    - Search uses cached list (reliable + fast)
    - GUI shows version
    - Connected status shows tenant name (best-effort)

REFERENCE
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

LOGGING
  - Single logfile only (append; never overwritten)
  - Timestamp on every line
  - SOA changes logged with BEFORE/AFTER + Actor
  - RunId included for correlation

REQUIREMENTS
  - Windows PowerShell 5.1 OR PowerShell 7+ (must run in STA for WinForms)
  - Module: ExchangeOnlineManagement

AUTHOR
  Peter Schmidt

VERSION
  1.8.1 (2026-01-05)
    - FIX: Load all mailboxes no longer fails with "You cannot call a method on a null-valued expression"
    - FIX: Removed PS7-only null-coalescing operator '??' (PS5.1 compatible)
#>

# --- Load WinForms early ---
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    [System.Windows.Forms.Application]::EnableVisualStyles()
} catch {
    Write-Error "Failed to load WinForms assemblies. Error: $($_.Exception.Message)"
    return
}

#region Globals
$Script:ScriptVersion = "1.8.1"
$Script:ToolName      = "Mailbox SOA Manager"
$Script:RunId         = [Guid]::NewGuid().ToString()

# Use current location (as in your script) for Logs/Exports
$Script:LogDir        = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:ExportDir     = Join-Path -Path (Get-Location) -ChildPath "Exports"   # created on-demand
$Script:LogFile       = Join-Path -Path $Script:LogDir -ChildPath "MailboxSOAManager.log"

$Script:IsConnected   = $false
$Script:ExoActor      = $null
$Script:TenantName    = $null

# Cache + paging
$Script:MailboxCache     = @()
$Script:CurrentView      = @()
$Script:CacheLoaded      = $false
$Script:PageSize         = 50
$Script:PageIndex        = 0
$Script:CurrentQueryText = ""

New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
#endregion

#region Logging
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $winUser = "$env:USERDOMAIN\$env:USERNAME"
    $actor = if ($Script:ExoActor) { $Script:ExoActor } else { "EXO:unknown" }
    $tenant = if ($Script:TenantName) { $Script:TenantName } else { "Tenant:unknown" }
    $line = "[$ts][$Level][RunId:$($Script:RunId)][Win:$winUser][$actor][$tenant] $Message"
    Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8
}
#endregion

#region STA guard (AUTO RELAUNCH)
function Ensure-STA {
    try {
        $apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
        if ($apt -eq [System.Threading.ApartmentState]::STA) { return $true }

        Write-Log "Not running in STA mode (ApartmentState=$apt). Attempting self-relaunch in STA..." "WARN"

        $scriptPath = $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrWhiteSpace($scriptPath) -or -not (Test-Path $scriptPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "This GUI must run in STA mode.`nRun:`n  powershell.exe -STA -File .\ExchangeSOAManger-GUI.ps1",
                $Script:ToolName,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $false
        }

        $exe = if ($PSVersionTable.PSEdition -eq "Core") { "pwsh.exe" } else { "powershell.exe" }
        $args = @("-NoProfile","-ExecutionPolicy","Bypass","-STA","-File","`"$scriptPath`"") -join " "
        Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory (Split-Path -Parent $scriptPath) | Out-Null
        Write-Log ("Launched new process: {0} {1}" -f $exe, $args) "INFO"
        return $false
    } catch {
        Write-Log "Ensure-STA failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to validate STA mode.`n$($_.Exception.Message)",
            $Script:ToolName,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}
if (-not (Ensure-STA)) { return }
#endregion

Write-Log "$($Script:ToolName) v$($Script:ScriptVersion) starting (GUI init)..." "INFO"

#region Module helpers
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Required module '$Name' is not installed.`nInstall it now (CurrentUser)?",
            "Missing Module",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
            throw "Module '$Name' not installed."
        }

        Write-Log "Installing module '$Name' (Scope=CurrentUser)..." "INFO"
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log "Module '$Name' installed." "INFO"
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log "Module loaded: $Name" "INFO"
}
#endregion

#region Helpers
function Safe-ToString {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    try { return $Value.ToString() } catch { return "" }
}

function Get-PropValue {
    param(
        [Parameter(Mandatory)][object]$Object,
        [Parameter(Mandatory)][string]$Name
    )
    try {
        if ($null -eq $Object) { return $null }
        $p = $Object.PSObject.Properties[$Name]
        if ($null -ne $p) { return $p.Value }
    } catch { }
    return $null
}

function Get-SOAIndicator {
    param([object]$IsExchangeCloudManaged)
    if ($IsExchangeCloudManaged -eq $true)  { return "☁ Online" }
    if ($IsExchangeCloudManaged -eq $false) { return "🏢 On-Prem" }
    return "? Unknown"
}

function Convert-ToGridRow {
    param([Parameter(Mandatory)]$MailboxObject)

    if ($null -eq $MailboxObject) { return $null }

    $display = Safe-ToString (Get-PropValue $MailboxObject "DisplayName")
    $alias   = Safe-ToString (Get-PropValue $MailboxObject "Alias")
    $psa     = Safe-ToString (Get-PropValue $MailboxObject "PrimarySmtpAddress")

    $rtd     = Safe-ToString (Get-PropValue $MailboxObject "RecipientTypeDetails")
    $dirSync = Get-PropValue $MailboxObject "IsDirSynced"
    $cloud   = Get-PropValue $MailboxObject "IsExchangeCloudManaged"

    [PSCustomObject]@{
        DisplayName                 = $display
        Alias                       = $alias
        PrimarySmtpAddress          = $psa
        RecipientTypeDetails        = $rtd
        IsDirSynced                 = $dirSync
        IsExchangeCloudManaged      = $cloud
        SOAIndicator                = (Get-SOAIndicator $cloud)
    }
}

function Text-Matches {
    param([object]$Text,[string]$Query)
    if ([string]::IsNullOrWhiteSpace($Query)) { return $true }
    if ($null -eq $Text) { return $false }
    return ($Text.ToString().IndexOf($Query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Get-ExoActorBestEffort {
    try {
        $ci = Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($ci) {
            $info = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
            if ($info -and $info.UserPrincipalName) { return "EXO:$($info.UserPrincipalName)" }
        }
    } catch { }
    return "EXO:unknown"
}

function Get-TenantNameBestEffort {
    # Best effort: Get-OrganizationConfig name (works for most orgs)
    try {
        $org = Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1
        if ($org -and $org.Name) { return $org.Name }
    } catch { }

    # Fallback: UPN domain
    try {
        if ($Script:ExoActor -and $Script:ExoActor -like "EXO:*") {
            $upn = $Script:ExoActor.Replace("EXO:","")
            if ($upn -like "*@*") { return ($upn.Split("@")[-1]) }
        }
    } catch { }

    return "Unknown"
}

function Get-AllMailboxesSafe {
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $errors = New-Object System.Collections.Generic.List[string]

    $exoCmd = Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue
    if ($exoCmd) {
        try {
            $splat = @{
                ResultSize  = 'Unlimited'
                ErrorAction = 'Stop'
            }

            if ($exoCmd.Parameters.ContainsKey('PropertySets')) {
                $splat['PropertySets'] = 'Minimum'
            }

            if ($exoCmd.Parameters.ContainsKey('Properties')) {
                $splat['Properties'] = @(
                    'DisplayName','Alias','PrimarySmtpAddress',
                    'RecipientTypeDetails','IsDirSynced','IsExchangeCloudManaged'
                )
            }

            Write-Log ("Get-AllMailboxesSafe: Trying Get-EXOMailbox with params: {0}" -f ($splat.Keys -join ',')) "INFO"
            $raw = @(Get-EXOMailbox @splat)

            Write-Log "Get-AllMailboxesSafe: Get-EXOMailbox returned count=$($raw.Count)" "INFO"
            if ($raw.Count -gt 0) { return $raw }
            $errors.Add("Get-EXOMailbox returned 0 objects.")
        } catch {
            $msg = "Get-EXOMailbox failed: $($_.Exception.Message)"
            $errors.Add($msg)
            Write-Log "Get-AllMailboxesSafe: $msg" "WARN"
        }
    } else {
        $errors.Add("Get-EXOMailbox not available in this session/module.")
        Write-Log "Get-AllMailboxesSafe: Get-EXOMailbox not available." "INFO"
    }

    # Fallback
    try {
        Write-Log "Get-AllMailboxesSafe: Falling back to Get-Mailbox -ResultSize Unlimited" "INFO"
        $raw2 = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
        Write-Log "Get-AllMailboxesSafe: Get-Mailbox returned count=$($raw2.Count)" "INFO"
        return $raw2
    } catch {
        $msg2 = "Get-Mailbox fallback failed: $($_.Exception.Message)"
        $errors.Add($msg2)
        Write-Log "Get-AllMailboxesSafe: $msg2" "ERROR"
        $combined = ($errors | Select-Object -Unique) -join "`r`n- "
        throw "Load all mailboxes failed.`r`n- $combined"
    }
}
#endregion

#region Mailbox ops
function Get-MailboxDetails {
    param([Parameter(Mandatory)][string]$Identity)
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,Alias,PrimarySmtpAddress,RecipientTypeDetails,IsDirSynced,IsExchangeCloudManaged,ExchangeGuid,ExternalDirectoryObjectId

    $usr = $null
    try {
        $usr = Get-User -Identity $Identity -ErrorAction Stop |
            Select-Object DisplayName,UserPrincipalName,ImmutableId,RecipientTypeDetails,WhenChangedUTC
    } catch {
        Write-Log "Get-User failed (non-fatal) for '$Identity': $($_.Exception.Message)" "WARN"
    }

    [PSCustomObject]@{ Mailbox = $mbx; User = $usr }
}

function Export-MailboxSOASettings {
    param([Parameter(Mandatory)][string]$Identity)

    New-Item -ItemType Directory -Path $Script:ExportDir -Force | Out-Null

    $details = Get-MailboxDetails -Identity $Identity
    $mbx = $details.Mailbox
    $soa = Get-SOAIndicator $mbx.IsExchangeCloudManaged

    $smtpSafe = Safe-ToString $mbx.PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($smtpSafe)) { $smtpSafe = ($Identity -replace '[^a-zA-Z0-9\.\-_@]','_') }

    $safeId  = ($smtpSafe -replace '[^a-zA-Z0-9\.\-_@]','_')
    $stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
    $path    = Join-Path $Script:ExportDir "$safeId-MailboxSOASettings-$stamp.json"

    $export = [PSCustomObject]@{
        ExportType     = "Mailbox SOA Settings"
        ExportedAt     = (Get-Date).ToString("o")
        ToolName       = $Script:ToolName
        ToolVersion    = $Script:ScriptVersion
        RunId          = $Script:RunId
        TenantName     = $Script:TenantName
        Identity       = $Identity
        SOASettings    = [PSCustomObject]@{
            PrimarySmtpAddress      = $smtpSafe
            DisplayName             = Safe-ToString $mbx.DisplayName
            IsDirSynced             = $mbx.IsDirSynced
            IsExchangeCloudManaged  = $mbx.IsExchangeCloudManaged
            SOAIndicator            = $soa
        }
        MailboxDetails = $mbx
        UserDetails    = $details.User
    }

    $export | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8
    Write-Log "Export-MailboxSOASettings completed for '$Identity'. Path='$path' SOA='$soa' IsExchangeCloudManaged='$($mbx.IsExchangeCloudManaged)'" "INFO"
    return $path
}

function Set-MailboxSOACloudManaged {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][bool]$EnableCloudManaged
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $targetValue = [bool]$EnableCloudManaged

    $before = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    Write-Log "SOA change requested for '$Identity'. TargetIsExchangeCloudManaged=$targetValue (Before=$($before.IsExchangeCloudManaged); IsDirSynced=$($before.IsDirSynced))" "INFO"

    if ($before.IsDirSynced -ne $true) {
        $msg = "Mailbox '$Identity' is not DirSynced (IsDirSynced=$($before.IsDirSynced)). Change blocked."
        Write-Log $msg "WARN"
        throw $msg
    }

    if ($before.IsExchangeCloudManaged -eq $targetValue) {
        $msg = "No change needed for '$Identity'. IsExchangeCloudManaged already '$targetValue'."
        Write-Log $msg "INFO"
        return $msg
    }

    Set-Mailbox -Identity $Identity -IsExchangeCloudManaged $targetValue -ErrorAction Stop
    Write-Log "Set-Mailbox executed for '$Identity' IsExchangeCloudManaged=$targetValue" "INFO"

    $after = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    $changed = ($after.IsExchangeCloudManaged -eq $targetValue)
    Write-Log "SOA change result for '$Identity'. Before=$($before.IsExchangeCloudManaged) After=$($after.IsExchangeCloudManaged) Expected=$targetValue Success=$changed" "INFO"

    if (-not $changed) {
        return "Executed, but verification did not match expected value. Before='$($before.IsExchangeCloudManaged)' After='$($after.IsExchangeCloudManaged)' Expected='$targetValue'."
    }
    return "Updated. IsExchangeCloudManaged is now '$($after.IsExchangeCloudManaged)'."
}
#endregion

#region Cache + Paging
function Reset-ViewToCache {
    $Script:CurrentView = @($Script:MailboxCache)
    $Script:PageIndex = 0
    $Script:CurrentQueryText = ""
}

function Apply-SearchToCache {
    param([string]$QueryText)

    $q = ""
    if ($null -ne $QueryText) { $q = $QueryText.Trim() }

    $Script:CurrentQueryText = $q
    $Script:PageIndex = 0

    if ([string]::IsNullOrWhiteSpace($q)) {
        $Script:CurrentView = @($Script:MailboxCache)
        return
    }

    $Script:CurrentView = @(
        $Script:MailboxCache | Where-Object {
            (Text-Matches $_.DisplayName $q) -or
            (Text-Matches $_.Alias $q) -or
            (Text-Matches $_.PrimarySmtpAddress $q)
        }
    )
}

function Get-PageSlice {
    param([array]$Items,[int]$PageIndex,[int]$PageSize)

    if (-not $Items -or $Items.Count -eq 0) { return @() }
    if ($PageSize -le 0) { $PageSize = 50 }

    $count = $Items.Count
    $start = $PageIndex * $PageSize
    if ($start -ge $count) { return @() }

    $end = [Math]::Min($start + $PageSize - 1, $count - 1)
    return @($Items[$start..$end])
}

function Get-TotalPages {
    param([array]$Items,[int]$PageSize)
    if (-not $Items -or $Items.Count -eq 0) { return 0 }
    if ($PageSize -le 0) { $PageSize = 50 }
    return [int][Math]::Ceiling($Items.Count / [double]$PageSize)
}

function New-MailboxGridDataTable {
    $dt = New-Object System.Data.DataTable "Mailboxes"
    [void]$dt.Columns.Add("DisplayName", [string])
    [void]$dt.Columns.Add("Alias", [string])
    [void]$dt.Columns.Add("PrimarySmtpAddress", [string])
    [void]$dt.Columns.Add("RecipientTypeDetails", [string])
    [void]$dt.Columns.Add("IsDirSynced", [string])
    [void]$dt.Columns.Add("IsExchangeCloudManaged", [string])
    [void]$dt.Columns.Add("SOA (Exchange Attributes)", [string])
    return $dt
}

function Convert-PageToDataTable {
    param([array]$PageItems)

    $dt = New-MailboxGridDataTable

    $items = @()
    if ($PageItems) { $items = @($PageItems) } else { $items = @() }

    foreach ($x in $items) {
        if ($null -eq $x) { continue }

        $row = $dt.NewRow()
        $row["DisplayName"]              = Safe-ToString $x.DisplayName
        $row["Alias"]                    = Safe-ToString $x.Alias
        $row["PrimarySmtpAddress"]       = Safe-ToString $x.PrimarySmtpAddress
        $row["RecipientTypeDetails"]     = Safe-ToString $x.RecipientTypeDetails
        $row["IsDirSynced"]              = if ($null -eq $x.IsDirSynced) { "" } else { Safe-ToString $x.IsDirSynced }
        $row["IsExchangeCloudManaged"]   = if ($null -eq $x.IsExchangeCloudManaged) { "" } else { Safe-ToString $x.IsExchangeCloudManaged }
        $row["SOA (Exchange Attributes)"] = Safe-ToString $x.SOAIndicator
        [void]$dt.Rows.Add($row)
    }

    return $dt
}
#endregion

#region EXO connect/disconnect
function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"
        Write-Log "Connecting to Exchange Online..." "INFO"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null

        $Script:IsConnected = $true
        $Script:ExoActor    = Get-ExoActorBestEffort
        $Script:TenantName  = Get-TenantNameBestEffort

        Write-Log "Connected to Exchange Online. Tenant='$($Script:TenantName)'" "INFO"
        return $true
    } catch {
        $Script:IsConnected = $false
        $Script:ExoActor = "EXO:unknown"
        $Script:TenantName = $null
        Write-Log "Connect-EXO failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange Online.`n$($_.Exception.Message)",
            "Connect Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}

function Disconnect-EXO {
    try {
        if ($Script:IsConnected) {
            Write-Log "Disconnecting from Exchange Online..." "INFO"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Disconnected from Exchange Online." "INFO"
        }
    } catch {
        Write-Log "Disconnect-EXO warning: $($_.Exception.Message)" "WARN"
    } finally {
        $Script:IsConnected = $false
        $Script:ExoActor = $null
        $Script:TenantName = $null
        $Script:MailboxCache     = @()
        $Script:CurrentView      = @()
        $Script:CacheLoaded      = $false
        $Script:PageIndex        = 0
        $Script:CurrentQueryText = ""
    }
}
#endregion

#region GUI
try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($Script:ToolName) v$($Script:ScriptVersion) - Exchange Online (IsExchangeCloudManaged)"
    $form.Size = New-Object System.Drawing.Size(1100, 700)
    $form.StartPosition = "CenterScreen"

    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Text = "Connect"
    $btnConnect.Location = New-Object System.Drawing.Point(12, 12)
    $btnConnect.Size = New-Object System.Drawing.Size(110, 30)

    $btnDisconnect = New-Object System.Windows.Forms.Button
    $btnDisconnect.Text = "Disconnect"
    $btnDisconnect.Location = New-Object System.Drawing.Point(130, 12)
    $btnDisconnect.Size = New-Object System.Drawing.Size(110, 30)
    $btnDisconnect.Enabled = $false

    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Text = "Status: Not connected"
    $lblConn.Location = New-Object System.Drawing.Point(260, 18)
    $lblConn.AutoSize = $true

    $grpBrowse = New-Object System.Windows.Forms.GroupBox
    $grpBrowse.Text = "Browse & Search"
    $grpBrowse.Location = New-Object System.Drawing.Point(12, 55)
    $grpBrowse.Size = New-Object System.Drawing.Size(1060, 170)

    $btnLoadAll = New-Object System.Windows.Forms.Button
    $btnLoadAll.Text = "Load all mailboxes"
    $btnLoadAll.Location = New-Object System.Drawing.Point(16, 30)
    $btnLoadAll.Size = New-Object System.Drawing.Size(170, 30)
    $btnLoadAll.Enabled = $false

    $lblLoadHint = New-Object System.Windows.Forms.Label
    $lblLoadHint.Text = "Tip: Load all first for fast browsing/search."
    $lblLoadHint.Location = New-Object System.Drawing.Point(200, 36)
    $lblLoadHint.AutoSize = $true

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(16, 70)
    $txtSearch.Size = New-Object System.Drawing.Size(520, 25)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Search"
    $btnSearch.Location = New-Object System.Drawing.Point(548, 67)
    $btnSearch.Size = New-Object System.Drawing.Size(110, 30)
    $btnSearch.Enabled = $false

    $btnClearSearch = New-Object System.Windows.Forms.Button
    $btnClearSearch.Text = "Clear"
    $btnClearSearch.Location = New-Object System.Drawing.Point(666, 67)
    $btnClearSearch.Size = New-Object System.Drawing.Size(110, 30)
    $btnClearSearch.Enabled = $false

    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = "◀ Prev"
    $btnPrev.Location = New-Object System.Drawing.Point(16, 110)
    $btnPrev.Size = New-Object System.Drawing.Size(90, 30)
    $btnPrev.Enabled = $false

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = "Next ▶"
    $btnNext.Location = New-Object System.Drawing.Point(112, 110)
    $btnNext.Size = New-Object System.Drawing.Size(90, 30)
    $btnNext.Enabled = $false

    $lblPage = New-Object System.Windows.Forms.Label
    $lblPage.Text = "Page: -"
    $lblPage.Location = New-Object System.Drawing.Point(220, 116)
    $lblPage.AutoSize = $true

    $lblPageSize = New-Object System.Windows.Forms.Label
    $lblPageSize.Text = "Page size:"
    $lblPageSize.Location = New-Object System.Drawing.Point(360, 116)
    $lblPageSize.AutoSize = $true

    $cmbPageSize = New-Object System.Windows.Forms.ComboBox
    $cmbPageSize.Location = New-Object System.Drawing.Point(430, 112)
    $cmbPageSize.Size = New-Object System.Drawing.Size(90, 25)
    $cmbPageSize.DropDownStyle = 'DropDownList'
    [void]$cmbPageSize.Items.AddRange(@("25","50","100","200"))
    $cmbPageSize.SelectedItem = "50"
    $cmbPageSize.Enabled = $false

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Count: -"
    $lblCount.Location = New-Object System.Drawing.Point(548, 116)
    $lblCount.AutoSize = $true

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(16, 145)
    $lblStatus.Size = New-Object System.Drawing.Size(1020, 20)

    # Grid (bind to DataTable)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(16, 235)
    $grid.Size = New-Object System.Drawing.Size(1056, 220)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.AutoGenerateColumns = $true

    $Script:GridBinding = New-Object System.Windows.Forms.BindingSource
    $grid.DataSource = $Script:GridBinding

    $grpDetails = New-Object System.Windows.Forms.GroupBox
    $grpDetails.Text = "Selected mailbox details"
    $grpDetails.Location = New-Object System.Drawing.Point(12, 465)
    $grpDetails.Size = New-Object System.Drawing.Size(1060, 170)

    $txtDetails = New-Object System.Windows.Forms.TextBox
    $txtDetails.Location = New-Object System.Drawing.Point(16, 25)
    $txtDetails.Size = New-Object System.Drawing.Size(1028, 95)
    $txtDetails.Multiline = $true
    $txtDetails.ScrollBars = "Vertical"
    $txtDetails.ReadOnly = $true

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh details"
    $btnRefresh.Location = New-Object System.Drawing.Point(16, 130)
    $btnRefresh.Size = New-Object System.Drawing.Size(160, 30)
    $btnRefresh.Enabled = $false

    $btnExportSOA = New-Object System.Windows.Forms.Button
    $btnExportSOA.Text = "Export current mailbox SOA settings (JSON)"
    $btnExportSOA.Location = New-Object System.Drawing.Point(186, 130)
    $btnExportSOA.Size = New-Object System.Drawing.Size(300, 30)
    $btnExportSOA.Enabled = $false

    $btnEnableCloud = New-Object System.Windows.Forms.Button
    $btnEnableCloud.Text = "Enable cloud SOA (true)"
    $btnEnableCloud.Location = New-Object System.Drawing.Point(500, 130)
    $btnEnableCloud.Size = New-Object System.Drawing.Size(180, 30)
    $btnEnableCloud.Enabled = $false

    $btnRevertOnPrem = New-Object System.Windows.Forms.Button
    $btnRevertOnPrem.Text = "Revert to on-prem SOA (false)"
    $btnRevertOnPrem.Location = New-Object System.Drawing.Point(690, 130)
    $btnRevertOnPrem.Size = New-Object System.Drawing.Size(220, 30)
    $btnRevertOnPrem.Enabled = $false

    $lblFoot = New-Object System.Windows.Forms.Label
    $lblFoot.Location = New-Object System.Drawing.Point(12, 640)
    $lblFoot.Size = New-Object System.Drawing.Size(1060, 22)
    $lblFoot.Text = "v$($Script:ScriptVersion) | Log: $($Script:LogFile) | Export: .\Exports (created only when exporting)"
    $lblFoot.AutoSize = $false

    $form.Controls.AddRange(@($btnConnect,$btnDisconnect,$lblConn,$grpBrowse,$grid,$grpDetails,$lblFoot))
    $grpBrowse.Controls.AddRange(@(
        $btnLoadAll,$lblLoadHint,$txtSearch,$btnSearch,$btnClearSearch,
        $btnPrev,$btnNext,$lblPage,$lblPageSize,$cmbPageSize,$lblCount,$lblStatus
    ))
    $grpDetails.Controls.AddRange(@($txtDetails,$btnRefresh,$btnExportSOA,$btnEnableCloud,$btnRevertOnPrem))

    $Script:SelectedIdentity = $null

    function Update-PagingUI {
        $totalPages = Get-TotalPages -Items $Script:CurrentView -PageSize $Script:PageSize
        $totalItems = if ($Script:CurrentView) { $Script:CurrentView.Count } else { 0 }

        if ($totalPages -eq 0) {
            $lblPage.Text = "Page: -"
            $lblCount.Text = "Count: 0"
            $btnPrev.Enabled = $false
            $btnNext.Enabled = $false
            return
        }

        if ($Script:PageIndex -lt 0) { $Script:PageIndex = 0 }
        if ($Script:PageIndex -gt ($totalPages - 1)) { $Script:PageIndex = $totalPages - 1 }

        $lblPage.Text = "Page: $($Script:PageIndex + 1) / $totalPages"
        $lblCount.Text = "Count: $totalItems"
        $btnPrev.Enabled = ($Script:PageIndex -gt 0)
        $btnNext.Enabled = ($Script:PageIndex -lt ($totalPages - 1))
    }

    function Bind-GridFromCurrentView {
        # HARDEN: never let binding source receive $null
        $view = @()
        if ($Script:CurrentView) { $view = @($Script:CurrentView) } else { $view = @() }

        $pageItems = Get-PageSlice -Items $view -PageIndex $Script:PageIndex -PageSize $Script:PageSize
        if (-not $pageItems) { $pageItems = @() }

        $dt = Convert-PageToDataTable -PageItems $pageItems
        if ($null -eq $dt) { $dt = New-MailboxGridDataTable }

        $Script:GridBinding.DataSource = $dt
        $Script:GridBinding.ResetBindings($true)
        Update-PagingUI
    }

    function Reset-SelectionAndDetails {
        $Script:SelectedIdentity = $null
        $txtDetails.Clear()
        $btnRefresh.Enabled = $false
        $btnExportSOA.Enabled = $false
        $btnEnableCloud.Enabled = $false
        $btnRevertOnPrem.Enabled = $false
    }

    function Set-UiConnectedState {
        param([bool]$Connected)

        $btnConnect.Enabled      = -not $Connected
        $btnDisconnect.Enabled   = $Connected
        $btnLoadAll.Enabled      = $Connected
        $btnSearch.Enabled       = $Connected
        $cmbPageSize.Enabled     = $Connected

        if (-not $Connected) {
            $lblConn.Text = "Status: Not connected"
            $Script:GridBinding.DataSource = New-MailboxGridDataTable
            $Script:GridBinding.ResetBindings($true)
            Reset-SelectionAndDetails
            $lblStatus.Text = ""
            $lblPage.Text = "Page: -"
            $lblCount.Text = "Count: -"
            $btnPrev.Enabled = $false
            $btnNext.Enabled = $false
            $btnClearSearch.Enabled = $false
        } else {
            $tn = if ($Script:TenantName) { $Script:TenantName } else { "Unknown" }
            $lblConn.Text = "Status: Connected to Exchange Online (Tenant: $tn)"
            $lblStatus.Text = "Connected. Click 'Load all mailboxes' to browse/search."
        }
    }

    function Show-Details {
        param([string]$Identity)

        $details = Get-MailboxDetails -Identity $Identity
        $mbx = $details.Mailbox
        $usr = $details.User
        $soa = Get-SOAIndicator $mbx.IsExchangeCloudManaged

        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("Mailbox:")
        $lines.Add("  DisplayName               : $($mbx.DisplayName)")
        $lines.Add("  PrimarySmtpAddress        : $($mbx.PrimarySmtpAddress)")
        $lines.Add("  RecipientTypeDetails      : $($mbx.RecipientTypeDetails)")
        $lines.Add("  IsDirSynced               : $($mbx.IsDirSynced)")
        $lines.Add("  IsExchangeCloudManaged    : $($mbx.IsExchangeCloudManaged)")
        $lines.Add("  SOA (Exchange Attributes) : $soa")
        $lines.Add("  ExchangeGuid              : $($mbx.ExchangeGuid)")
        $lines.Add("  ExternalDirectoryObjectId : $($mbx.ExternalDirectoryObjectId)")

        if ($usr) {
            $lines.Add("")
            $lines.Add("User:")
            $lines.Add("  UserPrincipalName         : $($usr.UserPrincipalName)")
            $lines.Add("  ImmutableId               : $($usr.ImmutableId)")
            $lines.Add("  WhenChangedUTC            : $($usr.WhenChangedUTC)")
        }

        $txtDetails.Lines = $lines.ToArray()
        $btnRefresh.Enabled      = $true
        $btnExportSOA.Enabled    = $true
        $btnEnableCloud.Enabled  = $true
        $btnRevertOnPrem.Enabled = $true
    }

    # Events
    $btnConnect.Add_Click({
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if (Connect-EXO) { Set-UiConnectedState -Connected $true }
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnDisconnect.Add_Click({
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Disconnect-EXO
            Set-UiConnectedState -Connected $false
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $cmbPageSize.Add_SelectedIndexChanged({
        try {
            $Script:PageSize = [int]$cmbPageSize.SelectedItem
            $Script:PageIndex = 0
            Write-Log "PageSize changed to $($Script:PageSize)" "INFO"
            Bind-GridFromCurrentView
            Reset-SelectionAndDetails
        } catch { }
    })

    $btnPrev.Add_Click({
        if ($Script:PageIndex -gt 0) {
            $Script:PageIndex--
            Write-Log "Paging Prev. PageIndex=$($Script:PageIndex)" "INFO"
            Bind-GridFromCurrentView
            Reset-SelectionAndDetails
        }
    })

    $btnNext.Add_Click({
        $totalPages = Get-TotalPages -Items $Script:CurrentView -PageSize $Script:PageSize
        if ($Script:PageIndex -lt ($totalPages - 1)) {
            $Script:PageIndex++
            Write-Log "Paging Next. PageIndex=$($Script:PageIndex)" "INFO"
            Bind-GridFromCurrentView
            Reset-SelectionAndDetails
        }
    })

    $btnLoadAll.Add_Click({
        if (-not $Script:IsConnected) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Load ALL mailboxes into local cache?`nThis enables fast browsing and searching.",
            "Load all mailboxes",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $lblStatus.Text = "Loading all mailboxes..."
            $form.Refresh()

            Write-Log "LoadAll clicked." "INFO"
            $raw = @(Get-AllMailboxesSafe)
            if (-not $raw) { $raw = @() }

            $cache = New-Object System.Collections.Generic.List[object]
            foreach ($m in $raw) {
                if ($null -eq $m) { continue }
                $row = Convert-ToGridRow $m
                if ($null -ne $row) { [void]$cache.Add($row) }
            }

            $Script:MailboxCache = @($cache)
            $Script:CacheLoaded  = $true

            Write-Log "LoadAll: Converted to cache rows count=$($Script:MailboxCache.Count)" "INFO"

            Reset-ViewToCache
            Bind-GridFromCurrentView
            Reset-SelectionAndDetails

            $btnClearSearch.Enabled = $true
            $lblStatus.Text = "Loaded $($Script:MailboxCache.Count) mailboxes. Use paging + search."
            Write-Log "LoadAll success. CachedCount=$($Script:MailboxCache.Count)" "INFO"

            if ($Script:MailboxCache.Count -eq 0) {
                $lblStatus.Text = "Loaded 0 mailboxes. Check RBAC permissions and log file."
                Write-Log "LoadAll WARNING: Cache is empty (0). Likely permission/RBAC or command returned none." "WARN"
            }
        } catch {
            Write-Log "LoadAll failed: $($_.Exception.Message)" "ERROR"
            $lblStatus.Text = "Load all failed. Check log for details."
            [System.Windows.Forms.MessageBox]::Show(
                "Load all mailboxes failed.`n`n$($_.Exception.Message)`n`nLog:`n$($Script:LogFile)",
                "Load all failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnSearch.Add_Click({
        if (-not $Script:IsConnected) { return }

        $qTrim = ""
        if ($null -ne $txtSearch.Text) { $qTrim = $txtSearch.Text.Trim() }
        Write-Log "Search clicked. Query='$qTrim' CacheLoaded=$($Script:CacheLoaded)" "INFO"

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if (-not $Script:CacheLoaded) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please click 'Load all mailboxes' first to enable reliable searching and browsing.",
                    "Search",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
                return
            }

            Apply-SearchToCache -QueryText $qTrim
            Bind-GridFromCurrentView
            Reset-SelectionAndDetails

            $matches = if ($Script:CurrentView) { $Script:CurrentView.Count } else { 0 }
            if ([string]::IsNullOrWhiteSpace($qTrim)) {
                $lblStatus.Text = "Showing all cached mailboxes ($matches)."
            } else {
                $lblStatus.Text = "Search '$qTrim' matched $matches mailbox(es)."
            }
            $btnClearSearch.Enabled = $true
        } catch {
            Write-Log "Search failed: $($_.Exception.Message)" "ERROR"
            $lblStatus.Text = "Search failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "Search failed.`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnClearSearch.Add_Click({
        $txtSearch.Text = ""
        Write-Log "ClearSearch clicked." "INFO"
        if ($Script:CacheLoaded) {
            Reset-ViewToCache
            Bind-GridFromCurrentView
            $lblStatus.Text = "Showing all cached mailboxes ($($Script:MailboxCache.Count))."
        } else {
            $Script:GridBinding.DataSource = New-MailboxGridDataTable
            $Script:GridBinding.ResetBindings($true)
            $lblStatus.Text = "Cleared results."
        }
        Reset-SelectionAndDetails
    })

    # Selection (DataTable)
    $grid.Add_SelectionChanged({
        try {
            if ($grid.SelectedRows.Count -gt 0) {
                $row = $grid.SelectedRows[0]
                $smtp = $row.Cells["PrimarySmtpAddress"].Value
                if ($smtp) {
                    $Script:SelectedIdentity = $smtp.ToString()
                    Write-Log "Selection changed. SelectedIdentity='$($Script:SelectedIdentity)'" "INFO"
                    Show-Details -Identity $Script:SelectedIdentity
                }
            }
        } catch {
            Write-Log "SelectionChanged warning: $($_.Exception.Message)" "WARN"
        }
    })

    $btnRefresh.Add_Click({
        if (-not $Script:SelectedIdentity) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try { Show-Details -Identity $Script:SelectedIdentity }
        catch {
            Write-Log "Refresh failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Refresh failed.`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnExportSOA.Add_Click({
        if (-not $Script:SelectedIdentity) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $path = Export-MailboxSOASettings -Identity $Script:SelectedIdentity
            [System.Windows.Forms.MessageBox]::Show(
                "Exported CURRENT mailbox SOA settings to:`n$path",
                "Export Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } catch {
            Write-Log "Export failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Export failed.`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnEnableCloud.Add_Click({
        if (-not $Script:SelectedIdentity) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Enable cloud SOA (IsExchangeCloudManaged = TRUE) for:`n`n$($Script:SelectedIdentity)",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $true
            [System.Windows.Forms.MessageBox]::Show($msg,"Done",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            Show-Details -Identity $Script:SelectedIdentity

            if ($Script:CacheLoaded) {
                $updated = Convert-ToGridRow (Get-Mailbox -Identity $Script:SelectedIdentity -ErrorAction Stop)
                for ($i=0; $i -lt $Script:MailboxCache.Count; $i++) {
                    if ($Script:MailboxCache[$i].PrimarySmtpAddress -eq $Script:SelectedIdentity) {
                        $Script:MailboxCache[$i] = $updated
                        break
                    }
                }
                Apply-SearchToCache -QueryText $txtSearch.Text
                Bind-GridFromCurrentView
            }
        } catch {
            Write-Log "Enable cloud SOA failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Enable cloud SOA failed.`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnRevertOnPrem.Add_Click({
        if (-not $Script:SelectedIdentity) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Revert to on-prem SOA (IsExchangeCloudManaged = FALSE) for:`n`n$($Script:SelectedIdentity)",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $false
            [System.Windows.Forms.MessageBox]::Show($msg,"Done",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            Show-Details -Identity $Script:SelectedIdentity

            if ($Script:CacheLoaded) {
                $updated = Convert-ToGridRow (Get-Mailbox -Identity $Script:SelectedIdentity -ErrorAction Stop)
                for ($i=0; $i -lt $Script:MailboxCache.Count; $i++) {
                    if ($Script:MailboxCache[$i].PrimarySmtpAddress -eq $Script:SelectedIdentity) {
                        $Script:MailboxCache[$i] = $updated
                        break
                    }
                }
                Apply-SearchToCache -QueryText $txtSearch.Text
                Bind-GridFromCurrentView
            }
        } catch {
            Write-Log "Revert to on-prem SOA failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Revert failed.`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $form.Add_FormClosing({
        Write-Log "Application closing requested." "INFO"
        try { Disconnect-EXO } catch { }
        Write-Log "Application closed." "INFO"
    })

    # Init
    Set-UiConnectedState -Connected $false
    $Script:GridBinding.DataSource = New-MailboxGridDataTable
    $Script:GridBinding.ResetBindings($true)

    Write-Log "$($Script:ToolName) GUI starting (Application.Run)..." "INFO"
    [System.Windows.Forms.Application]::Run($form)

} catch {
    Write-Log "FATAL: GUI failed to start. Error=$($_.Exception.Message)" "ERROR"
    [System.Windows.Forms.MessageBox]::Show(
        "GUI failed to start.`n$($_.Exception.Message)`n`nCheck log:`n$($Script:LogFile)",
        "$($Script:ToolName) v$($Script:ScriptVersion)",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}
#endregion
