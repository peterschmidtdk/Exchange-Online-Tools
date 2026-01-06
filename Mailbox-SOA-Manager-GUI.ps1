<#
.SYNOPSIS
  Exchange Mailbox SOA Manager (GUI) for Exchange Online - Cloud-managed Exchange attributes (SOA) toggle.

.DESCRIPTION
  GUI tool to view and change mailbox Exchange attribute SOA state via IsExchangeCloudManaged:
    - Enable cloud management     : Set-Mailbox -IsExchangeCloudManaged $true
    - Revert to on-prem management: Set-Mailbox -IsExchangeCloudManaged $false

  Browse/Search:
    - Connect/Disconnect to Exchange Online
    - Load all mailboxes into local cache (required for fast browsing/search)
    - Paging (Prev/Next + Page size)
    - Search uses cached list
    - Filter by SOA Status: All / Online / On-Prem

  Grid Columns:
    - DisplayName
    - PrimarySMTP
    - SOA Status (Online / On-Prem / Unknown)
    - DirSynced

  UI Enhancement:
    - Rows with SOA Status = Online are highlighted light green (when not selected)

REFERENCE
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

LOGGING
  - Single logfile only (append; never overwritten)
  - Timestamp on every line
  - SOA changes logged with BEFORE/AFTER + Actor + Tenant
  - RunId included for correlation

REQUIREMENTS
  - PowerShell 7+ (Windows) started with -STA (script can auto-relaunch in STA)
  - Module: ExchangeOnlineManagement

AUTHOR
  Peter Schmidt (www.msdigest.net)

VERSION
  2.5.9 (2026-01-06)
    - Add SOA filter dropdown: All / Online / On-Prem (default All)
    - Highlight rows with SOAStatus=Online in light green
#>

#region PS7 Requirement
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This tool requires PowerShell 7+. Start with: pwsh.exe -STA -File .\MailboxSOAManager-GUI.ps1"
    return
}
#endregion

#region Load WinForms
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    [System.Windows.Forms.Application]::EnableVisualStyles()
} catch {
    Write-Error "Failed to load required assemblies. Error: $($_.Exception.Message)"
    return
}
#endregion

#region Globals
$Script:ToolName      = "Exchange Mailbox SOA Manager"
$Script:ScriptVersion = "2.5.9"
$Script:RunId         = [Guid]::NewGuid().ToString()

$Script:LogDir   = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:LogFile  = Join-Path -Path $Script:LogDir -ChildPath "MailboxSOAManager.log"
New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null

$Script:IsConnected = $false
$Script:ExoActor    = $null
$Script:TenantName  = $null

$Script:MailboxCache     = @()  # full cached list (MailboxGridRow)
$Script:CurrentView      = @()  # current filtered list (MailboxGridRow)
$Script:CacheLoaded      = $false
$Script:PageSize         = 50
$Script:PageIndex        = 0
$Script:CurrentQueryText = ""
$Script:CurrentFilter    = "All"  # All | Online | On-Prem
$Script:SelectedIdentity = $null
#endregion

#region Types (Strong binding for WinForms)
class MailboxGridRow {
    [string]$DisplayName
    [string]$PrimarySMTP
    [string]$SOAStatus
    [string]$DirSynced

    MailboxGridRow([string]$displayName, [string]$primarySmtp, [string]$soaStatus, [string]$dirSynced) {
        $this.DisplayName = $displayName
        $this.PrimarySMTP = $primarySmtp
        $this.SOAStatus   = $soaStatus
        $this.DirSynced   = $dirSynced
    }
}
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

function Write-LogException {
    param(
        [Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$Context = "Exception"
    )
    Write-Log "${Context}: $($ErrorRecord.Exception.Message)" "ERROR"
    Write-Log "${Context} (ToString): $($ErrorRecord.Exception.ToString())" "DEBUG"
    if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.PositionMessage) {
        Write-Log "${Context} (Position): $($ErrorRecord.InvocationInfo.PositionMessage)" "DEBUG"
    }
    if ($ErrorRecord.ScriptStackTrace) {
        Write-Log "${Context} (Stack): $($ErrorRecord.ScriptStackTrace)" "DEBUG"
    }
}
#endregion

#region STA guard
function Ensure-STA {
    try {
        $apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
        if ($apt -eq [System.Threading.ApartmentState]::STA) { return $true }

        Write-Log "Not running in STA mode (ApartmentState=$apt). Attempting self-relaunch in STA." "WARN"
        $scriptPath = $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrWhiteSpace($scriptPath) -or -not (Test-Path $scriptPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "This GUI must run in STA mode.`nRun:`n  pwsh.exe -STA -File .\MailboxSOAManager-GUI.ps1",
                $Script:ToolName,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $false
        }

        Start-Process -FilePath "pwsh.exe" -ArgumentList @(
            "-NoProfile","-STA","-ExecutionPolicy","Bypass","-File","`"$scriptPath`""
        ) -WorkingDirectory (Split-Path -Parent $scriptPath) | Out-Null

        return $false
    } catch {
        Write-Log "Ensure-STA failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}
if (-not (Ensure-STA)) { return }
#endregion

Write-Log "$($Script:ToolName) v$($Script:ScriptVersion) starting." "INFO"

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
function Get-SOAStatus {
    param([object]$IsExchangeCloudManaged)
    if ($IsExchangeCloudManaged -eq $true)  { return "Online" }
    if ($IsExchangeCloudManaged -eq $false) { return "On-Prem" }
    return "Unknown"
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
            if ($info -and $info.UserPrincipalName) {
                $u = ([string]$info.UserPrincipalName).Trim().TrimEnd(';')
                return "EXO:$u"
            }
        }
    } catch { }
    return "EXO:unknown"
}

function Get-TenantNameBestEffort {
    try {
        $org = Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1
        if ($org -and $org.Name) {
            return ([string]$org.Name).Trim().TrimEnd(';')
        }
    } catch { }

    try {
        if ($Script:ExoActor -and $Script:ExoActor.StartsWith("EXO:")) {
            $upn = $Script:ExoActor.Replace("EXO:","").Trim().TrimEnd(';')
            if ($upn -like "*@*") { return ($upn.Split("@")[-1]).Trim().TrimEnd(';') }
        }
    } catch { }

    return "Unknown"
}

function Get-AllMailboxesSafe {
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }
    Write-Log "Get-AllMailboxesSafe: Using Get-Mailbox -ResultSize Unlimited" "INFO"

    $raw = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged)

    Write-Log "Get-AllMailboxesSafe: Get-Mailbox returned count=$($raw.Count)" "INFO"
    return $raw
}

function Convert-ToRow {
    param([Parameter(Mandatory)]$MailboxObject)

    $smtp = ""
    if ($MailboxObject.PrimarySmtpAddress) { $smtp = [string]$MailboxObject.PrimarySmtpAddress }

    $dn  = [string]$MailboxObject.DisplayName
    $soa = (Get-SOAStatus $MailboxObject.IsExchangeCloudManaged)
    $ds  = if ($null -eq $MailboxObject.IsDirSynced) { "" } else { [string]$MailboxObject.IsDirSynced }

    return [MailboxGridRow]::new($dn, $smtp, $soa, $ds)
}
#endregion

#region Mailbox ops
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

#region Paging + filter/search
function Reset-ViewToCache {
    $Script:CurrentView = @($Script:MailboxCache)
    $Script:PageIndex = 0
    $Script:CurrentQueryText = ""
    $Script:CurrentFilter = "All"
}

function Apply-SearchAndFilterToCache {
    param(
        [string]$QueryText,
        [string]$Filter
    )

    $q = ""
    if ($null -ne $QueryText) { $q = $QueryText.Trim() }

    $f = "All"
    if ($null -ne $Filter -and -not [string]::IsNullOrWhiteSpace($Filter)) { $f = $Filter.Trim() }

    $Script:CurrentQueryText = $q
    $Script:CurrentFilter    = $f
    $Script:PageIndex        = 0

    $items = @($Script:MailboxCache)

    # Filter first
    switch ($f) {
        "Online"  { $items = @($items | Where-Object { $_.SOAStatus -eq "Online" }) }
        "On-Prem" { $items = @($items | Where-Object { $_.SOAStatus -eq "On-Prem" }) }
        default   { } # All
    }

    # Then search
    if (-not [string]::IsNullOrWhiteSpace($q)) {
        $items = @(
            $items | Where-Object {
                (Text-Matches $_.DisplayName $q) -or
                (Text-Matches $_.PrimarySMTP $q)
            }
        )
    }

    $Script:CurrentView = $items
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
        Write-LogException -ErrorRecord $_ -Context "Connect-EXO failed"
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
            Write-Log "Disconnecting from Exchange Online." "INFO"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Write-Log "Disconnected from Exchange Online." "INFO"
        }
    } catch {
        Write-LogException -ErrorRecord $_ -Context "Disconnect-EXO warning"
    } finally {
        $Script:IsConnected = $false
        $Script:ExoActor = $null
        $Script:TenantName = $null

        $Script:MailboxCache     = @()
        $Script:CurrentView      = @()
        $Script:CacheLoaded      = $false
        $Script:PageIndex        = 0
        $Script:CurrentQueryText = ""
        $Script:CurrentFilter    = "All"
        $Script:SelectedIdentity = $null
    }
}
#endregion

#region GUI
try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($Script:ToolName) v$($Script:ScriptVersion)"
    $form.Size = New-Object System.Drawing.Size(1100, 700)
    $form.StartPosition = "CenterScreen"

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.RowCount = 4
    $root.ColumnCount = 1
    $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52))) | Out-Null
    $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 92))) | Out-Null
    $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32))) | Out-Null
    $form.Controls.Add($root)

    # Top bar
    $top = New-Object System.Windows.Forms.Panel
    $top.Dock = 'Fill'

    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Text = "Connect"
    $btnConnect.Location = New-Object System.Drawing.Point(12, 10)
    $btnConnect.Size = New-Object System.Drawing.Size(110, 30)

    $btnDisconnect = New-Object System.Windows.Forms.Button
    $btnDisconnect.Text = "Disconnect"
    $btnDisconnect.Location = New-Object System.Drawing.Point(130, 10)
    $btnDisconnect.Size = New-Object System.Drawing.Size(110, 30)
    $btnDisconnect.Enabled = $false

    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Text = "Status: Not connected"
    $lblConn.Location = New-Object System.Drawing.Point(260, 16)
    $lblConn.AutoSize = $true

    $btnOpenLog = New-Object System.Windows.Forms.Button
    $btnOpenLog.Text = "View Logfile"
    $btnOpenLog.Location = New-Object System.Drawing.Point(960, 10)
    $btnOpenLog.Size = New-Object System.Drawing.Size(110, 30)

    $top.Controls.AddRange(@($btnConnect,$btnDisconnect,$lblConn,$btnOpenLog))
    $root.Controls.Add($top,0,0)

    # Browse panel
    $browse = New-Object System.Windows.Forms.Panel
    $browse.Dock = 'Fill'

    $btnLoadAll = New-Object System.Windows.Forms.Button
    $btnLoadAll.Text = "Load all mailboxes (cache)"
    $btnLoadAll.Location = New-Object System.Drawing.Point(12, 10)
    $btnLoadAll.Size = New-Object System.Drawing.Size(220, 30)
    $btnLoadAll.Enabled = $false

    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = "◀ Prev"
    $btnPrev.Location = New-Object System.Drawing.Point(260, 10)
    $btnPrev.Size = New-Object System.Drawing.Size(90, 30)
    $btnPrev.Enabled = $false

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = "Next ▶"
    $btnNext.Location = New-Object System.Drawing.Point(356, 10)
    $btnNext.Size = New-Object System.Drawing.Size(90, 30)
    $btnNext.Enabled = $false

    $lblPage = New-Object System.Windows.Forms.Label
    $lblPage.Text = "Page: -"
    $lblPage.Location = New-Object System.Drawing.Point(460, 16)
    $lblPage.AutoSize = $true

    $lblPageSize = New-Object System.Windows.Forms.Label
    $lblPageSize.Text = "Page size:"
    $lblPageSize.Location = New-Object System.Drawing.Point(560, 16)
    $lblPageSize.AutoSize = $true

    $cmbPageSize = New-Object System.Windows.Forms.ComboBox
    $cmbPageSize.Location = New-Object System.Drawing.Point(630, 12)
    $cmbPageSize.Size = New-Object System.Drawing.Size(90, 25)
    $cmbPageSize.DropDownStyle = 'DropDownList'
    [void]$cmbPageSize.Items.AddRange(@("25","50","100","200"))
    $cmbPageSize.SelectedItem = "50"
    $cmbPageSize.Enabled = $false

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "SOA Filter:"
    $lblFilter.Location = New-Object System.Drawing.Point(740, 16)
    $lblFilter.AutoSize = $true

    $cmbFilter = New-Object System.Windows.Forms.ComboBox
    $cmbFilter.Location = New-Object System.Drawing.Point(810, 12)
    $cmbFilter.Size = New-Object System.Drawing.Size(110, 25)
    $cmbFilter.DropDownStyle = 'DropDownList'
    [void]$cmbFilter.Items.AddRange(@("All","Online","On-Prem"))
    $cmbFilter.SelectedItem = "All"
    $cmbFilter.Enabled = $false

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Count: -"
    $lblCount.Location = New-Object System.Drawing.Point(940, 16)
    $lblCount.AutoSize = $true

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(12, 50)
    $txtSearch.Size = New-Object System.Drawing.Size(520, 25)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Search (cache)"
    $btnSearch.Location = New-Object System.Drawing.Point(540, 48)
    $btnSearch.Size = New-Object System.Drawing.Size(120, 30)
    $btnSearch.Enabled = $false

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = "Clear"
    $btnClear.Location = New-Object System.Drawing.Point(668, 48)
    $btnClear.Size = New-Object System.Drawing.Size(90, 30)
    $btnClear.Enabled = $false

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(780, 54)
    $lblStatus.Size = New-Object System.Drawing.Size(290, 20)

    $browse.Controls.AddRange(@(
        $btnLoadAll,$btnPrev,$btnNext,$lblPage,$lblPageSize,$cmbPageSize,
        $lblFilter,$cmbFilter,$lblCount,
        $txtSearch,$btnSearch,$btnClear,$lblStatus
    ))
    $root.Controls.Add($browse,0,1)

    # Grid area
    $gridPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $gridPanel.Dock = 'Fill'
    $gridPanel.RowCount = 2
    $gridPanel.ColumnCount = 1
    $gridPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $gridPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 48))) | Out-Null
    $root.Controls.Add($gridPanel,0,2)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $false
    $grid.AutoGenerateColumns = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.RowHeadersVisible = $false

    # Ensure text is visible in all themes
    $grid.BackgroundColor = [System.Drawing.SystemColors]::Window
    $grid.DefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Window
    $grid.DefaultCellStyle.ForeColor = [System.Drawing.SystemColors]::WindowText
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText

    $gridPanel.Controls.Add($grid,0,0)

    # Manual columns (with Name for Cells["..."] access)
    $grid.Columns.Clear() | Out-Null

    $col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col1.Name = "DisplayName"
    $col1.HeaderText = "DisplayName"
    $col1.DataPropertyName = "DisplayName"
    $grid.Columns.Add($col1) | Out-Null

    $col2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col2.Name = "PrimarySMTP"
    $col2.HeaderText = "PrimarySMTP"
    $col2.DataPropertyName = "PrimarySMTP"
    $grid.Columns.Add($col2) | Out-Null

    $col3 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col3.Name = "SOAStatus"
    $col3.HeaderText = "SOA Status"
    $col3.DataPropertyName = "SOAStatus"
    $grid.Columns.Add($col3) | Out-Null

    $col4 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col4.Name = "DirSynced"
    $col4.HeaderText = "DirSynced"
    $col4.DataPropertyName = "DirSynced"
    $grid.Columns.Add($col4) | Out-Null

    # Row highlight (Online = light green)
    $OnlineRowColor = [System.Drawing.Color]::PaleGreen
    $DefaultRowColor = [System.Drawing.SystemColors]::Window

    $grid.Add_RowPrePaint({
        param($sender,$e)
        try {
            $row = $sender.Rows[$e.RowIndex]
            if ($null -eq $row -or $row.IsNewRow) { return }

            # Keep default selection styling when selected
            if ($row.Selected) { return }

            $soa = $row.Cells["SOAStatus"].Value
            if ($soa -eq "Online") {
                $row.DefaultCellStyle.BackColor = $OnlineRowColor
            } else {
                $row.DefaultCellStyle.BackColor = $DefaultRowColor
            }
        } catch { }
    })

    $actions = New-Object System.Windows.Forms.Panel
    $actions.Dock = 'Fill'

    $btnEnableCloud = New-Object System.Windows.Forms.Button
    $btnEnableCloud.Text = "Enable SOA = Online"
    $btnEnableCloud.Location = New-Object System.Drawing.Point(12, 8)
    $btnEnableCloud.Size = New-Object System.Drawing.Size(170, 32)
    $btnEnableCloud.Enabled = $false

    $btnRevertOnPrem = New-Object System.Windows.Forms.Button
    $btnRevertOnPrem.Text = "Revert SOA = On-Prem"
    $btnRevertOnPrem.Location = New-Object System.Drawing.Point(190, 8)
    $btnRevertOnPrem.Size = New-Object System.Drawing.Size(190, 32)
    $btnRevertOnPrem.Enabled = $false

    $btnRefreshRow = New-Object System.Windows.Forms.Button
    $btnRefreshRow.Text = "Refresh selected"
    $btnRefreshRow.Location = New-Object System.Drawing.Point(388, 8)
    $btnRefreshRow.Size = New-Object System.Drawing.Size(140, 32)
    $btnRefreshRow.Enabled = $false

    $actions.Controls.AddRange(@($btnEnableCloud,$btnRevertOnPrem,$btnRefreshRow))
    $gridPanel.Controls.Add($actions,0,1)

    # Footer
    $footer = New-Object System.Windows.Forms.Label
    $footer.Dock = 'Fill'
    $footer.TextAlign = 'MiddleLeft'
    $footer.Padding = New-Object System.Windows.Forms.Padding(10,0,0,0)
    $footer.Text = "v$($Script:ScriptVersion) | Toggle mailbox SOA (IsExchangeCloudManaged) | Log: $($Script:LogFile)"
    $root.Controls.Add($footer,0,3)

    # Strongly-typed BindingList for page view
    $PageBinding = [System.ComponentModel.BindingList[MailboxGridRow]]::new()
    $grid.DataSource = $PageBinding

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
        $pageItems = Get-PageSlice -Items $Script:CurrentView -PageIndex $Script:PageIndex -PageSize $Script:PageSize

        $PageBinding.RaiseListChangedEvents = $false
        $PageBinding.Clear()
        foreach ($x in @($pageItems)) { [void]$PageBinding.Add($x) }
        $PageBinding.RaiseListChangedEvents = $true
        $PageBinding.ResetBindings()

        $grid.Refresh()
        Update-PagingUI

        Write-Log "Grid bind diagnostics: Filter='$($Script:CurrentFilter)' Query='$($Script:CurrentQueryText)' PageItems=$($pageItems.Count) BindingCount=$($PageBinding.Count) GridRows=$($grid.Rows.Count)" "INFO"
    }

    function Reset-Selection {
        $Script:SelectedIdentity = $null
        $btnEnableCloud.Enabled = $false
        $btnRevertOnPrem.Enabled = $false
        $btnRefreshRow.Enabled = $false
    }

    function Set-UiConnectedState {
        param([bool]$Connected)

        $btnConnect.Enabled      = -not $Connected
        $btnDisconnect.Enabled   = $Connected
        $btnLoadAll.Enabled      = $Connected
        $btnSearch.Enabled       = $Connected
        $cmbPageSize.Enabled     = $Connected
        $cmbFilter.Enabled       = $Connected

        if (-not $Connected) {
            $lblConn.Text = "Status: Not connected"
            $PageBinding.Clear()
            $lblStatus.Text = ""
            $lblPage.Text = "Page: -"
            $lblCount.Text = "Count: -"
            $btnPrev.Enabled = $false
            $btnNext.Enabled = $false
            $btnClear.Enabled = $false
            $cmbFilter.SelectedItem = "All"
            Reset-Selection
        } else {
            $tn = if ($Script:TenantName) { $Script:TenantName } else { "Unknown" }
            $lblConn.Text = "Status: Connected to Exchange Online (Tenant: $tn)"
            $lblStatus.Text = "Connected"
        }
    }

    function Rebuild-ViewAndBind {
        if (-not $Script:CacheLoaded) { return }
        $filter = if ($cmbFilter.SelectedItem) { [string]$cmbFilter.SelectedItem } else { "All" }
        Apply-SearchAndFilterToCache -QueryText $txtSearch.Text -Filter $filter
        Bind-GridFromCurrentView
        Reset-Selection
    }

    # Events
    $btnOpenLog.Add_Click({
        try {
            if (-not (Test-Path $Script:LogFile)) { New-Item -Path $Script:LogFile -ItemType File -Force | Out-Null }
            Start-Process -FilePath $Script:LogFile | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to open log file.`n$($_.Exception.Message)`n`nPath:`n$($Script:LogFile)",
                "Open log failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })

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
            if ($Script:CacheLoaded) {
                Bind-GridFromCurrentView
                Reset-Selection
            }
        } catch { }
    })

    $cmbFilter.Add_SelectedIndexChanged({
        try {
            $sel = if ($cmbFilter.SelectedItem) { [string]$cmbFilter.SelectedItem } else { "All" }
            Write-Log "SOA Filter changed to '$sel'" "INFO"
            Rebuild-ViewAndBind
        } catch { }
    })

    $btnPrev.Add_Click({
        if ($Script:PageIndex -gt 0) {
            $Script:PageIndex--
            Write-Log "Paging Prev. PageIndex=$($Script:PageIndex)" "INFO"
            Bind-GridFromCurrentView
            Reset-Selection
        }
    })

    $btnNext.Add_Click({
        $totalPages = Get-TotalPages -Items $Script:CurrentView -PageSize $Script:PageSize
        if ($Script:PageIndex -lt ($totalPages - 1)) {
            $Script:PageIndex++
            Write-Log "Paging Next. PageIndex=$($Script:PageIndex)" "INFO"
            Bind-GridFromCurrentView
            Reset-Selection
        }
    })

    $btnLoadAll.Add_Click({
        if (-not $Script:IsConnected) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $lblStatus.Text = "Loading..."
            [System.Windows.Forms.Application]::DoEvents()

            Write-Log "LoadAll clicked." "INFO"

            $raw = Get-AllMailboxesSafe
            $cache = foreach ($m in $raw) { Convert-ToRow $m }

            $Script:MailboxCache = @($cache)
            $Script:CacheLoaded  = $true

            # Apply default filter/search (All, no query)
            $cmbFilter.SelectedItem = "All"
            Apply-SearchAndFilterToCache -QueryText "" -Filter "All"
            Bind-GridFromCurrentView
            Reset-Selection

            $btnClear.Enabled = $true
            $lblStatus.Text = "Loaded"
            $lblCount.Text = "Count: $($Script:MailboxCache.Count)"

            Write-Log "LoadAll success. CachedCount=$($Script:MailboxCache.Count)" "INFO"
        } catch {
            Write-LogException -ErrorRecord $_ -Context "LoadAll failed"
            $lblStatus.Text = "Load failed"
            [System.Windows.Forms.MessageBox]::Show(
                "Load all mailboxes failed.`n`n$($_.Exception.Message)`n`nLog:`n$($Script:LogFile)",
                "Load all failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $btnSearch.Add_Click({
        if (-not $Script:IsConnected) { return }

        $qTrim = if ($txtSearch.Text) { $txtSearch.Text.Trim() } else { "" }
        $fSel  = if ($cmbFilter.SelectedItem) { [string]$cmbFilter.SelectedItem } else { "All" }
        Write-Log "Search clicked. Query='$qTrim' Filter='$fSel' CacheLoaded=$($Script:CacheLoaded)" "INFO"

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if (-not $Script:CacheLoaded) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please click 'Load all mailboxes (cache)' first.",
                    "Search",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
                return
            }

            Apply-SearchAndFilterToCache -QueryText $qTrim -Filter $fSel
            Bind-GridFromCurrentView
            Reset-Selection

            $matches = if ($Script:CurrentView) { $Script:CurrentView.Count } else { 0 }
            $lblStatus.Text = "Matches: $matches"
            $btnClear.Enabled = $true
        } catch {
            Write-LogException -ErrorRecord $_ -Context "Search failed"
            [System.Windows.Forms.MessageBox]::Show(
                "Search failed.`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnClear.Add_Click({
        $txtSearch.Text = ""
        $cmbFilter.SelectedItem = "All"
        if ($Script:CacheLoaded) {
            Apply-SearchAndFilterToCache -QueryText "" -Filter "All"
            Bind-GridFromCurrentView
            Reset-Selection
            $lblStatus.Text = "Showing all"
        }
    })

    $grid.Add_SelectionChanged({
        try {
            if ($grid.SelectedRows.Count -gt 0) {
                $smtp = $grid.SelectedRows[0].Cells["PrimarySMTP"].Value
                if ($smtp) {
                    $Script:SelectedIdentity = $smtp.ToString()
                    $btnEnableCloud.Enabled = $true
                    $btnRevertOnPrem.Enabled = $true
                    $btnRefreshRow.Enabled = $true
                }
            }
        } catch {
            Write-LogException -ErrorRecord $_ -Context "SelectionChanged warning"
        }
    })

    $btnRefreshRow.Add_Click({
        if (-not $Script:SelectedIdentity) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $mbx = Get-Mailbox -Identity $Script:SelectedIdentity -ErrorAction Stop |
                Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

            $updated = Convert-ToRow $mbx
            for ($i=0; $i -lt $Script:MailboxCache.Count; $i++) {
                if ($Script:MailboxCache[$i].PrimarySMTP -eq $Script:SelectedIdentity) {
                    $Script:MailboxCache[$i] = $updated
                    break
                }
            }

            Rebuild-ViewAndBind
            Write-Log "Refresh selected completed for '$($Script:SelectedIdentity)'" "INFO"
        } catch {
            Write-LogException -ErrorRecord $_ -Context "Refresh selected failed"
            [System.Windows.Forms.MessageBox]::Show(
                "Refresh failed.`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnEnableCloud.Add_Click({
        if (-not $Script:SelectedIdentity) { return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Enable SOA = Online for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = TRUE.",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $true
            [System.Windows.Forms.MessageBox]::Show($msg,"Done",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $btnRefreshRow.PerformClick()
        } catch {
            Write-LogException -ErrorRecord $_ -Context "Enable cloud SOA failed"
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
            "Revert SOA = On-Prem for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = FALSE.`n`nWARNING: Next sync may overwrite cloud values with on-prem values.",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $false
            [System.Windows.Forms.MessageBox]::Show($msg,"Done",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $btnRefreshRow.PerformClick()
        } catch {
            Write-LogException -ErrorRecord $_ -Context "Revert to on-prem SOA failed"
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

    Set-UiConnectedState -Connected $false

    Write-Log "GUI starting (Application.Run)..." "INFO"
    [System.Windows.Forms.Application]::Run($form)

} catch {
    Write-LogException -ErrorRecord $_ -Context "FATAL: GUI failed to start"
    [System.Windows.Forms.MessageBox]::Show(
        "GUI failed to start.`n$($_.Exception.Message)`n`nCheck log:`n$($Script:LogFile)",
        "$($Script:ToolName) v$($Script:ScriptVersion)",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}
#endregion
