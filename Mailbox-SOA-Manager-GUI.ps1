<#
.SYNOPSIS
  Mailbox SOA Manager (GUI) for Exchange Online - Cloud-managed Exchange attributes (SOA) toggle.

.DESCRIPTION
  GUI tool to view and change mailbox Exchange attribute SOA state via IsExchangeCloudManaged:
    - Enable cloud management  : IsExchangeCloudManaged = $true
    - Revert to on-prem management: IsExchangeCloudManaged = $false
  Includes SOA indicator column in the list and detailed change logging.

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
  Peter

VERSION
  1.4 (2026-01-04)
#>

# --- Load WinForms early (so we can show MessageBoxes even before GUI starts) ---
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    [System.Windows.Forms.Application]::EnableVisualStyles()
} catch {
    Write-Error "Failed to load WinForms assemblies. This tool must run on Windows with WinForms available. Error: $($_.Exception.Message)"
    return
}

#region Globals
$Script:ToolName     = "Mailbox SOA Manager"
$Script:RunId        = [Guid]::NewGuid().ToString()
$Script:LogDir       = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:ExportDir    = Join-Path -Path (Get-Location) -ChildPath "Exports"
$Script:LogFile      = Join-Path -Path $Script:LogDir -ChildPath "MailboxSOAManager.log"
$Script:IsConnected  = $false
$Script:ExoActor     = $null  # populated after connect (best-effort)

New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
New-Item -ItemType Directory -Path $Script:ExportDir -Force | Out-Null
#endregion

#region Logging (single logfile, timestamp per line)
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $winUser = "$env:USERDOMAIN\$env:USERNAME"
    $actor = if ($Script:ExoActor) { $Script:ExoActor } else { "EXO:unknown" }
    $line = "[$ts][$Level][RunId:$($Script:RunId)][Win:$winUser][$actor] $Message"
    Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8
}
#endregion

#region STA guard (AUTO RELAUNCH)
function Ensure-STA {
    try {
        $apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
        if ($apt -eq [System.Threading.ApartmentState]::STA) {
            return $true
        }

        Write-Log "Not running in STA mode (ApartmentState=$apt). Attempting self-relaunch in STA..." "WARN"

        $scriptPath = $MyInvocation.MyCommand.Path
        if ([string]::IsNullOrWhiteSpace($scriptPath) -or -not (Test-Path $scriptPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "This GUI must run in STA mode, but the script path could not be detected for auto-relaunch.`n`nPlease run it like:`n  powershell.exe -STA -File .\MailboxSOAManager-GUI.ps1",
                "Mailbox SOA Manager",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $false
        }

        $exe = if ($PSVersionTable.PSEdition -eq "Core") { "pwsh.exe" } else { "powershell.exe" }

        $args = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-STA",
            "-File", "`"$scriptPath`""
        ) -join " "

        Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory (Split-Path -Parent $scriptPath) | Out-Null

        Write-Log "Launched new process: $exe $args" "INFO"
        return $false
    } catch {
        Write-Log "Ensure-STA failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to validate STA mode.`n`n$($_.Exception.Message)",
            "Mailbox SOA Manager",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
}

if (-not (Ensure-STA)) { return }
#endregion

Write-Log "$($Script:ToolName) starting (GUI init)..." "INFO"

#region Module helpers
function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Required module '$Name' is not installed.`n`nInstall it now (CurrentUser)?",
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

#region SOA indicator helper
function Get-SOAIndicator {
    param([object]$IsExchangeCloudManaged)
    if ($IsExchangeCloudManaged -eq $true)  { return "☁ Online" }
    if ($IsExchangeCloudManaged -eq $false) { return "🏢 On-Prem" }
    return "? Unknown"
}
#endregion

#region EXO connect/disconnect
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

function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"
        Write-Log "Connecting to Exchange Online..." "INFO"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
        $Script:IsConnected = $true
        $Script:ExoActor = Get-ExoActorBestEffort
        Write-Log "Connected to Exchange Online." "INFO"
        return $true
    } catch {
        $Script:IsConnected = $false
        $Script:ExoActor = "EXO:unknown"
        Write-Log "Connect-EXO failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange Online.`n`n$($_.Exception.Message)",
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
    }
}
#endregion

#region Mailbox ops
function Search-Mailboxes {
    param(
        [Parameter(Mandatory)][string]$QueryText,
        [int]$Max = 200
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $q = $QueryText.Trim()
    if ([string]::IsNullOrWhiteSpace($q)) { return @() }

    $filter = "DisplayName -like '*$q*' -or Alias -like '*$q*' -or PrimarySmtpAddress -like '*$q*'"
    Write-Log "Search-Mailboxes started. Query='$q' Max=$Max Filter='$filter'" "INFO"

    $items = Get-Mailbox -ResultSize $Max -Filter $filter -ErrorAction Stop |
        Select-Object `
            DisplayName,
            Alias,
            PrimarySmtpAddress,
            RecipientTypeDetails,
            IsDirSynced,
            IsExchangeCloudManaged,
            @{Name="SOA (Exchange Attributes)"; Expression={ Get-SOAIndicator $_.IsExchangeCloudManaged }}

    Write-Log "Search-Mailboxes completed. Results=$($items.Count)" "INFO"
    return @($items)
}

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

    [PSCustomObject]@{
        Mailbox = $mbx
        User    = $usr
    }
}

function Export-MailboxBackup {
    param([Parameter(Mandatory)][string]$Identity)

    $details = Get-MailboxDetails -Identity $Identity
    $safeId  = ($details.Mailbox.PrimarySmtpAddress.ToString() -replace '[^a-zA-Z0-9\.\-_@]','_')
    $stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
    $path    = Join-Path $Script:ExportDir "$safeId-MailboxSOA-Backup-$stamp.json"

    $details | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8
    Write-Log "Export-MailboxBackup completed for '$Identity'. Path='$path'" "INFO"
    return $path
}

function Set-MailboxSOACloudManaged {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][bool]$EnableCloudManaged
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $targetValue = [bool]$EnableCloudManaged

    # BEFORE
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

    # APPLY
    try {
        Set-Mailbox -Identity $Identity -IsExchangeCloudManaged $targetValue -ErrorAction Stop
        Write-Log "Set-Mailbox executed for '$Identity' IsExchangeCloudManaged=$targetValue" "INFO"
    } catch {
        Write-Log "Set-Mailbox FAILED for '$Identity'. Error=$($_.Exception.Message)" "ERROR"
        throw
    }

    # AFTER
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

#region GUI (WinForms)
try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$($Script:ToolName) - Exchange Online (IsExchangeCloudManaged)"
    $form.Size = New-Object System.Drawing.Size(980, 640)
    $form.StartPosition = "CenterScreen"

    # Top bar
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

    # Search
    $grpSearch = New-Object System.Windows.Forms.GroupBox
    $grpSearch.Text = "Search mailboxes (DisplayName, Alias, Primary SMTP)"
    $grpSearch.Location = New-Object System.Drawing.Point(12, 55)
    $grpSearch.Size = New-Object System.Drawing.Size(940, 120)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(16, 30)
    $txtSearch.Size = New-Object System.Drawing.Size(720, 25)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Search"
    $btnSearch.Location = New-Object System.Drawing.Point(750, 27)
    $btnSearch.Size = New-Object System.Drawing.Size(170, 30)
    $btnSearch.Enabled = $false

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(16, 65)
    $grid.Size = New-Object System.Drawing.Size(904, 45)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.AutoGenerateColumns = $true

    # Details
    $grpDetails = New-Object System.Windows.Forms.GroupBox
    $grpDetails.Text = "Selected mailbox details"
    $grpDetails.Location = New-Object System.Drawing.Point(12, 185)
    $grpDetails.Size = New-Object System.Drawing.Size(940, 260)

    $txtDetails = New-Object System.Windows.Forms.TextBox
    $txtDetails.Location = New-Object System.Drawing.Point(16, 30)
    $txtDetails.Size = New-Object System.Drawing.Size(904, 165)
    $txtDetails.Multiline = $true
    $txtDetails.ScrollBars = "Vertical"
    $txtDetails.ReadOnly = $true

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh details"
    $btnRefresh.Location = New-Object System.Drawing.Point(16, 205)
    $btnRefresh.Size = New-Object System.Drawing.Size(170, 32)
    $btnRefresh.Enabled = $false

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Text = "Export backup (JSON)"
    $btnBackup.Location = New-Object System.Drawing.Point(196, 205)
    $btnBackup.Size = New-Object System.Drawing.Size(190, 32)
    $btnBackup.Enabled = $false

    $btnEnableCloud = New-Object System.Windows.Forms.Button
    $btnEnableCloud.Text = "Enable cloud SOA (true)"
    $btnEnableCloud.Location = New-Object System.Drawing.Point(396, 205)
    $btnEnableCloud.Size = New-Object System.Drawing.Size(210, 32)
    $btnEnableCloud.Enabled = $false

    $btnRevertOnPrem = New-Object System.Windows.Forms.Button
    $btnRevertOnPrem.Text = "Revert to on-prem SOA (false)"
    $btnRevertOnPrem.Location = New-Object System.Drawing.Point(616, 205)
    $btnRevertOnPrem.Size = New-Object System.Drawing.Size(250, 32)
    $btnRevertOnPrem.Enabled = $false

    # Footer
    $lblFoot = New-Object System.Windows.Forms.Label
    $lblFoot.Location = New-Object System.Drawing.Point(12, 460)
    $lblFoot.Size = New-Object System.Drawing.Size(940, 130)
    $lblFoot.Text =
"Notes:
- Single logfile (append): $($Script:LogFile)
- Each log line is timestamped and includes RunId + Actor.
- 'SOA (Exchange Attributes)' indicator is based on IsExchangeCloudManaged:
    ☁ Online  = Exchange Online is SOA for Exchange attributes
    🏢 On-Prem = On-premises is SOA for Exchange attributes
    ? Unknown = Not set/unknown
Exports: $($Script:ExportDir)
"
    $lblFoot.AutoSize = $false

    # Add controls
    $form.Controls.AddRange(@($btnConnect,$btnDisconnect,$lblConn,$grpSearch,$grpDetails,$lblFoot))
    $grpSearch.Controls.AddRange(@($txtSearch,$btnSearch,$grid))
    $grpDetails.Controls.AddRange(@($txtDetails,$btnRefresh,$btnBackup,$btnEnableCloud,$btnRevertOnPrem))

    # State
    $Script:SelectedIdentity = $null

    function Set-UiConnectedState {
        param([bool]$Connected)

        $btnConnect.Enabled        = -not $Connected
        $btnDisconnect.Enabled     = $Connected
        $btnSearch.Enabled         = $Connected

        $btnRefresh.Enabled        = $false
        $btnBackup.Enabled         = $false
        $btnEnableCloud.Enabled    = $false
        $btnRevertOnPrem.Enabled   = $false

        if ($Connected) {
            $lblConn.Text = "Status: Connected to Exchange Online"
        } else {
            $lblConn.Text = "Status: Not connected"
            $grid.DataSource = $null
            $txtDetails.Clear()
            $Script:SelectedIdentity = $null
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
        $btnBackup.Enabled       = $true
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

    $btnSearch.Add_Click({
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $q = $txtSearch.Text
            Write-Log "UI Search clicked. Query='$q'" "INFO"
            $results = Search-Mailboxes -QueryText $q -Max 200

            if ($results.Count -eq 0) {
                $grid.DataSource = $null
                $txtDetails.Text = "No results."
                $Script:SelectedIdentity = $null
                $btnRefresh.Enabled = $false
                $btnBackup.Enabled  = $false
                $btnEnableCloud.Enabled = $false
                $btnRevertOnPrem.Enabled= $false
                return
            }

            $grid.DataSource = $results
            $txtDetails.Text = "Select a mailbox row to see details."
        } catch {
            Write-Log "UI Search failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Search failed.`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $grid.Add_SelectionChanged({
        try {
            if ($grid.SelectedRows.Count -gt 0) {
                $row = $grid.SelectedRows[0]
                $smtp = $row.Cells["PrimarySmtpAddress"].Value
                if ($smtp) {
                    $Script:SelectedIdentity = $smtp.ToString()
                    Write-Log "UI selection changed. SelectedIdentity='$($Script:SelectedIdentity)'" "INFO"
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
        try {
            Write-Log "UI Refresh clicked. Identity='$($Script:SelectedIdentity)'" "INFO"
            Show-Details -Identity $Script:SelectedIdentity
        } catch {
            Write-Log "Refresh failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Refresh failed.`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnBackup.Add_Click({
        if (-not $Script:SelectedIdentity) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-Log "UI Backup clicked. Identity='$($Script:SelectedIdentity)'" "INFO"
            $path = Export-MailboxBackup -Identity $Script:SelectedIdentity
            [System.Windows.Forms.MessageBox]::Show(
                "Backup exported:`n$path",
                "Export Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } catch {
            Write-Log "Backup export failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Backup export failed.`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnEnableCloud.Add_Click({
        if (-not $Script:SelectedIdentity) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Enable cloud SOA for Exchange attributes for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = TRUE.",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Enable cloud SOA cancelled by user. Identity='$($Script:SelectedIdentity)'" "INFO"
            return
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-Log "Enable cloud SOA initiated. Identity='$($Script:SelectedIdentity)'" "INFO"
            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $true
            [System.Windows.Forms.MessageBox]::Show(
                $msg,
                "Done",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            Show-Details -Identity $Script:SelectedIdentity
            $btnSearch.PerformClick()
        } catch {
            Write-Log "Enable cloud SOA failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Enable cloud SOA failed.`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnRevertOnPrem.Add_Click({
        if (-not $Script:SelectedIdentity) { return }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Revert SOA back to on-prem for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = FALSE.`n`nWARNING: Next sync may overwrite cloud values with on-prem values.",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Revert to on-prem SOA cancelled by user. Identity='$($Script:SelectedIdentity)'" "INFO"
            return
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Write-Log "Revert to on-prem SOA initiated. Identity='$($Script:SelectedIdentity)'" "INFO"

            $backupPrompt = [System.Windows.Forms.MessageBox]::Show(
                "Do you want to export a backup (JSON) before reverting?",
                "Backup recommended",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($backupPrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
                $path = Export-MailboxBackup -Identity $Script:SelectedIdentity
                Write-Log "Backup created before revert. Identity='$($Script:SelectedIdentity)' Path='$path'" "INFO"
            } else {
                Write-Log "Backup skipped before revert. Identity='$($Script:SelectedIdentity)'" "WARN"
            }

            $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $false
            [System.Windows.Forms.MessageBox]::Show(
                $msg,
                "Done",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            Show-Details -Identity $Script:SelectedIdentity
            $btnSearch.PerformClick()
        } catch {
            Write-Log "Revert to on-prem SOA failed: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Revert failed.`n`n$($_.Exception.Message)",
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

    # Init + Run
    Set-UiConnectedState -Connected $false
    Write-Log "$($Script:ToolName) GUI starting (Application.Run)..." "INFO"
    [System.Windows.Forms.Application]::Run($form)

} catch {
    Write-Log "FATAL: GUI failed to start. Error=$($_.Exception.Message)" "ERROR"
    [System.Windows.Forms.MessageBox]::Show(
        "GUI failed to start.`n`n$($_.Exception.Message)`n`nCheck log:`n$($Script:LogFile)",
        "Mailbox SOA Manager",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}
#endregion
