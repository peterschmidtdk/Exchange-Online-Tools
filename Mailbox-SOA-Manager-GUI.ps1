<#
.SYNOPSIS
  Mailbox SOA Manager Tool (GUI) for Exchange Online - Cloud-managed Exchange attributes (SOA) toggle.

.DESCRIPTION
  In Exchange Hybrid environments, Microsoft introduced a per-mailbox switch to transfer the
  Source of Authority (SOA) for Exchange attributes from on-premises to Exchange Online.

  This tool provides a Windows GUI to:
    - Search for EXO mailboxes
    - View IsDirSynced + IsExchangeCloudManaged state
    - Enable cloud management (IsExchangeCloudManaged = $true)
    - Revert to on-premises management (IsExchangeCloudManaged = $false)
    - Export a small backup (JSON) of mailbox + user properties before reverting

  Microsoft reference:
    - Set-Mailbox -Identity <User> -IsExchangeCloudManaged $true/$false
    - Requires appropriate admin roles (Exchange Admin recommended)
    - For Entra Connect Sync tenants, Microsoft recommends Entra Connect Sync 2.5.190.0+,
      and waiting after on-prem changes before switching a mailbox to cloud-managed. :contentReference[oaicite:1]{index=1}

IMPORTANT NOTES
  - This does NOT migrate mailboxes. It only changes where Exchange *attributes* are managed.
  - After you set IsExchangeCloudManaged=$false, next sync cycle will overwrite cloud values
    with on-prem values (per Microsoft). Backup/export any needed cloud-only changes first. :contentReference[oaicite:2]{index=2}

REQUIREMENTS
  - Windows PowerShell 5.1 OR PowerShell 7+ started with -STA
  - Module: ExchangeOnlineManagement

AUTHOR
  Peter

VERSION
  1.0 (2026-01-02)
#>

#region Safety / STA check
try {
    $apt = [System.Threading.Thread]::CurrentThread.GetApartmentState()
    if ($apt -ne [System.Threading.ApartmentState]::STA) {
        Write-Warning "This GUI must run in STA mode. Start PowerShell with -STA and re-run."
        Write-Warning "Examples:"
        Write-Warning "  Windows PowerShell: powershell.exe -STA -File .\SOA-MailboxTool-GUI.ps1"
        Write-Warning "  PowerShell 7+:      pwsh.exe -STA -File .\SOA-MailboxTool-GUI.ps1"
        return
    }
} catch {
    # If we can't detect, continue (best effort)
}
#endregion

#region Globals
$Script:LogDir     = Join-Path -Path (Get-Location) -ChildPath "Logs"
$Script:ExportDir  = Join-Path -Path (Get-Location) -ChildPath "Exports"
$Script:LogFile    = Join-Path -Path $Script:LogDir -ChildPath "SOA-MailboxTool.log"
$Script:IsConnected = $false

New-Item -ItemType Directory -Path $Script:LogDir -Force | Out-Null
New-Item -ItemType Directory -Path $Script:ExportDir -Force | Out-Null
#endregion

#region Logging
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Message"
    Add-Content -Path $Script:LogFile -Value $line
}
#endregion

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

        Write-Log "Installing module $Name for CurrentUser..." "INFO"
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            Write-Log "Install-Module failed: $($_.Exception.Message)" "ERROR"
            throw
        }
    }

    Import-Module $Name -ErrorAction Stop
    Write-Log "Module loaded: $Name" "INFO"
}
#endregion

#region EXO connect/disconnect
function Connect-EXO {
    try {
        Ensure-Module -Name "ExchangeOnlineManagement"

        Write-Log "Connecting to Exchange Online..." "INFO"
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
        $Script:IsConnected = $true
        Write-Log "Connected to Exchange Online." "INFO"
        return $true
    } catch {
        $Script:IsConnected = $false
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
        }
    } catch {
        Write-Log "Disconnect-EXO warning: $($_.Exception.Message)" "WARN"
    } finally {
        $Script:IsConnected = $false
        Write-Log "Disconnected (or no session)." "INFO"
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

    # Use OPATH filter for performance (best effort)
    $filter = "DisplayName -like '*$q*' -or Alias -like '*$q*' -or PrimarySmtpAddress -like '*$q*'"

    $items = Get-Mailbox -ResultSize $Max -Filter $filter -ErrorAction Stop |
        Select-Object DisplayName,Alias,PrimarySmtpAddress,RecipientTypeDetails,IsDirSynced,IsExchangeCloudManaged

    return @($items)
}

function Get-MailboxDetails {
    param([Parameter(Mandatory)][string]$Identity)

    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,Alias,PrimarySmtpAddress,RecipientTypeDetails,IsDirSynced,IsExchangeCloudManaged,ExchangeGuid,ExternalDirectoryObjectId

    # Get-User gives some additional attributes that admins often want to backup/see
    $usr = $null
    try {
        $usr = Get-User -Identity $Identity -ErrorAction Stop |
            Select-Object DisplayName,UserPrincipalName,ImmutableId,RecipientTypeDetails,WhenChangedUTC
    } catch {
        # Not fatal
    }

    [PSCustomObject]@{
        Mailbox = $mbx
        User    = $usr
    }
}

function Export-MailboxBackup {
    param(
        [Parameter(Mandatory)][string]$Identity
    )

    $details = Get-MailboxDetails -Identity $Identity
    $safeId  = ($details.Mailbox.PrimarySmtpAddress.ToString() -replace '[^a-zA-Z0-9\.\-_@]','_')
    $stamp   = Get-Date -Format "yyyyMMdd-HHmmss"
    $path    = Join-Path $Script:ExportDir "$safeId-SOA-Backup-$stamp.json"

    $details | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8
    Write-Log "Exported backup for $Identity to $path" "INFO"
    return $path
}

function Set-MailboxSOACloudManaged {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][bool]$EnableCloudManaged
    )
    if (-not $Script:IsConnected) { throw "Not connected to Exchange Online." }

    $targetValue = [bool]$EnableCloudManaged

    # Pre-check
    $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop |
        Select-Object DisplayName,PrimarySmtpAddress,IsDirSynced,IsExchangeCloudManaged

    if ($mbx.IsDirSynced -ne $true) {
        throw "Mailbox '$Identity' is not DirSynced (IsDirSynced=$($mbx.IsDirSynced)). This switch is intended for directory-synchronized users."
    }

    if ($mbx.IsExchangeCloudManaged -eq $targetValue) {
        return "No change needed. IsExchangeCloudManaged is already '$targetValue'."
    }

    Set-Mailbox -Identity $Identity -IsExchangeCloudManaged $targetValue -ErrorAction Stop
    Write-Log "Set-Mailbox $Identity IsExchangeCloudManaged=$targetValue" "INFO"
    return "Updated. IsExchangeCloudManaged is now '$targetValue'."
}
#endregion

#region GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "SOA Mailbox Tool - Exchange Online (IsExchangeCloudManaged)"
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

# Footer info
$lblFoot = New-Object System.Windows.Forms.Label
$lblFoot.Location = New-Object System.Drawing.Point(12, 460)
$lblFoot.Size = New-Object System.Drawing.Size(940, 130)
$lblFoot.Text =
"Notes:
- This tool changes IsExchangeCloudManaged only (does not migrate mailboxes).
- Before enabling cloud SOA, Microsoft recommends allowing your normal sync cycle + 24 hours after on-prem mailbox attribute changes.
- Before reverting to on-prem SOA, export/backup cloud-side changes you want to keep; next sync will overwrite cloud values with on-prem values.
Log: $($Script:LogFile)
Exports: $($Script:ExportDir)
"
$lblFoot.AutoSize = $false

# Add controls
$form.Controls.AddRange(@($btnConnect,$btnDisconnect,$lblConn,$grpSearch,$grpDetails,$lblFoot))
$grpSearch.Controls.AddRange(@($txtSearch,$btnSearch,$grid))
$grpDetails.Controls.AddRange(@($txtDetails,$btnRefresh,$btnBackup,$btnEnableCloud,$btnRevertOnPrem))

# State tracking
$Script:SelectedIdentity = $null

function Set-UiConnectedState {
    param([bool]$Connected)

    $btnConnect.Enabled    = -not $Connected
    $btnDisconnect.Enabled = $Connected
    $btnSearch.Enabled     = $Connected

    $btnRefresh.Enabled    = $false
    $btnBackup.Enabled     = $false
    $btnEnableCloud.Enabled= $false
    $btnRevertOnPrem.Enabled= $false

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

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Mailbox:")
    $lines.Add("  DisplayName              : $($mbx.DisplayName)")
    $lines.Add("  PrimarySmtpAddress       : $($mbx.PrimarySmtpAddress)")
    $lines.Add("  RecipientTypeDetails     : $($mbx.RecipientTypeDetails)")
    $lines.Add("  IsDirSynced              : $($mbx.IsDirSynced)")
    $lines.Add("  IsExchangeCloudManaged   : $($mbx.IsExchangeCloudManaged)")
    $lines.Add("  ExchangeGuid             : $($mbx.ExchangeGuid)")
    $lines.Add("  ExternalDirectoryObjectId: $($mbx.ExternalDirectoryObjectId)")

    if ($usr) {
        $lines.Add("")
        $lines.Add("User:")
        $lines.Add("  UserPrincipalName        : $($usr.UserPrincipalName)")
        $lines.Add("  ImmutableId              : $($usr.ImmutableId)")
        $lines.Add("  WhenChangedUTC           : $($usr.WhenChangedUTC)")
    }

    $txtDetails.Lines = $lines.ToArray()

    $btnRefresh.Enabled     = $true
    $btnBackup.Enabled      = $true
    $btnEnableCloud.Enabled = $true
    $btnRevertOnPrem.Enabled= $true
}

# Events
$btnConnect.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        if (Connect-EXO) {
            Set-UiConnectedState -Connected $true
        }
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnDisconnect.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Disconnect-EXO
        Set-UiConnectedState -Connected $false
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnSearch.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $q = $txtSearch.Text
        Write-Log "Search query: $q" "INFO"
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
        Write-Log "Search failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Search failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$grid.Add_SelectionChanged({
    try {
        if ($grid.SelectedRows.Count -gt 0) {
            $row = $grid.SelectedRows[0]
            $smtp = $row.Cells["PrimarySmtpAddress"].Value
            if ($smtp) {
                $Script:SelectedIdentity = $smtp.ToString()
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
        Show-Details -Identity $Script:SelectedIdentity
    } catch {
        Write-Log "Refresh failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Refresh failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnBackup.Add_Click({
    if (-not $Script:SelectedIdentity) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
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
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnEnableCloud.Add_Click({
    if (-not $Script:SelectedIdentity) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Enable cloud SOA for Exchange attributes for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = TRUE.",
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $true
        [System.Windows.Forms.MessageBox]::Show(
            $msg,
            "Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        Show-Details -Identity $Script:SelectedIdentity
    } catch {
        Write-Log "Enable cloud SOA failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Enable cloud SOA failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnRevertOnPrem.Add_Click({
    if (-not $Script:SelectedIdentity) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Revert SOA back to on-prem for:`n`n$($Script:SelectedIdentity)`n`nThis sets IsExchangeCloudManaged = FALSE.`n`nWARNING: Next sync may overwrite cloud values with on-prem values.",
        "Confirm",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Encourage backup before revert
        $backupPrompt = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to export a backup (JSON) before reverting?",
            "Backup recommended",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($backupPrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
            $path = Export-MailboxBackup -Identity $Script:SelectedIdentity
            Write-Log "Backup before revert: $path" "INFO"
        }

        $msg = Set-MailboxSOACloudManaged -Identity $Script:SelectedIdentity -EnableCloudManaged $false
        [System.Windows.Forms.MessageBox]::Show(
            $msg,
            "Done",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        Show-Details -Identity $Script:SelectedIdentity
    } catch {
        Write-Log "Revert to on-prem SOA failed: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Revert failed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$form.Add_FormClosing({
    try { Disconnect-EXO } catch {}
})

# Initialize UI
Set-UiConnectedState -Connected $false
Write-Log "SOA Mailbox Tool started." "INFO"

[System.Windows.Forms.Application]::Run($form)
#endregion
