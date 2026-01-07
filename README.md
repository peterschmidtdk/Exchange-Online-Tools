# Exchange Mailbox SOA Manager

## Overview
**Exchange Mailbox SOA Manager** is a **PowerShell 7** (WinForms) GUI tool for Exchange Online that lets you **view** and **switch** the *State of Authority (SOA)* for Exchange attributes on **directory-synced mailboxes**.

It toggles whether Exchange attributes for a synced mailbox should be managed from:

- **Exchange Online (Cloud-managed)** → *SOA = Online* (`IsExchangeCloudManaged = True`)
- **Exchange On-Premises (On-prem managed)** → *SOA = On-Prem* (`IsExchangeCloudManaged = False`)

It is done by converting the Mailbox State of Authority (SOA) of the mailbox. Can be done on individual mailboxes.

Microsoft reference:
- *Enable Exchange attributes in Microsoft Entra ID for cloud management* (Microsoft Learn)  
  https://learn.microsoft.com/en-us/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management

---

## Why this tool exists
In hybrid environments, mailboxes are often **directory-synced** from on-premises Active Directory. Traditionally, that means Exchange attributes are authored on-premises (Exchange tools or AD attributes). Microsoft introduced support for shifting the **authoritative source** for certain Exchange attributes to the **cloud** for synced objects.

This tool provides a safe GUI-based way to:

- See which mailboxes are currently **cloud-managed** vs **on-prem managed**
- Switch SOA state per mailbox quickly
- Log all changes to a single logfile for auditing and troubleshooting

---
##Requirements

### PowerShell
- **PowerShell 7+** required (`pwsh.exe`)

### STA mode (important for WinForms)
WinForms requires **Single-Threaded Apartment (STA)** mode:

### Exchange Online PowerShell module: ExchangeOnlineManagement
Connectivity to Exchange Online endpoints
Appropriate Exchange Online permissions to run Get-Mailbox and Set-Mailbox

### Permissions / roles in Exchange Online & Microsoft 365
At minimum, the signed-in admin account must be able to:
- Run Get-Mailbox across the target scope
- Run Set-Mailbox -IsExchangeCloudManaged ... on the target recipients

Roles needed, one of the two:
- Exchange Administrator
- Global Administrator

---

## Key features

### 1) Connect / Disconnect to Exchange Online
- Connects using the **ExchangeOnlineManagement** module
- Shows connection state in the GUI
- Displays detected **tenant name** (best effort)

### 2) Load and cache all mailboxes
- Click **Load all mailboxes (cache)** to download mailboxes into memory
- Enables **fast paging** and **instant search**
- Shows how many mailboxes exist in the tenant

### 3) Browse mailboxes with paging
- Navigate large tenants using:
  - **Prev / Next**
  - Configurable **page size** (25 / 50 / 100 / 200)

### 4) Search mailboxes (cached)
- Search by:
  - `DisplayName`
  - `PrimarySMTP`
- Search runs against the cached dataset (fast and consistent)

### 5) View mailbox SOA state in the grid
The grid displays these columns:

- **DisplayName**
- **PrimarySMTP**
- **SOA Status** (Online / On-Prem / Unknown)
- **DirSynced**

### 6) Change SOA state for the selected mailbox
When a mailbox row is selected:

- **Enable SOA = Online**
  - Sets `IsExchangeCloudManaged = True`
- **Revert SOA = On-Prem**
  - Sets `IsExchangeCloudManaged = False`

Safety behavior:
- The tool **blocks changes** if the mailbox is not **DirSynced**.

### 7) Refresh selected mailbox
- Uses **Refresh selected** to re-check the mailbox and update the cache/grid after changes.

### 8) Logging (single logfile, timestamp per line)
- Logs are written to one file:
  - `.\Logs\MailboxSOAManager.log`
- Every log line includes timestamp + run details (RunId, user, tenant).
- **Open log** button opens the logfile directly.

---


```powershell
pwsh.exe -STA -ExecutionPolicy Bypass -File .\MailboxSOAManager-GUI.ps1
