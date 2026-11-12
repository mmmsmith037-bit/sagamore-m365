# sagamore-m365

Installs and manages the Sagamore Alliance SharePoint + Microsoft Teams environment: security groups,
the team site, a locked-down folder structure (client Invoices are Owner-only), one Teams channel per
client, and a local **Permission Manager** dashboard for changing who can see what.

One file drives everything: `config/sagamore.config.json`. The dashboard edits it; the installer makes the
tenant match it; the audit proves the tenant matches it.

```
dashboard/index.html  ──edits──▶  config/sagamore.config.json  ──Install-Sagamore.ps1──▶  SharePoint + Teams
                                                     ▲                                          │
                                                     └──────── Test-SagamoreInstall.ps1 ◀───────┘ (audit)
```

## Requirements

| What | Why |
|---|---|
| Windows PowerShell **7.4+** (`pwsh`) | PnP.PowerShell 3.x requires it |
| `Install-Module PnP.PowerShell -Scope CurrentUser` | The only module used |
| A **Global Administrator** of the Sagamore tenant, once | To register the app and grant consent (step 2) |
| The person running the installer must be an Owner in config or listed in `site.adminUpns` | So they never lose access to a folder they are locking down |
| SharePoint Administrator role for the installer account | Only for setting the site's external-sharing level (warns and continues otherwise) |

## Install day runbook

```powershell
git clone <repo-url> sagamore-m365
cd sagamore-m365
./tests/Invoke-Tests.ps1                     # offline tests: must be all green
```

**1. Create the config**

```powershell
Copy-Item config/sagamore.config.example.json config/sagamore.config.json
```

Edit `config/sagamore.config.json`:
- `tenant.name` — the part before `.sharepoint.com`
- `people` — every staff member's real UPN and role (`Owner`, `Member`, `Support`)
- `site.adminUpns` — your own admin UPN (the account that will run the installer)
- `clients` — client folder names (or import them later with step 6)

The real config is git-ignored; it never gets committed.

**2. Register the app (once per tenant, Global Admin)**

```powershell
./scripts/Register-SagamoreApp.ps1
```

This creates the Entra ID app "Sagamore M365 Installer" with delegated permissions only. It can never do more
than the signed-in user is allowed to do. The script saves the Client ID into the config. If the AmShot MSP
holds Global Admin, they run this one step.

**3. Preview**

```powershell
./scripts/Install-Sagamore.ps1 -PlanOnly      # no sign-in: shows every folder and permission
./scripts/Install-Sagamore.ps1 -WhatIf        # signs in, reads the tenant, changes nothing
```

**4. Install**

```powershell
./scripts/Install-Sagamore.ps1
```

Takes a few minutes. A new Microsoft 365 group can take several minutes before Teams accepts it; the script
retries automatically. A log and a permissions CSV are written to `logs/`, and the CSV is also uploaded to
`_Admin/Permissions-Log`.

**5. Audit**

```powershell
./scripts/Test-SagamoreInstall.ps1            # read-only; exit code 0 = tenant matches config
```

**6. Import Joe's existing client folders (optional)**

```powershell
./scripts/Import-SagamoreClients.ps1 -SourcePath 'D:\Joe\Clients' -WhatIf    # preview the mapping
./scripts/Import-SagamoreClients.ps1 -SourcePath 'D:\Joe\Clients'            # folders + permissions
./scripts/Import-SagamoreClients.ps1 -SourcePath 'D:\Joe\Clients' -UploadFiles
```

Folders are created and **locked before any file is uploaded**, so an Invoices file is never briefly visible to staff.

## Changing permissions later (dashboard)

Open `dashboard/index.html` in Edge or Chrome → **Open config…** → pick `config/sagamore.config.json`.

- Click a **person** to see exactly what they can open in every folder.
- Click a **folder** to switch between "same as parent" and custom access, add/remove people or groups, or add subfolders.
- **Save config** writes straight back to the file (Edge/Chrome). Other browsers download it instead.
- Then run `./scripts/Install-Sagamore.ps1 -PermissionsOnly` (the dashboard shows the command).

The dashboard never signs in to Microsoft 365. It only edits the config file, and the installer applies it.
To preview with sample data, run `python -m http.server` in the repo and open `http://localhost:8000/dashboard/#example`.

## Permission model

| Where | Owner | Member | Support |
|---|---|---|---|
| Library root | Full Control | Edit | Read |
| `_Admin` | Full Control | — | — |
| `Internal` | Full Control | Edit | Read |
| `Internal/HR` | Full Control | Edit | — |
| `_Templates` | Full Control | Read | Read |
| `Clients` | Full Control | Read | Read |
| `Clients/<client>` | Full Control | Edit* | Read* |
| `Clients/<client>/Invoices` (any depth) | Full Control | — | — |

\* Default from `clientTemplate`; per-client overrides (for example "only Jacquelin works on Meridian") go in that client's `access`.

Rules the code enforces:
- The **Owner** group always keeps Full Control wherever permissions are customised. It can't be removed by config or dashboard.
- The site's own Owners group (Microsoft 365 group owners = Owner people + `site.adminUpns`) keeps Full Control everywhere, so admins are never locked out.
- Any folder named **Invoices** anywhere under a client is forced to Owner-only, including folders staff create later. Re-run `-PermissionsOnly` or schedule it.
- The library no longer inherits the site's "Team members can edit" default, so Support really is read-only.
- **Add Only** is a custom SharePoint permission level: Contribute without delete.
- External sharing defaults to *existing guests only*, and only owners can share.
- Re-running is safe. Each run makes SharePoint match the config exactly for every folder the config manages. Folders not in config are left alone, except the Invoices sweep.

## Scripts

| Script | Purpose |
|---|---|
| `scripts/Register-SagamoreApp.ps1` | One-time app registration; writes Client ID to config |
| `scripts/Install-Sagamore.ps1` | Install / re-apply. `-PlanOnly`, `-WhatIf`, `-PermissionsOnly`, `-SkipTeams` |
| `scripts/Test-SagamoreInstall.ps1` | Read-only audit with CSV report |
| `scripts/Import-SagamoreClients.ps1` | Import local client folders (+ files with `-UploadFiles`) |
| `scripts/Uninstall-Sagamore.ps1` | Remove the group, site and Team (30-day restore). Requires `-ConfirmAlias` |

## Known behaviour

- **Teams channel folders**: each channel also creates an empty folder at the library root. That's how Teams works.
  The per-client "Client Files" tab points at `Clients/<client>` instead.
- **Channel visibility**: every Team member can read every standard channel's conversation, even for clients whose
  folder they can't open. If client conversations must be private, switch those to private channels. That isn't automated yet.
- **Access below a hidden parent**: someone with access to a subfolder but not its parent can only reach it from a
  link or "Shared with me". The dashboard warns when that happens.
- **Removing people**: the installer warns about extra members in the security groups or Team but never removes
  anyone automatically. Remove them in Entra ID / Teams.

## Tests

`./tests/Invoke-Tests.ps1` runs offline (no tenant):
- config validation and folder-plan logic (Invoices lock, Owner protection, ordering, name sanitising)
- **parity**: the dashboard's JavaScript and the installer's PowerShell produce the identical plan and validation errors
- **contract**: every PnP cmdlet and parameter the scripts use exists in the installed PnP.PowerShell
- PSScriptAnalyzer (errors) and parse checks

These tests do not replace a first run against a real tenant. Do the first install with `-WhatIf`, then for real,
then run the audit.
