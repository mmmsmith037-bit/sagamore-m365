# sagamore-m365

Installs and manages the Sagamore Alliance SharePoint + Microsoft Teams environment: security groups,
the team site, a locked-down folder structure, one **private** Teams channel per client, and a local
**Permission Manager** dashboard for changing who can see what.

**Clients are silos.** Nobody but the Owners can see a client until they are invited to it by name. An invitation
opens that client's folder *and* adds them to its private Teams channel; removing it takes both away. Client
Invoices folders stay Owner-only even for invited people.

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
- `clients` — client folder names, and under each client's `access` the people invited to it by UPN
  (or import the clients later with step 6 and invite people in the dashboard)

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
| `Clients/<client>` | Full Control | invited only* | invited only* |
| `Clients/<client>/Invoices` (any depth) | Full Control | — | — |
| Teams channel per client | channel owner | invited only* | invited only* |

\* Each client's `access` lists the invited people by UPN and their level, for example
`"jacquelin@sagamorealliance.com": "Edit"`. A client with no `access` is Owner-only. Whole groups
("all Members", "all Support") are rejected anywhere under a client, so a silo can't be opened by accident.

## Client silos and Teams channels

- Every client channel is a **private** channel. Standard channels are only used for `teams.extraChannels` (e.g. Internal).
- Channel owners are the Owner people. Channel members are exactly the people in that client's `access`.
  Every run invites anyone missing and **removes anyone who isn't invited**. Installer admins (`site.adminUpns`)
  are never removed.
- To invite someone: dashboard → click the client folder → pick the person → **Invite** → Save, then run
  `./scripts/Install-Sagamore.ps1 -PermissionsOnly`. To uninvite, click × next to them and do the same.
- Someone given access only to a client *subfolder* (e.g. Laci "Add Only" on Meridian Deliverables) gets that
  folder but is **not** in the channel. The dashboard warns that they can only reach it by link.
- Each client channel has a **Client Files** tab pointing at `Clients/<client>`. Files dropped in the channel's
  chat go to the private channel's own SharePoint site instead. Teams limits that site to channel members, so the silo still
  holds, but invoices belong in the Client Files tab's `Invoices` folder.
- If a client channel already exists as a *standard* channel (for example created by hand), Teams can't
  convert it. The installer reports **SILO BREACH** and leaves it alone, and the audit fails. Move anything you need out of it, delete it in Teams, and re-run.

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

- **Standard channel folders**: standard channels (Internal) create a folder at the library root. That's how Teams works.
  Private client channels don't. Each gets its own small SharePoint site instead.
- **Team membership**: everyone in config is a member of the Team itself (needed before they can be invited
  to a private channel). That only shows them the standard channels. Client channels stay invisible until invited.
- **Access below a hidden parent**: someone with access to a subfolder but not its parent can only reach it from a
  link or "Shared with me". The dashboard warns when that happens.
- **Removing people**: the installer warns about extra members in the security groups or Team but never removes
  them from those automatically. Remove them in Entra ID / Teams. Client channel membership *is* enforced (see above).

## Tests

`./tests/Invoke-Tests.ps1` runs offline (no tenant):
- config validation and folder-plan logic (Invoices lock, Owner protection, client-silo rule, ordering, name sanitising)
- private-channel plan: who is invited to which client channel, and the invite/remove comparison
- **parity**: the dashboard's JavaScript and the installer's PowerShell produce the identical folder plan, channel invitations and validation errors
- **contract**: every PnP cmdlet and parameter the scripts use exists in the installed PnP.PowerShell
- PSScriptAnalyzer (errors) and parse checks

These tests do not replace a first run against a real tenant. Do the first install with `-WhatIf`, then for real,
then run the audit.
