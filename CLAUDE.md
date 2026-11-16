# sagamore-m365: notes for Claude

PowerShell 7.4 + PnP.PowerShell 3.x installer, audit and Permission Manager dashboard for the Sagamore Alliance
SharePoint/Teams tenant. Read README.md first; it is the source of truth for behaviour.

## Rules
- Run `./tests/Invoke-Tests.ps1` after every change. All tests are offline and must stay green.
- Plan logic exists twice: `src/Sagamore.Common.psm1` (PowerShell) and `dashboard/plan.js` (JS). Any change to
  validation, folder plan or channel plan must be made in **both**. `tests/Parity.Tests.ps1` enforces this.
- Every PnP cmdlet/parameter must exist in the installed module (`tests/Contract.Tests.ps1`).
- `config/sagamore.config.json` is the real tenant config: git-ignored, never commit it.
- Never run anything that signs in to or changes the tenant (`Install-Sagamore.ps1` without `-PlanOnly`,
  `Test-SagamoreInstall.ps1`, `Register-SagamoreApp.ps1`, `Import-SagamoreClients.ps1`, `Uninstall-Sagamore.ps1`)
  unless Marshall asks for it in the conversation. `-PlanOnly` is offline and always safe.
- `git` is not on PATH on this machine.

## Status (2026-11-16)
Code complete: installer, audit, client import, uninstall, dashboard, and the one-click local runner
(`scripts/Start-SagamoreDashboard.ps1`). 76 offline tests pass.
Not done yet (needs Marshall / tenant access, can't be done offline):
1. Create `config/sagamore.config.json` with real UPNs, roles and clients.
2. `Register-SagamoreApp.ps1` (Global Admin, possibly AmShot).
3. First install: `-WhatIf`, then for real, then the audit.
4. Import Joe's existing client folders.
5. Hand the dashboard over to Joe and train staff.
