# CLAUDE.md

Guidance for AI assistants (Claude Code and others) working in this repository.

## What this repository is

**PSforO365** is a collection of PowerShell scripts and helper modules for
administering Microsoft 365 / Office 365, Exchange (Online and on‑premises),
Active Directory, SharePoint, and Microsoft Teams. It is a script *library*,
not a single application: each `.ps1` is a standalone tool an administrator
runs ad hoc or from a scheduled task. There is no build step, package manifest,
test suite, or CI pipeline.

Scripts target **Windows PowerShell 5.1** and the classic Microsoft admin
modules (`ExchangeOnlineManagement`, `MSOnline`, `AzureAD`/`Az`, `ActiveDirectory`).
They are intended to run on a Windows admin workstation with the relevant
modules installed and network access to the tenant / on‑prem servers.

## Repository layout

Most scripts live at the repository root. Notable groupings:

- **Shared modules**
  - `PSO365-Utilities.psm1` — the shared helper module. Provides logging,
    connection helpers, module/prereq validation, CSV helpers, and a
    destructive‑action confirmation prompt. Newer/hardened scripts import this.
  - `HybridManagement.psm1` — Microsoft Graph–based helpers for Exchange hybrid
    application/agent management and hybrid connectivity testing.
- **Automation hub**
  - `Invoke-O365Automation.ps1` — interactive (or `-Task`‑driven) menu that
    orchestrates the most common workflows by calling the individual scripts.
- **Connection**
  - `Connectto365.ps1` — connects to Exchange Online (modern EXO module, or
    `-LegacyMode` basic‑auth PSSession).
- **Exchange / mailbox tools** — `DisableIMAPandPOP.ps1`, `Forwarding.ps1`,
  `FullMailboxPermissionsforAdmin.ps1`, `Get-MbPermsAll.ps1`,
  `Exchange-SetUserProxy.ps1`, `MoveRequestReport.ps1`, `Impersonation.ps1`,
  `ExportO365UserInfo.ps1`, etc.
- **Identity / AD tools** — `UPNChange.PS1`, `licenses.ps1`, `CheckNTNames.ps1`,
  `Audit-ADTrusts.ps1`, `OU_permissions.ps1`, `Get-TokenSizeReport.ps1`,
  `PrivilegedUser3.0.ps1`, `New-ADAssetReport.ps1`.
- **Reports / health** (large third‑party or vendor scripts, mostly left as‑is)
  — `HealthChecker.ps1`, `Test-ExchangeServerHealth.ps1`,
  `Get-ExchangeEnvironmentReport.ps1`, `Get-ExchangeCertificateReport.ps1`,
  `m365secureScore.ps1`.
- **`Exchange Hybrid Discovery/`** — a self‑contained bundle of discovery /
  reporting scripts (some duplicate root scripts) used during hybrid migrations.
- **`Bulk User Creation/`** — `NewUserPS.ps1` plus a sample `NewUserCSV.csv`.
- **`Batch-Analysis-of-Permission-threads-master/`** — vendored third‑party tool
  (Excel macro + guide) for mailbox delegate analysis.
- **`.zip` files, `.docx`, `.txt`, and extensionless files** — assorted
  reference material, packaged scripts, and snippets. Treat these as data/docs,
  not code to maintain. Do **not** unpack or rewrite them unless asked.

## Conventions for the hardened scripts

Several root scripts were refactored to a common house style (see commit
`8ffefca` "harden existing scripts and add automation hub"). When **editing or
adding** a first‑party admin script, follow this pattern:

1. **Version pin & comment‑based help** at the top:
   ```powershell
   #Requires -Version 5.1
   <#
   .SYNOPSIS  ...
   .DESCRIPTION ...
   .PARAMETER ...
   .EXAMPLE ...
   .NOTES ...
   #>
   ```
2. **`[CmdletBinding(SupportsShouldProcess)]`** with a typed `param()` block.
   Use `SupportsShouldProcess` so `-WhatIf` works for any change‑making script.
3. **Bootstrap the shared module** with a graceful fallback:
   ```powershell
   $utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
   if (Test-Path $utilitiesPath) {
       Import-Module $utilitiesPath -Force
       Initialize-Log -ScriptName "<ScriptName>"
   }
   else {
       function Write-Log { param([string]$Message, [string]$Level = 'INFO') Write-Host "[$Level] $Message" }
   }
   ```
4. **Use the shared helpers** rather than re‑implementing them:
   - `Write-Log -Message ... -Level INFO|WARN|ERROR|SUCCESS`
   - `Connect-O365` / `Connect-MSOnline`
   - `Assert-ModuleAvailable -ModuleName ...`, `Assert-RunningAsAdmin`
   - `Get-CsvFilePath`, `Import-CsvSafe -Path ... -RequiredColumns ...`
   - `Confirm-DestructiveAction -ActionDescription "..." [-Force]`
5. **Safety first.** Validate inputs, detect duplicates/existing objects before
   creating, wrap per‑item operations in `try/catch`, and print a summary of
   successes / failures / skips at the end.
6. **Gate destructive or bulk operations** behind `Confirm-DestructiveAction`
   (or native `ShouldProcess`) and honor a `-Force` switch to bypass prompts for
   automation.
7. **Don't hardcode tenant‑specific values** (domains, UPNs, OU paths). Take
   them as parameters with sensible documented defaults.

`.NOTES` in the hardened scripts often record an "Improved / Original" diff —
keep that history note accurate if you further change the script.

Not every file follows this style. Large vendor/report scripts (`HealthChecker.ps1`,
`Test-ExchangeServerHealth.ps1`, the `Get-Exchange*Report.ps1` family) are
effectively third‑party — **do not reformat or refactor them wholesale**; make
only the minimal change requested.

## The shared utilities module

`PSO365-Utilities.psm1` exports (and is the canonical source for):

- Logging: `Initialize-Log`, `Write-Log` (writes colored console output and,
  once initialized, appends to a timestamped file under `.\Logs`).
- Connection: `Connect-O365` (EXO modern, `-LegacyMode` for PSSession),
  `Connect-MSOnline`.
- Validation: `Assert-ModuleAvailable`, `Assert-RunningAsAdmin`.
- Files: `Get-CsvFilePath` (WinForms file picker), `Import-CsvSafe`
  (existence/empty/required‑column validation).
- Safety: `Confirm-DestructiveAction`.

If you add a broadly useful helper, put it here and add it to the
`Export-ModuleMember` list at the bottom of the module.

## Connection & auth notes

- Prefer the **modern** `ExchangeOnlineManagement` module path; `-LegacyMode`
  basic‑auth PSSession exists only for old tenants and is deprecated by Microsoft.
- `HybridManagement.psm1` and some older scripts acquire Microsoft Graph tokens
  via ADAL with the well‑known client ID `1950a258-227b-4e31-a9cf-717495945fc2`.
- Never commit real credentials, tokens, tenant IDs, or customer data. Sample
  CSVs (e.g. `Bulk User Creation/NewUserCSV.csv`, `Exchange Hybrid Discovery/import.csv`)
  are illustrative — keep them generic.

## Running / verifying

There is no automated test harness and this environment is Linux, so scripts
generally **cannot be executed here** (they need Windows + tenant connectivity).
When changing a script:

- Validate syntax mentally / with PowerShell parsing if available; do not assume
  a script "passes" — say so plainly if you could not run it.
- Keep comment‑based help (`.EXAMPLE`, `.PARAMETER`) in sync with actual params.
- If `pwsh` is available, you can sanity‑check parsing with:
  `pwsh -NoProfile -Command "[void][System.Management.Automation.Language.Parser]::ParseFile('<file>',[ref]$null,[ref]$null)"`

## Git workflow

- Default branch: `master`. Remote: `bergeronk/psforo365` (GitHub).
- Develop on the branch you were assigned for the task; create it locally if
  needed. Push with `git push -u origin <branch>`.
- Commit messages: short imperative subject describing the change
  (e.g. "harden DisableIMAPandPOP and add WhatIf support").
- **Do not** open a pull request unless explicitly asked.
- Line endings: scripts are Windows‑oriented; preserve existing CRLF/encoding of
  a file you edit rather than normalizing it.

## Quick orientation for common requests

- "Add a new admin script" → follow the hardened‑script pattern above, import
  `PSO365-Utilities.psm1`, and consider wiring it into `Invoke-O365Automation.ps1`.
- "Fix/improve an existing root script" → check whether it already uses the
  shared module; align it with the conventions, keep it `-WhatIf`‑safe.
- "Touch a big report/health script" → minimal change only; treat as vendor code.
