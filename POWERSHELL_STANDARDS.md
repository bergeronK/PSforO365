# PowerShell Scripting Standards

A project-agnostic checklist for writing consistent, repeatable, maintainable PowerShell. Copy this file into any project; it doesn't assume a specific module or repo layout.

---

## 1. Script Header

Every `.ps1` and `.psm1` file starts with:

```powershell
#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    One-line summary.
.DESCRIPTION
    Full description of what the script does and why.
.PARAMETER UserUPN
    The UPN of the target user.
.INPUTS
    System.String. UPNs can be piped in.
.OUTPUTS
    PSCustomObject with UserUPN, Status, Timestamp.
.EXAMPLE
    .\Disable-MailboxProtocol.ps1 -UserUPN "user@domain.com" -WhatIf
.NOTES
    Version: 1.0.0
    Requires: PowerShell 7.0+, ExchangeOnlineManagement 3.x
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [string]$UserUPN
)

Set-StrictMode -Version Latest
```

- Pin a `#Requires -Version` and pin module versions where the cloud API surface changes often (Graph, ExchangeOnlineManagement).
- State a PowerShell 7 vs. Windows PowerShell 5.1 stance explicitly per project — don't leave it implicit. If a script depends on a Windows-only module (e.g., `ActiveDirectory`), say so in `.NOTES`.
- Support pipeline input (`ValueFromPipeline` / `ValueFromPipelineByPropertyName`) for any parameter that represents "one of many target objects" — it's what makes bulk scripts composable.

---

## 2. Parameters & Validation

Validate at the boundary, not deep in the logic:

```powershell
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserUPN,

    [ValidateSet('E1', 'E3', 'E5', 'F1')]
    [string]$LicenseSku = 'E3',

    [ValidateRange(1, 365)]
    [int]$RetentionDays = 30,

    [switch]$Force
)
```

Everything that changes between environments or runs is a parameter (tenant, identities, file paths, SKUs, feature flags). Things that are genuinely fixed (script version, default retry count) can be constants.

---

## 3. ShouldProcess, WhatIf, and Confirm — get the mechanics right

`SupportsShouldProcess` only works if the code actually gates the action behind `ShouldProcess()`. Declaring the attribute without the gate gives users a `-WhatIf` switch that does nothing — silently.

```powershell
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [switch]$Force
)

if ($Force -and -not $PSBoundParameters.ContainsKey('Confirm')) {
    $ConfirmPreference = 'None'
}

if ($PSCmdlet.ShouldProcess($UserUPN, "Disable IMAP/POP")) {
    Set-CASMailbox -Identity $UserUPN -ImapEnabled $false -Confirm:$false
}
```

Rules of thumb:
- `ConfirmImpact` default is `Medium` and does **not** prompt by default (`$ConfirmPreference` defaults to `High`). Reserve `High` for genuinely destructive, hard-to-reverse actions (bulk delete, irreversible config changes).
- `-WhatIf` alone suppresses the action; `-Confirm` alone prompts. Add your own `-Force` switch (pattern above) — most users try `-Force` before they discover `-Confirm:$false`.
- When calling other ShouldProcess-aware cmdlets inside a gated block, pass `-Confirm:$false` to them — otherwise `$ConfirmPreference` cascades and users get double-prompted.
- `ShouldProcess` only intercepts cmdlet calls. If your script shells out to an external binary or calls a .NET method directly to make the change, `-WhatIf` won't help — gate those manually.

---

## 4. Error Handling — know which stream you're catching

Most Exchange Online / Graph cmdlets raise **non-terminating** errors by default. A bare `try/catch` around them does **not** catch failures unless you force the error to terminate:

```powershell
try {
    Set-Mailbox -Identity $UserUPN -Type Shared -ErrorAction Stop
    Write-Verbose "Converted $UserUPN to shared mailbox."
    $successCount++
}
catch {
    Write-Error "Failed to convert ${UserUPN}: $_"
    $failCount++
}
```

- Set `-ErrorAction Stop` **per call**, not `$ErrorActionPreference = 'Stop'` globally. A global override changes behavior of every cmdlet in the session (including ones you didn't intend) and makes failures harder to localize.
- Classify failures explicitly:
  - **Fail fast** — anything that invalidates the entire run: missing module, failed auth/connection, invalid config. Stop immediately, don't process any items.
  - **Fail gracefully** — anything scoped to a single target object: one user out of 500 not found. Log it, increment a counter, move to the next item.
  - If you're unsure which a given failure is, ask: "does this error mean the rest of the run is meaningless?" If yes, fail fast.
- For cloud API throttling (HTTP 429 / `TooManyRequests` from Graph or EXO), implement retry with exponential backoff — don't treat a throttle as a hard failure:

```powershell
function Invoke-WithRetry {
    param([scriptblock]$ScriptBlock, [int]$MaxAttempts = 5)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try { return & $ScriptBlock }
        catch {
            if ($attempt -eq $MaxAttempts -or $_.Exception.Message -notmatch '429|throttl') { throw }
            Start-Sleep -Seconds ([math]::Pow(2, $attempt))
        }
    }
}
```

---

## 5. Output Stream Discipline

| Stream | Use for |
|---|---|
| Pipeline output (`[PSCustomObject]`) | Data the caller should be able to filter, sort, pipe, or export |
| `Write-Verbose` | Step-by-step progress detail (`-Verbose` to see it) |
| `Write-Information` | User-facing status messages that aren't pipeline data |
| `Write-Warning` | Recoverable problems worth flagging |
| `Write-Error` | Errors (non-terminating unless `-ErrorAction Stop`) |
| `Write-Host` | Almost never — only for interactive console formatting that must not be captured |

```powershell
[PSCustomObject]@{
    UserUPN   = $UserUPN
    Status    = 'Converted'
    Timestamp = Get-Date
}
```

Never `Write-Host` data the caller might need — it can't be captured, piped, or redirected. Reserve it for things like a colored interactive menu.

---

## 6. Idempotency & Repeatability

A repeatable script produces the same outcome whether run once or ten times, and works unmodified against a different tenant/environment given different parameters.

```powershell
if (-not (Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue)) {
    if ($PSCmdlet.ShouldProcess($GroupName, "Create distribution group")) {
        New-DistributionGroup -Name $GroupName
    }
}
else {
    Write-Verbose "$GroupName already exists, skipping."
    $skipCount++
}
```

Mental model:

```
1. Validate prerequisites (module, connection, config)   ← fail fast
2. Connect / authenticate
3. For each target item:
   a. Check current state
   b. Skip if already in desired state (not an error)
   c. Apply change if needed, gated by ShouldProcess
   d. Log result
4. Report summary: succeeded / skipped / failed counts
```

Anti-patterns to avoid:
- Hardcoded environment-specific values
- No state check before a write operation
- Stopping the entire run on the first per-item failure
- No output — caller can't tell what happened
- A script that silently assumes a previous script already ran, with no validation

---

## 7. Module Structure

For shared logic, use a real module with a manifest, not just a loose `.psm1`:

```
MyProject/
├── Modules/
│   └── MyProject.Utilities/
│       ├── MyProject.Utilities.psd1   ← manifest: version, exported functions, required modules
│       └── MyProject.Utilities.psm1
├── Scripts/
│   └── Disable-MailboxProtocol.ps1    ← thin orchestrator, calls module functions
└── Tests/
    └── MyProject.Utilities.Tests.ps1
```

```powershell
# MyProject.Utilities.psd1
@{
    RootModule        = 'MyProject.Utilities.psm1'
    ModuleVersion      = '1.2.0'
    FunctionsToExport  = @('Write-Log', 'Connect-Tenant', 'Assert-ModuleAvailable')
    RequiredModules    = @(@{ ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.4.0' })
}
```

- Use a manifest (`.psd1`) with explicit `FunctionsToExport` — avoid `'*'`, it's flagged by PSScriptAnalyzer and makes the module's public surface unclear.
- Pin `RequiredModules` versions for cloud SDKs (Graph, EXO) — these change behavior across versions and break scripts silently otherwise.
- Internal helpers stay unexported; use `$script:` scope for module state, never `$global:`.
- Break up any script over ~300–400 lines into module functions + a thin orchestrator script.

---

## 8. Naming Conventions

| Element | Convention | Example |
|---|---|---|
| Scripts | `Verb-Noun.ps1` (approved verbs) | `Disable-MailboxProtocol.ps1` |
| Functions | `Verb-Noun` (run `Get-Verb` to check) | `Get-MailboxPermission` |
| Parameters | `PascalCase` | `$UserUPN`, `$RetentionDays` |
| Local variables | `camelCase` | `$mailboxList`, `$failCount` |
| Module-scope variables | `$script:camelCase` | `$script:LogFile` |
| Constants | `$UPPER_SNAKE` | `$MAX_RETRY_ATTEMPTS = 5` |

There are competing community conventions for local variable casing — the convention itself matters less than applying it consistently across the whole project.

---

## 9. Linting and Testing

- **PSScriptAnalyzer** is mandatory, not optional. Add a `PSScriptAnalyzerSettings.psd1` at the project root and run it in CI / pre-commit:
  ```powershell
  Invoke-ScriptAnalyzer -Path . -Recurse -Settings ./PSScriptAnalyzerSettings.psd1 -EnableExit
  ```
- **Pester** for anything beyond a trivial script — at minimum, unit-test pure logic functions in the utility module, mocking cloud calls (`Connect-ExchangeOnline`, `Get-MgUser`, etc.) so tests don't require a live tenant:
  ```powershell
  Describe 'Get-DesiredState' {
      It 'returns Skip when group already exists' {
          Mock Get-DistributionGroup { return @{ Name = 'Finance' } }
          (Get-DesiredState -GroupName 'Finance').Action | Should -Be 'Skip'
      }
  }
  ```

---

## 10. Security

- Never hardcode credentials, tokens, or secrets in script source.
- Use `Get-Credential` for interactive runs; use `Microsoft.PowerShell.SecretManagement` or a cloud secret store (Key Vault, etc.) for automation.
- Prefer managed identity / certificate-based app auth over stored passwords for unattended connections.
- Add secret scanning (e.g., gitleaks) as a pre-commit or CI gate so a credential never lands in git history in the first place — rotating after the fact is far more expensive than catching it before commit.
- Sign scripts and set an appropriate execution policy in environments where script integrity matters; don't rely on `-ExecutionPolicy Bypass` as a permanent fix.

---

## 11. Performance Basics

- Don't grow arrays with `+=` in a loop over large collections — it reallocates every iteration. Use a `[System.Collections.Generic.List[T]]` or let the pipeline emit objects directly.
- Prefer pipeline-native cmdlets over manual `foreach` + accumulation when processing large result sets from EXO/Graph — it avoids holding the whole set in memory.
- Page large Graph/EXO queries instead of pulling entire directories into memory at once.

---

## 12. Minimal Script Template

```powershell
#Requires -Version 7.0
<# .SYNOPSIS  One-line summary. #>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [string]$UserUPN
)

begin {
    Set-StrictMode -Version Latest
    $successCount = 0
    $skipCount    = 0
    $failCount    = 0
}

process {
    try {
        # check current state, skip if already correct, otherwise:
        if ($PSCmdlet.ShouldProcess($UserUPN, "Apply change")) {
            # do the work, -ErrorAction Stop on the cmdlet call
            $successCount++
        }
        else {
            $skipCount++
        }
    }
    catch {
        Write-Error "Failed for ${UserUPN}: $_"
        $failCount++
    }
}

end {
    [PSCustomObject]@{
        Succeeded = $successCount
        Skipped   = $skipCount
        Failed    = $failCount
    }
}
```

---

## 13. Pre-Merge Checklist

- [ ] `#Requires` and comment-based help present
- [ ] `[CmdletBinding(SupportsShouldProcess)]` on anything that writes, with the action actually gated
- [ ] Parameters validated (`ValidateSet`/`ValidatePattern`/`ValidateRange` as applicable)
- [ ] `-ErrorAction Stop` on calls inside `try`, not a global `$ErrorActionPreference`
- [ ] Per-item failures logged and counted, not fatal to the whole run
- [ ] Pipeline output is structured objects, not `Write-Host`
- [ ] Re-running the script against unchanged state is a no-op (skips, not duplicates/errors)
- [ ] No hardcoded secrets or environment-specific values
- [ ] `Invoke-ScriptAnalyzer` passes with no errors
- [ ] Pester tests exist for non-trivial logic
