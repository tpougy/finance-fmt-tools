# Phase 4: Installation & Registration - Context

**Gathered:** 2026-07-11
**Status:** Ready for planning
**Mode:** Auto-generated (discuss skipped via workflow.skip_discuss)

<domain>
## Phase Boundary

A non-admin user can install and uninstall the add-in on 64-bit Excel with a single PowerShell command, and the installed add-in survives transient errors without being silently disabled by Excel.

</domain>

<decisions>
## Implementation Decisions

### Claude's Discretion
All implementation choices are at Claude's discretion — discuss phase was skipped per user setting (full autonomous run, `/gsd-autonomous`). Use ROADMAP phase goal, success criteria, and codebase conventions to guide decisions.

### Carried-over fixed identity from Phase 3 (not discretionary — must be reused verbatim)
Phase 3's `Connect.cs` already declares and documents the fixed COM identity this installer MUST register against: GUID (CLSID) `881EFDF3-424C-4240-BCA0-714DAC2B9CD7`, ProgId `FinanceFmtTools.Connect`, AssemblyName `FinanceFmtTools.ComAddin`, Version `1.0.0.0`, Excel discovery key `HKCU\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect`. Do not invent new values — read them from `src/FinanceFmtTools.ComAddin/Connect.cs`'s header comment and `FinanceFmtTools.ComAddin.csproj`.

### Environment constraint (same as Phase 3 — carries forward)
This dev environment is Linux/WSL with no Windows, no Excel, no PowerShell, no registry. The installer/uninstaller PowerShell scripts themselves can be written and reviewed for correctness here, but actually RUNNING them (registering the add-in in HKCU, installing the compiled DLL, verifying `DoNotDisableAddinList`, confirming the Ribbon tab appears after Excel restart) cannot be executed or tested in this environment. Expect this phase's verification to also come back `human_needed` for the actual install/uninstall run, same as Phase 3. Write correct, complete PowerShell that a human can run on their real Windows+Excel machine — do not skip or fake this reality.

</decisions>

<code_context>
## Existing Code Insights

The legacy VBA installer (`Install-FinanceFmtTools.ps1`, still present in this repo's root, distributing `FinanceFmtTools.xlam`) is a close analog for HKCU-only, no-admin registration patterns, GitHub Releases download flow, and Excel COM automation checks (`Excel.Application` process detection) — but it registers a `.xlam` Add-In, not a COM Shared Add-in, so the registry key shape differs (Excel `Add-Ins` collection vs. `HKCU\Software\Microsoft\Office\Excel\Addins\<ProgId>` COM add-in registration with `LoadBehavior`, `FriendlyName`, `Description`, and the `DoNotDisableAddinList` INST-03 key). CLAUDE.md's stated project inspiration, the sibling `outlook-classic-delay-send` project, is the closer analog for this exact COM-add-in-registration pattern — check if it has an installer script to reuse patterns from (its own BUILD.md/install scripts were referenced in Phase 3's research).

</code_context>

<specifics>
## Specific Ideas

No specific requirements — discuss phase skipped. Refer to ROADMAP phase description and success criteria (INST-01, INST-02, INST-03 in REQUIREMENTS.md).

</specifics>

<deferred>
## Deferred Ideas

None — discuss phase skipped.

</deferred>
