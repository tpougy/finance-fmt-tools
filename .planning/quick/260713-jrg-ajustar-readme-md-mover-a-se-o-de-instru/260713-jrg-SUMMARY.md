---
phase: quick-260713-jrg
plan: 01
status: complete
subsystem: docs
tags: [readme, documentation, installer, vba-migration]

dependency-graph:
  requires: []
  provides:
    - "README.md with `## Instalação` as the first top-level section (right after the title/intro block)"
    - "`### Atualizando da versão VBA` now documents the v2.1.0 automatic legacy VBA removal (Remove-LegacyVbaAddin in scripts/install.ps1)"
  affects:
    - README.md

tech-stack:
  added: []
  patterns:
    - "Blockquote note style (`> ...`) reused for the new automatic-migration callout, matching existing README blockquotes (VBA→C# migration note in Introdução, checksum note in Instalação)"

key-files:
  created: []
  modified:
    - README.md

decisions:
  - "Removed the `---` separator that used to sit immediately after the old `### Atualizando da versão VBA` content (between old Instalação and Referência rápida do ribbon), keeping the one before it — leaves exactly one `---` between `## Outros formatos` and `## Referência rápida do ribbon` post-move, per plan's Task 1 instruction to remove exactly one of the two flanking separators."
  - "Wrote the new blockquote sentence verbatim as specified in the plan (sourced from RELEASE_NOTES.md / scripts/install.ps1 behavior), with no paraphrasing."

metrics:
  duration: "~10 minutes"
  completed: 2026-07-13
---

# Quick Task 260713-jrg: Move installation section to top of README, document VBA auto-migration Summary

Reordered `README.md` so `## Instalação` is the first `##` section (right after the title/description block, before `## Introdução`), and added a blockquote note to `### Atualizando da versão VBA` explaining that `scripts/install.ps1` (since v2.1.0) auto-detects and removes a legacy VBA `.xlam` install before installing the C# version — no manual steps required.

## What Was Built

**Task 1 — Move `## Instalação` to the top** (commit `b37d8d2`):
- Cut the entire `## Instalação` section (heading through "Depois disso, rode o instalador acima normalmente.", including the nested `### Atualizando da versão VBA` subsection) from its original position between `## Outros formatos` and `## Referência rápida do ribbon`.
- Removed the `---` separator that used to follow the section (kept the one before it), leaving exactly one `---` between `## Outros formatos` and `## Referência rápida do ribbon`.
- Pasted the section, unchanged, immediately after the `---` that closes the title block and before `## Introdução`, adding a new `---` after the pasted content.
- No text inside the section was altered in this task — pure reorder.

**Task 2 — Add VBA auto-migration note + final integrity check** (commit `bec4451`):
- Inserted a new blockquote paragraph immediately after the `### Atualizando da versão VBA` heading (before the existing "Se você já tinha o add-in antigo..." line), stating verbatim that since v2.1.0 `scripts/install.ps1` automatically detects, unregisters (via COM), and deletes a legacy VBA `.xlam` install before installing the C# version, with the old manual steps kept only as a fallback reference.
- Left the rest of the subsection (4-step manual list, closing sentence) untouched.

## Verification

- Task 1 gate: `diff` of `## ` heading order against the expected 9-section order, `diff` of `### ` heading order (with "Atualizando da versão VBA" now first), separator count == 9, `#### ` count == 4 — all passed (`TASK1_OK`).
- Task 2 gate: new blockquote sentence present verbatim as the first non-blank line after the `### Atualizando da versão VBA` heading, heading counts unchanged (`##`=9, `###`=11, `####`=4), 9 anchor phrases (one per original top-level section) still present verbatim, `git status --porcelain` shows only `README.md` modified — all passed (`TASK2_OK`).
- Final manual re-check: `## ` and `### ` heading listings, separator count, and `git log`/`git status` all confirm the expected end state.

## Deviations from Plan

None — plan executed exactly as written. Both tasks followed the `<action>` specifications literally (exact cut/paste boundaries, exact verbatim blockquote text, no other text altered).

## Threat Flags

None — this is a documentation-only change; no new network endpoints, auth paths, file access patterns, or schema changes were introduced. The plan's own threat register (T-quick-jrg-01/02/03) already covers content-loss and stale-documentation risks, both mitigated by the automated gates above.

## Self-Check: PASSED

- FOUND: README.md
- FOUND: commit b37d8d2 (Task 1 — move `## Instalação` to top)
- FOUND: commit bec4451 (Task 2 — add VBA auto-migration note)
