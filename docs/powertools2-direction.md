# PowerTools 2 — direction

Written 2026-11-03 from a review of this repository (the "trial" generation) and the user's
own account of what VBA gave them for twenty years. This is not a plan; it is the set of
decisions and lessons that PT2 starts from, so that the next generation is designed with them
in hand rather than rediscovered. Names, folders and stack versions are placeholders.

## Why there is a PT2

This repo proved the architecture — a headless Core that owns every decision, Excel as the
GUI, state in sidecars and source files, byte-faithful patching instead of Excel saving — and
it proved it on real plants (FVZ, FVT, Jadlog). It also accumulated the debt of being the
trial: `net48` + Excel-DNA in-process, `dynamic` COM calls across the add-in project, WinForms
living inside Excel's process, a portable launcher that never reached 1:1 with the add-in.
PT2 keeps the architecture and drops the hosting model.

## What VBA did right — the properties PT2 must preserve

These are the reasons many never left VBA, restated as requirements rather than nostalgia.

| VBA property | What it bought | PT2 requirement |
|---|---|---|
| **Excel is a cooperator, not a target.** One line of an Excel call did what a page of code would. | Formulas as a compute engine; ranges as data structures; sorting/filtering/pivoting for free. | The host hands Excel the parts Excel is good at (display, filter, formulas, recalc) and never re-implements them. |
| **No GUI to build.** Textboxes, buttons anywhere, a form in minutes. | The sheet *is* the form; every intermediate result visible at every change. | Sheets are the primary UI (already the repo rule: blue borders, one column per action, status columns). Forms only for things a sheet cannot ask. |
| **Always runs in debug.** An exception lands you on the line; edit it; continue. | Shortest possible loop from "it broke" to "I see why" to "it works". | Host runs out-of-process with a debugger attachable at any time; every failure names the row/file/cell it was handling; structured `_Log` sheet with jump links in every workspace. |
| **Reach anything.** .NET, COM, a `.bat`, a `.vbs`, run from a cell. | No wall between the sheet and the machine. | The host is an ordinary .NET process: shell out, call anything, read any file, no sandbox. Excel is one of its instruments, not its cage. |
| **The macro is in the file.** Send the workbook, the tool travels with it. | Distribution was free. | Never fully recoverable in a signed-add-in world — see *Distribution*. The nearest substitute: one `.exe`, no install, no Trust Center, Excel launched by it. |

## The architecture decision (agreed)

**Excel is an instrument.** PT2 is an out-of-process **.NET 8+ host** — console/tray/service,
one executable — that does all work on files and drives Excel over COM only to render and to
collect gestures. Nothing of PT2 runs inside Excel's process.

- **Core stays headless** (the netstandard Core of this repo is the seed: join, diff, patchers,
  sidecars, mnemonic tables) and gains nothing that needs Excel. Every decision is testable
  without Office installed.
- **The host launches controlled Excel sessions**, one per purpose: an IO Bridge session, an
  SWP Bridge session, a CSV workspace — each a dedicated `EXCEL.EXE` the host owns (started,
  configured, shut down by the host; never the user's open Excel), with only that tool's
  workbooks, pane and gestures inside it. Crash isolation and clean teardown come free; the
  "closing the only workbook takes Excel down" class of bug disappears because the host owns
  the lifetime.
- **Gestures travel back over COM events** (sheet double-click, cell change, workbook close) to
  the host, which recomputes and re-renders. The thin event shim is the only in-Excel code and
  it is generic (it knows nothing about bridges).
- **No ribbon.** The Tools Hub pane pattern (a docked task pane fed by the host) replaces it;
  the pane is a renderer like the sheets.
- **Modern .NET** for the host: current C#, current BCL, `System.Text.Json`, Spans, channels —
  the things `net48` denied. The COM interop uses the typed PIA where possible; `dynamic` is
  confined to one adapter file per Excel object kind, never spread through features.
- **Forms out of process.** Any dialog PT2 still needs is a window of the host (WPF/WinUI/
  Avalonia — pick once), owned to the Excel HWND for modality. The GDI+ jagged-text,
  STA-dashboard and silent-MessageBox workarounds of PT1 all cease to exist.

## What to keep from PT1, verbatim in spirit

- The repo rules in `CLAUDE.md`: FLD above everything, mnemonics are never rewritten (rules as
  data, a new shape is a row), formulas spelled for the file, plant CSVs recognized by content,
  the UI conventions (blue double-click border, one column per action, red attention outlines,
  pending yellow / applied blue / header blue / SWP orange, tab colours, never close a generated
  workbook, overrides name the opposite of the default, radio not two checkboxes, labels say
  what lands in the file, never ask for a number the file states, a session is a file).
- Byte-faithful file writes with three-way checks, atomic write + `.bak`, re-read before write.
- Sidecar-as-state, workbook-as-view; `Saved = true` on close; AutoFilter captured across
  re-renders.
- Docs as part of every change (`add-in-reference.md` per behaviour).
- The field-first verification habit: a feature is done when a real Excel opened a real file
  the patcher wrote.

## What to drop

- Excel-DNA and the packed `.xll` (and with it the pack-lock, bitness, mark-of-the-web,
  Trust Center).
- `net48` and in-process WinForms.
- The portable launcher as a product: its job (run without installing) is the host's default.
- `dynamic` outside the COM adapter layer.
- Registry-persisted settings as the first choice: PT1's HKCU keys exist for compatibility;
  PT2 keeps settings in files beside the data they describe (already the direction of
  `swp-bridge.json`).

## Distribution (the honest part)

The one VBA property that does not survive. Realistic target: a single self-contained `.exe`
(`dotnet publish -r win-x64 --self-contained`), no admin, no install, no Trust Center prompt
because nothing loads into Excel; the Excel it launches is whatever Office the machine has.
Colleagues get one file. The costs to accept: ~70 MB binary; SmartScreen on first run (signing
certificate or a known-publisher route is the real fix); Excel COM registration must exist on
the machine (it does wherever Office is installed).

## The spreadsheet-editor option (parked, not rejected)

Writing PT2's own grid editor — Python or .NET — with "Excel as a GUI" as an *option* rather
than a dependency was raised. The attraction: no Excel dependency for the headless 80%, full
control of rendering, cross-platform. The cost: giving up exactly the cooperator properties in
the table above (formulas, filters, pivots, recalculation, the user's own Excel habits), which
are the reason the sheets-as-UI rule works at all. Decision for now: PT2 targets Excel; the
Core/host split keeps the door open — a second renderer would attach where the COM adapter
sits, and nothing in Core would know.

## Open questions for the PT2 design phase

1. One host process for all tools, or one per tool? (Lean: one host, many Excel sessions.)
2. Event shim: VBA in a hidden add-in workbook the host injects, vs. COM event sinks from the
   host alone? Measure latency and reliability of `SheetBeforeDoubleClick` across the boundary.
3. Pane technology: a WebView2 pane fed by the host, or a native window docked beside Excel?
4. How much of PT1 Core ports unchanged to modern .NET (expected: almost all; the OpenXml SDK
   and the CSV engine are the two dependencies to re-verify).
5. Governance: start PT2 with a living contract from line one (intent-level clauses, evidence
   files, adversarial refute before "verified"), with a hard altitude cap learned from OP5.
6. Name.

## See also

- `docs/portable-host-plan.md` — the PT1-era tier-2 plan whose seam inventory (ExcelApp.App,
  ExcelStateGuard, the QueueAsMacro sites, AddIn.cs event wiring) is the port list for the host.
- `docs/io-bridge-port-plan.md` — IO Bridge features to bring to SWP Bridge parity; everything
  there that is Core-side carries straight into PT2.
- `CLAUDE.md` — the rules PT2 inherits.
