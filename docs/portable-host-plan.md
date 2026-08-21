# Portable host plan — the add-in as a standalone exe

Two tiers serve two different blockers. **Tier 1 is built** (`ExcelPowerTools.Portable.exe`,
the session launcher — see the add-in reference). This document is the plan for **tier 2**: a
pure COM-automation exe that drives Excel as its GUI with **no .xll involved at all**, for
machines whose policy blocks unsigned add-in files outright.

## Why the codebase is ready for this

- All logic lives in `ExcelPowerTools.Core` (netstandard2.0, zero Excel dependency): OpenXml
  read/patch, filter, diff, queue, sidecar, glossary, alarm strings, index rules. Untouched.
- The net48 layer talks to Excel through **late-bound `dynamic` COM** (`ExcelApp.App`) — the
  same code runs against an `Application` object from any origin.
- Event capture already uses **typed interop events** (`Microsoft.Office.Interop.Excel`
  `Application` + `+=` in `AddIn.cs`) — PIA event sinks work identically from an out-of-process
  automation client. The whole edit-capture layer (bridges, manual tracking) ports as-is.
- The renderers batch by design (`object[,]` writes, run-grouped paints, bounded AutoFit), so
  the higher per-call cost of cross-process COM stays acceptable.
- The SWP dashboard already runs on its own STA thread and never touches COM directly
  (everything dispatches) — the exe generalizes that pattern instead of inventing one.

## The actual seams (measured, 2026-08-21)

79 `ExcelDna.*` references across 31 files, but almost all inert or trivially fenced:

| Seam | Where | Port action |
|---|---|---|
| `ExcelApp.App => ExcelDnaUtil.Application` | `Interop/ExcelApp.cs` (1 line) | Becomes a settable host: the exe assigns the automation `Application`; the add-in head keeps Excel-DNA. |
| `ExcelDnaUtil.Application` in `ExcelStateGuard` ctor | same file (1 line) | Read from `ExcelApp.App` instead. |
| `ExcelAsyncUtil.QueueAsMacro` | 12 call sites / 7 files (dashboard 2, pane 4, timeline 1, CSV 5) | Replace with `Host.Post(Action)`: add-in head = QueueAsMacro; exe head = invoke on the Excel-owning STA thread with an `IOleMessageFilter` retry (Excel rejects calls mid-edit/modal → `RPC_E_SERVERCALL_RETRYLATER`). |
| `[ExcelCommand]`/`[ExcelFunction]` attributes | all Commands files | **Inert** — the exe references `ExcelDna.Integration.dll` as a plain library (zlib license) so the attributes resolve; nothing initializes it. Zero source edits. |
| `IExcelAddIn` bootstrap | `AddIn.cs` | Exe gets its own `Program`: launch/attach Excel, assign `ExcelApp.App`, run the SAME event-wiring body (typed PIA sink), message loop. |
| Ribbon (`ExcelRibbon`) + CTP pane | `RibbonController.cs`, `UI/PowerToolsPane*` | Add-in-only. The exe's launcher window lists the commands instead (the pane's section model reused as a WinForms list). |
| Home-sheet buttons (`OnAction = "PT_*"`) | bridge aux sheets | No macros exist without an add-in. Host capability flag: out-of-proc renders styled **hyperlink cells** and handles `SheetFollowHyperlink` (typed event exists) → same command dispatch. |

## Exe-side infrastructure (the genuinely new code)

1. **Instance lifecycle** — launch a dedicated `Excel.Application` (`Visible = true`,
   `UserControl = true`); never hijack the user's open instance (the bridge workspaces are
   ephemeral workbooks anyway). Detect death (COM call throws `RPC_E_DISCONNECTED`) → offer
   relaunch; on exe exit leave Excel to the user. Zombie hygiene: `FinalReleaseComObject`
   discipline on exit paths. Measured 2026-08-21: `Quit()` on a hidden automation instance
   that had loaded the add-in did NOT terminate `EXCEL.EXE` (it kept the packed .xll locked) —
   programmatic teardown must expect a surviving process and escalate (bounded wait → kill by
   the instance's own PID via `Hwnd`), never assume Quit sufficed.
2. **`IOleMessageFilter`** — standard `CoRegisterMessageFilter` retry filter on the Excel STA
   thread; converts busy-Excel rejections into bounded retries instead of exceptions.
3. **Threading** — one STA "Excel thread" owns every COM call and the PIA event sink (events
   arrive marshaled to it); WinForms launcher + dashboards stay on their own threads exactly as
   today; headless phases (OpenXml, diff) on worker threads as today. `Host.Post` targets the
   Excel thread.
4. **Command surface** — launcher window (sections from the pane model) + the existing
   dashboards; Home sheets get hyperlink cells.

## Scope and phases

Phase A ports the **IO Pipeline family only** (pulls, IO Bridge, SWP Bridge) — the audience
that can't install add-ins is exactly the audience of these workflows. The CSV pipeline,
comparison review, undo timeline and pane stay add-in-only until asked for (they lean on
UDFs/CTP/GetPivotData-style surfaces that don't translate 1:1).

- **A1 — host abstraction**: settable `ExcelApp.App`, `Host.Post` seam, fence `AddIn.cs` /
  ribbon / pane out of the exe compile (shared-source two-head layout: the exe csproj
  `Compile Include`s the io-pipeline + interop sources it needs, like Core is linked today).
  Add-in behavior provably unchanged (full suite + field smoke).
- **A2 — automation host**: lifecycle + message filter + event sink + hyperlink-buttons + the
  launcher window. Exit criterion: SWP Bridge full cycle (open → edit capture → preview →
  apply → refresh) from the exe on the FVZ files, add-in not installed.
- **A3 — hardening**: Excel-death recovery, second-launch behavior (attach-or-new prompt),
  SmartScreen story (sign or share-deploy), docs.

Estimated a few focused days; the risk is concentrated in A2's threading/lifecycle, which the
dashboard precedent de-risks.
