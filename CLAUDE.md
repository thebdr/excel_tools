# Excel Power Tools — repo rules

Excel is the GUI of this add-in: generated workbooks are views, the sidecars and the source files
are the state. These rules keep every generated sheet readable the same way.

## The FLD rule (SWP Bridge — above everything else)

The IO list's **Functional unit / Location / Device** (FU = a system inside a building, Location =
a system inside a FU, Device = the one labeled hardware piece) are the most important information
the SWP carries: the maintainer has SCADA and the T.S. manual, nothing else, and a missing FLD
turns a repair into a service call that can take weeks. Every write path must leave the FLD
visible — in the row's own cells (H / I / J, brutal overwrite) or appended to the descriptions
(append mode) — and any change that can drop it is a bug before it is anything else. Priority
order: a wrong description is tolerable when the device pointed at is the right one (once a
person stands in front of it, LEDs and tripped switches tell the rest); a wrong or missing
place is never tolerable. See `SwpDiff.ProposeEdits`, the "FLD rule" paragraph in
`docs/add-in-reference.md`, and the full statement of priorities in `docs/diagnostic-priorities.md`.

## UI conventions (Excel as the GUI)

- **Double-clickable cells carry a thin blue border, RGB(0,112,192)** (`xlContinuous`, `xlThin`)
  on EVERY sheet the add-in generates — the cells a double-click acts on, and only those. The
  border is the affordance; a gesture without the border, or a border without the gesture, is a
  bug. Helpers: `SwpBridgeSpineSheet.MarkDoubleClickable`, `IoBridgeBuilder.OutlineRows`, the
  PT_CSV Index "Status" column in `CsvWorkspace.StyleIndex` (one gesture cell per row — the
  status is where a state-related action belongs).
- **One column per action**: a double-click action lives on exactly one column of a sheet. The SWP
  Workspace's Verdict AND Coordinate both opening `_Preview` predates this rule and stays; never
  add a second column for an existing action, and never put two actions on one column.
- **Attention outlines are red, RGB(200,30,30)** — a medium `BorderAround` on the cells that need
  a decision (an IO Bridge virtual twin's CSV block); red interior RGB(255,199,206) marks conflicts
  and mandatory data that is missing (a T.S. ref on an IO-list alarm).
- Pending = yellow interior RGB(255,235,156); applied externally = blue interior RGB(221,235,247);
  header rows RGB(180,210,240); the SWP/CSV side of a spine in orange RGB(250,200,140).
- **Sheet tabs reuse the sheet's own palette** so a tab is found by colour, not by reading the
  strip: SWP Bridge Home light blue RGB(180,210,240), Workspace orange RGB(250,200,140), `_Preview`
  yellow RGB(255,235,156) (`SwpBridgeBuilder.ColorWorkspaceTabs`). A sheet that unfolds another
  sits **next to it** in the strip, never at the end where `EnsureSheet` appends it.
- **Reconfiguring never closes a generated workbook**: a settings change rebuilds INTO the live
  workbook (`SwpBridgeCommands.OpenWorkspace` reuse path + `SwpBridgeSession.RetireForRebuild`).
  Closing a bridge's only workbook takes Excel down with it.
- **A per-row override of a workspace default names the OPPOSITE of that default** and is stored
  as the mode it chose, never as a "flip" bool: the header says what ticking does
  (`[Overwrite FLD]` under an append default, `[Append FLD]` under an overwrite one —
  `SwpBridgeSpineSheet.FldModeHeader`; likewise `TrustConfigHeader`), and changing the default
  later can never silently reverse a row's recorded decision. Cleared = follow the default. An
  explicit per-row tick also bypasses any automatic scope the workspace setting applies — a
  gesture that does nothing on the row you ticked is a bug.
- **A setting's label says what lands in the file**, not what the code calls it: "Never overwrite
  CONFIG_PLANT descriptions", not "Trust CONFIG_PLANT". Field and sidecar keys stay put
  (`TrustConfigPlant` / `trustConfigPlant` — serialized schema); only the human text changes.
- Generated workbooks **never trigger Excel's own save prompt**: mark `Saved = true` in the
  `WorkbookBeforeClose` handler. Only unsaved changes to SOURCE FILES get a prompt, and it is the
  add-in's own (PT_CSV: Save All / close without saving / cancel).
- Sheets that mirror external state (the PT_CSV Index) rescan on sheet activation AND on window
  focus — the user must never have to switch sheets to see an "edited externally" flag.
- Windows of generated workbooks carry `PT_*` captions (PT_CSV, PT_IOBridge, PT_SWPBridge).
- Column widths autofit to the largest cell (no caps); on a sheet with an AutoFilter, allow for
  the drop-down button that covers the header's right end (caption chars × 1.15 + 3 width units),
  or the last glyph of a header-widest column disappears.
- WinForms text renders through GDI (`UI/GdiText.cs`); inside Excel the GDI+ default comes out
  jagged. New forms call `GdiText.Apply(this)` after building their controls.
- **Never call `MessageBox.Show`** in the add-in: it always plays a system sound. Use
  `Interop.Dialogs` (`Info`, `Confirm`, `YesNoCancel`, `OkCancel`, `RetryCancel`, `Message`) — the
  add-in's own silent modal box, owned by the invoking window (default: the Excel window).

## Build, test, pack

- `dotnet test tests/ExcelPowerTools.Tests` must stay green (Core is netstandard2.0, tests net8).
- The packed add-in is `src/ExcelPowerTools/bin/Release/net48/publish/ExcelPowerTools-AddIn64-packed.xll`;
  the pack step fails while an Excel holds it — wait for the lock, then build Release again.

## Docs

- `docs/add-in-reference.md` is the user-facing reference: every behavior change lands there in
  the same change set.
