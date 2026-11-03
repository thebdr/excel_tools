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
place is never tolerable. Two corollaries every write path follows: **an empty IO-side value never
erases what the SWP holds** (a missing FU / Location / Device / description / T.S. ref is missing
data, not a blank to write — G, J, L, T/U and the H/I overwrite all guard it), and in append mode a
row's **device blank travels only with the label writes it depends on** (the `FLD` group: toggled
together on `_Queue`, applied together or not at all). See `SwpDiff.ProposeEdits`, the "FLD rule"
paragraph in `docs/add-in-reference.md`, and the full statement of priorities in
`docs/diagnostic-priorities.md`.

## Mnemonics are never rewritten (SWP Bridge)

A mnemonic's **destination is data, not a rename**. The legacy tool mutated names so they would
ride an existing branch (`EMERGENCY02000` → `CABINETAL02000`, which worked only because both
prefixes are nine characters); that cost the identity every plant CSV knows the signal by, and it
could only express shapes that happened to line up with one already handled. `SwpMnemonicMap` is
the single table — one rule set per plant type, regex → sheet/tag/word/index/bit, `_Offsets` still
applied by the rule's object key — and it decides BOTH whether a mnemonic is a diagnostic and where
it lands. Consequences that are load-bearing: a new shape is **a row, never a branch**; an empty
destination means *recognized and deliberately unwritten* (the `EMERGENCY_ON` commands, AND/OR
group members after the `_0` one); an index declared as `map(Source,…)` **fails when that source is
absent or incoherent** rather than falling back to the mnemonic's own number; and a rule that does
not parse is inert-with-a-warning. `SwpAddressMap` is superseded — it survives only as the parity
reference in `SwpMnemonicMapParityTests`, and adding a shape there instead of to the table is a bug.

## Formulas are spelled for the FILE, not for the formula bar (SWP Bridge)

A formula the patcher writes into the XLSX XML uses the file-format spelling, which is NOT what
Excel displays: every function added after Excel 2007 carries `_xlfn.` (`_xlfn.CONCAT`,
`_xlfn.LET`), and every LET/LAMBDA parameter name carries `_xlpm.` (`_xlfn.LET(_xlpm.t,…)`, shown
back as `=LET(t,…)`). Miss the first and the cell reads `_xludf.<name>` → `#NAME?`; miss the second
and Excel's loader REMOVES the formula on open ("Removed Records: Formula from …"), keeping only
the cached value. Both reached the field. So a new formula shape is not done until a real Excel has
opened a file the patcher wrote and read the formula back (`HasFormula`, `.Formula`, the
recalculated `.Text`); a unit test that the string looks right proved nothing, twice.

## Plant-CSV recognition (IO Bridge + IO-list pulls)

Plant CSVs are recognized **by content, never by header text alone** (`CsvColumnsMap.Locate`):
column names when any row resolves them — row widths free, because the ISPC export's data rows
carry a trailing `;#;` marker column its header line lacks — and the file kind's legacy fixed
positions otherwise, gated on the data being signal-shaped (text names, numeric addresses). A
header check must never block a load; only non-signal data may (user ruling 2026-08-24).

The SWP Bridge **busfield** follows the same rule (`SwpCoherence.BusfieldColumns`): its columns moved
between IODriver versions — the older export has no Linked Master/Slave ID, so Reference Index /
Element / Bit sit two fields earlier — and fixed positions misread every older file (Element read the
`#` separator, Bit the IP address). Names first, current positions only when the header names nothing.
Its **Reference Index is compared only on rows that carry a linkage** (index, element or bit non-zero):
every modern export on file leaves all three 0 — ISPC, MSC and upstream alike — and comparing the 0
reported every modern ISPC/upstream module as a missing node.

**IO-list card sheets are found by name first, by content only when the names name nothing**
(`IoListCardSheets.Find`): the `R{rack}.S{slot}.C{card}` names when the workbook has any, else the
sheets whose row-1 header resolves Mnemonic + Bit — the pair the card reader itself requires, so a
listed sheet is one the reader will read. The safety PLC's list names its one card sheet by network
(`NET SAFETY 192.168.20`); a sorter list with real `R.S.C` sheets must never grow an extra ticked
sheet because a summary tab carries a Mnemonic column.

## SAFETY is the sorter's cabinet family and nothing else (SWP Bridge)

Plant type 4 (SAFETY, the Siemens safety PLC on its own OPC UA link) is MSC logic restricted to
cabinets: its mnemonic rules are exactly the sorter's CABINET + EMERGENCY rows (the safety tooling
writes sorter syntax on purpose), never the CLASSDIAG classes — a safety SWP is the cabinet sheet
alone (rules §9.2), so a CLASSDIAG mnemonic in a safety list is unclaimed, not routed to a sheet
that does not exist. Its one production input is the SCADA's `Alarms_*.txt` export (`Alarms List\`
beside the SWP, `AlarmLists\`, or flat): no busfield, no Plant_Inputs, and neither is listed.
Its signal rows carry Siemens absolute addresses with no node (`I.0.4` / `I0.4` / `%I12.7` —
`IoListLogic.TryParsePlcAddress`), which SAFETY alone accepts and writes as the TIA form `I0.4`;
on every other plant type a three-part address stays the unreadable-address error, because on a
sorter list it is a missing bit (field gap 2026-09-22: 454 rows in the list, 70 in the workspace).
`_Offsets` keeps SAFETY 1:1 zero-based — the real FVT export starts at `S1.CABINETAL00000`; the
20-item file at the FVT project root starting at `00100` is a generic placeholder, not evidence.

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
  later can never silently reverse a row's recorded decision. Cleared = follow the default. A
  gesture that does nothing on the row you ticked is a bug.
- **Two answers to ONE question are radio buttons, never two checkboxes.** A row cannot both keep
  CONFIG_PLANT's template whole and cut its tail off, so those are one `SwpConfigMode`, not two
  bools — the checkbox pair also admitted a third state (compose the description AND leave the
  formula alone) that was the stacking shape the mode exists to end. A mode that replaces
  independent switches keeps writing them in the sidecar, derived from itself, so an older build
  reading a newer file finds the setting it knows; migration is most-specific-first and any
  behaviour it changes goes into `LoadNotes`, never silently.
- **Keeping CONFIG_PLANT's text keeps its WORDS — it never starves the cell a kept formula shows.**
  A formula already displaying `T`/`U` after a template with nothing past its last placeholder
  (CABINET: `…%2 - ` then `&" "&U{row}`) has no sentence to stack onto; that cell is the alarm's
  only name and takes the description under either mode (`SwpDescriptionFormula.ShowsSuffixAfterBareTemplate`,
  read from the displayed text — unreadable means label-only). Withholding it left a French-only
  SCADA reading stale English (FVZ, 2026-09-10).
- **Description 2 is all a single-language SCADA shows.** With a second language other than English configured, any
  language-1 text landing there — the whole-description fallback, or a language-1 variable part
  beside a glossary translation — is a `_Log` warning, never silent; and a glossary match takes the
  IO list's language-2 variable part whenever the IO list WROTE one (`SwpDescriptions.Compose`):
  a part the dictionary empties still counts as written and nothing rides in its place — language
  1's raw part riding back in was the one path with no dictionary (FVT `CH1/CH2`, 2026-09-22).
  **Text landing in Description 2 always passes the language-2 dictionary last**, language 1's
  riding part and the whole-description fallback included; only a bus module's phrase +
  `[name - ip]` is identity that no dictionary touches.
- **A workspace default may carry ONE regex exception on the destination sheet name, and the
  levels resolve most-specific-first**: per-row tick → sheet rule → default (`SwpFldSheetRule`,
  `SwpConfigSheetRule` — both over `SwpSheetNameRule`, resolved in `SwpDiff.ProposeEdits`). The
  mode is the WHOLE answer: no invisible scope on top of it, because what departs from the
  default must be something a person can see. One pattern suffices for a binary setting — it names the
  sheets taking the OTHER mode — and it obeys the same rules as a per-row override: the caption
  names the mode it applies and flips with the default, the answer is stored as that mode and
  never as a flip. Because the rule moves rows with no visible gesture ON them, two things are
  mandatory: the tick on a governed sheet flips away from the RULE's mode (not the default, or
  the tick does nothing), and each build reports how many rows the pattern moved. A pattern that
  does not parse is refused at setup and inert-with-a-warning if it reaches a build — never fatal.
- **A setting's label says what lands in the file**, not what the code calls it: "Never overwrite
  CONFIG_PLANT descriptions", not "Trust CONFIG_PLANT". Field and sidecar keys stay put
  (serialized schema); only the human text changes. A key is added, never renamed, when the
  setting's SHAPE changes.
- **A dialog section is one QUESTION, not one control type**: SWP SYSTEM, IO LIST (its workbook,
  its cards, and the regexes deciding which of its mnemonics count), PRODUCTION FILES,
  INTERNATIONALIZATION (the glossaries, the second language, and how the glossary competes with
  the IO list), CONFIG_PLANT, ELECTRICAL_DESIGNATION. A setting lives with the thing it is about.
- **Never ask for a number the file already states.** The first-line and group ids come from the
  SWP's own tab names (`IL02_…`, `ILGRP02_CABINET` — `SwpPlantNumbering`), because the mnemonic
  table builds destination sheets from exactly those numbers: retyping them is re-deriving what
  the file says, and one wrong digit sends every row to a sheet that does not exist. Derived
  values stay editable and say where they came from.
- **A session is a FILE, not a folder convention.** `swp-bridge.json` at the SWP root is the
  default, not the only option: "Save session as…" names one anywhere, which is what lets one IO
  list drive two SWP systems as named sessions kept in one folder. Save as WRITES the file at once
  (settings + decisions + dictionaries + queue) and SWITCHES to it — the dialog, the running
  workspace it had open (rewritten with its running settings if the dialog is cancelled), and what
  Setup reopens next (running session → last session → last root). Picking an SWP never swaps a
  named session for a root's default sidecar. A session saved at its own root records NO root and
  stays self-locating (a system folder moved wholesale still resolves); only one saved elsewhere
  records which system it belongs to (`SwpBridgeSidecarStore.RootOf`). "Save as…" carries the
  session's history with it — that is its job;
  **"Save as new…" carries everything that is not history** (settings, file references, dictionaries,
  glossary corrections, strobe templates) and starts the queue and the per-row decisions empty
  (`SwpSidecarJson.FreshSession`), because a second system driven by the same IO list shares its
  mnemonics and would silently inherit the first one's decisions. A running workspace moves with a
  Save as, never with a Save as new. Reopened after another
  session's apply made a newer `NN_…` folder holding its SWP, a session follows the newest one
  (`SwpVersionFolders`) and says so.
- **File-dialog filters separate patterns with `;`** (`*.xlsx;*.xlsm`) — the Windows syntax; `,`
  matches nothing (checked against the shell's own matcher, `PathMatchSpecEx`). A filter naming an
  exact file hides every other name: the session filter listed only `swp-bridge.json`, so named
  sessions never showed up.
- Generated workbooks **never trigger Excel's own save prompt**: mark `Saved = true` in the
  `WorkbookBeforeClose` handler. Only unsaved changes to SOURCE FILES get a prompt, and it is the
  add-in's own (PT_CSV: Save All / close without saving / cancel).
- Sheets that mirror external state (the PT_CSV Index) rescan on sheet activation AND on window
  focus — the user must never have to switch sheets to see an "edited externally" flag.
- **A re-render never discards the user's view state**: a sheet that refills itself (an override
  edit recomputing the SWP spine) captures the AutoFilter criteria first and re-applies them after
  (`SwpBridgeSpineSheet.CaptureAutoFilter`/`RestoreAutoFilter`; colour/icon filters are the COM
  limit). Requiring a manual refresh instead is not an option — a stale view desyncs the queue
  from the recorded decisions, and Apply would write the old ones.
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
