# Excel Power Tools

A rebuild of the personal VBA macro collection (formerly in `PERSONAL.XLSB`) as a proper
**Excel-DNA add-in** written in C#. It loads as a single `.xll` you enable once in Excel —
no more flaky `PERSONAL.XLSB` startup — and moves the code off VBA onto .NET.

## What maps to what

| Original VBA | New home | Notes |
|---|---|---|
| `IoListAux.ChangeField` + `ufIoListAuxParameters` | `Commands/FieldEditorCommands.cs` + `UI/FieldEditorForm.cs` | Core split/rebuild logic lives in `ExcelPowerTools.Core/FieldStringEditor.cs` (unit-tested). Per-cell error `MsgBox` is now a single end-of-run summary. Each separator box has quick-pick buttons to its right — the usual `( ) . , - _ ;` plus whitespace tokens `\s` (space), `\n`, `\r`, `\r\n`, and Clear; each colour box has 6 colour swatches. |
| `PasteMultiSelectionValues.KeepValuesOnlyV2` (Ctrl+Q) | `Commands/SelectionCommands.cs` | Same Ctrl+Q shortcut, registered by the add-in. |
| `csvImportExport.*` | `Commands/CsvCommands.cs` + `Commands/CsvWorkspace.cs` | Multi-file CSV/TSV/TXT workspace, folder open, generated Index, per-sheet safe save/Save All, export and backup. |
| `ThisWorkbook` app events | `AddIn.cs` + `CsvCommands.OnWorkbook*` | Native delimited-file interception, edit tracking, Index refresh, external-change protection and close warning. |
| `aPlayground.exportSheetsAsNewFile` | `Commands/WorkbookCommands.cs` | Generalised: works on the active workbook and asks how many sheets per file (the VBA had a hard-coded path). |
| `xlReferences` / `xlCells` helpers | `Interop/RangeHelpers.cs` | `LastRow`, `LastColumn`, `LastRowInColumn`, `WorksheetExists`. |
| `xlFormulas.vbaIndexMultiValues` (UDF) | `Functions/WorksheetFunctions.cs` | Worksheet function `=PT.INDEXMULTI(range, positions)`. |

The dialog's saved parameters used to live on a hidden `DataStorage` sheet; an add-in has no
workbook, so they now persist per-user in the registry under
`HKCU\Software\ExcelPowerTools\FieldEditor`.

## User interface

**Comparison rebuild, 10 September 2026:** a fresh implementation is in progress under the
[approved comparison plan](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md).
The new saved-file reader and alignment core are currently tested through the repository's test suite,
including a read-only trial on five real sorter revisions. There is no new launchable review workflow yet.
The comparison ribbon/Hub descriptions below refer to the pre-existing implementation, not the rebuild.
See [the build checkpoint](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-build-2026-09-10.md)
for current support boundaries and validation evidence.

The **Power Tools** ribbon is a compact command strip, grouped by workflow rather than by implementation:

| Ribbon group | Direct actions |
|---|---|
| Power Tools | Tools Hub, Undo, Redo, Timeline |
| Compare & review | Compare revisions, Keyed table comparison, Apply selected changes, Exact two-file preview, Open Review workbook |
| Delimited workspace | Open data (files/folder split button), Save file, Save all, Index |
| Transform | Keep values, Smart fill, Change field, Regex replace |
| Clean & analyze | Natural sort (ascending/descending split button), Deduplicate, Column stats |

The large coloured icons identify the Hub and the three primary workflows; secondary actions keep familiar
Office icons. Save controls are disabled until their required source sheet or workspace is active, and
Undo/Redo reflect the Power Tools timeline.

**Tools Hub** is a docked, searchable task pane with category filters for Compare, Files, Transform, Data,
and Utilities. It exposes the complete catalog—including the less frequent file operations, workbook split,
the IO Pipeline tools, developer shortcuts, About, and persistent behavior switches—without allowing
future modules to overwhelm the ribbon. A separate pane is retained per Excel window. Existing transactional
editors and conflict prompts still open as dialogs from the Hub; they are intentionally not embedded as Form
objects in the task pane. The implemented **Comparison** workspace can select and integrity-validate an existing
`Review.xlsx`, display a validation summary for its project and embedded package state, and launch a protected
read-only copy. A validated decision sidecar adds only safe counts and shortened catalog/generation IDs to that
summary. Beneath it, an editable virtual ledger shows bounded master/candidate content and locations from the
validated catalog. It supports filtering, multi-selection and right-click actions for exact observation choices,
review statuses, and whole groups. Reviewer identity and an audit note are captured with the queued immutable
events. **Save new Review revision** writes the ordered pending batch only to a newly named workbook; the source
stays unchanged and its exact hydrated package/occurrence is rechecked on the protected derivation handle. Use
**Compare revisions → Review workbook in Hub...** on the ribbon, or **Open Review workbook** in the Hub, to enter
this workflow directly.

## Keyed table revision comparison

**Keyed table comparison...** is a secondary comparison-menu command. It accepts 2–20 distinct local
`.xlsx`/`.xlsm` files from one or more folders, in an explicit order, and compares every adjacent step
(`R1 -> R2`, `R2 -> R3`, and so on). Tables with the same sheet name are proposed automatically; other sheets
can be paired explicitly. Paired key headers may have different names when mapped explicitly; every other
populated column must retain the same unique, nonblank header in this first version. Added, removed, renamed,
or blank-headed data columns are reported as a blocking setup error rather than silently ignored. The
header row and one or more semicolon-separated key columns are always visible and every proposal must be
confirmed. Key preview reports row counts, blank keys, and duplicate-key rows before the report is built.

The macro-free review workbook has exactly five visible sheets: **Summary**, **Added**, **Removed**,
**Changed**, and **Duplicate Keys**. Values retain their exact stored type, formulas are compared by exact
stored formula text, and keyed row movement does not lose identity. Blank and duplicate keys remain visible but
are excluded from automatic matching. Formula-valued keys and formula coverage the adapter cannot prove are
blocked rather than guessed. Removed rows are review-only in this first working version.

**Added** and **Changed** include an **Apply** column. After marking actions `Yes` and saving the review, use
**Apply selected comparison changes...**. Apply re-reads every source and its raw-byte SHA-256, reconstructs the
comparison, reanchors each action by its explicit key, and performs a three-way check. Selected cell changes and
whole added rows are written to a new copy of the first revision; neither inputs nor the review are overwritten,
and apply refuses any output filename that already exists.
Added rows append after the last content row and formulas are translated to the new physical row. Ordinary
single-cell formulas are supported; grouped, array, dynamic-array, data-table, or untranslatable formulas stop
the operation. The private output stage is reopened and every written cell is verified before publication;
`.xlsm` VBA bytes must remain identical.

The keyed path has a 768 MiB aggregate input-package limit in addition to the OOXML adapter's per-file limit.
The primary revision-alignment project and positional exact preview are separate comparison commands.

## Multi-revision comparison (Alpha)

**Compare revisions** is the primary ribbon action and is also available in the Hub with an **ALPHA** badge.
It accepts 2–20 distinct local `.xlsx`/`.xlsm` revision files. Inventorying hashes the
read-only sources, proposes a filename/evidence-based order and parent graph, and then opens a setup window in
which the user can reorder revisions, edit labels and parent links, choose the master and optional trusted
baseline, and select a new `Review.xlsx` destination. The displayed order, parent links, master, and baseline
must be explicitly confirmed before the comparison starts.

The engine conservatively infers sheet identity (including supported rename evidence), headers, keys, and
logical row/column identity across adjacent lineage edges. It builds multi-revision history of stored values,
saved calculated values, and formulas, and groups related changes only when identity is sufficiently supported.
Ambiguities and unsupported coverage do
not become guessed matches: they produce stable, verbose findings and corrective actions on the generated
workbook's **Diagnostics** sheet. A blocked project therefore still yields a useful diagnostic Review workbook,
but no logical history is claimed for the unresolved branch.

Cell-content comparison prioritizes stored values, saved calculated values, and formulas. Saved calculated
values are compared even when the formula also changed. Sheet, row, and column structure is retained only to
align those channels. Styles, fills, comments, drawings, and package metadata are not comparison changes; they
are used only for safe reading and integrity checks. Blank-versus-missing storage is retained only for alignment
safety.

Inserted and deleted rows and columns are aligned across consecutive revisions. Repeated exact rows pair in
monotonic sequence and surplus occurrences become explicit insertions or deletions; a competing reorder remains
blocked. Tables identical in stored values, saved calculated values, and formulas use stable positional identity,
so duplicate rows or duplicate headers
do not make an unchanged table ambiguous. A blank-headed column may align when its neighboring position is
unique and at least 90% of its comparable cells agree; competing or weak candidates stay unresolved. For
dynamic-array cells, stored values, saved calculated values, and exact formulas remain available for Review with
a warning. Spill metadata is retained only for safety classification, and merge stays blocked until spill
semantics can be validated.

The alignment-correction workflow is implemented end to end. Before publication, an incomplete analysis—or a
complete analysis with active confirmed overrides—opens **Comparison Project — alignment review**. Its
structured editors can confirm or replace an exact sheet pair; set each side to Auto, No header, or a physical
header row with an optional later first-data row; and set one-to-four ordered Before → After key-column pairs.
An explicit choice can be cleared back to automatic inference. Confirm, replace, and clear actions are
immutable, append-only audit events tied to exact revision and physical-sheet identities. **Rerun analysis**
validates the queued generation and rebuilds logical identities from the command's retained immutable parsed
snapshots—it does not reread or modify the sources. A complete result can continue to Review; an incomplete one
requires the user to explicitly choose **Build diagnostic Review** before it is published.

An accepted result produces a newly created, macro-free, validated `Review.xlsx`. Its visible sheets project
changes, structural events, revisions, details, diagnostics, and the current history; integrity-checked embedded
JSON stores the Phase 2 lineage, active alignment-correction generation and append-only events, alignment
findings, and exact changed-subject history without persisting absolute source paths. This embedded state is not
yet a complete copy of the source packages or a finished resumable review project. For a complete analysis, a
separately hashed and manifest-bound Custom XML sidecar also freezes the exact revision evidence, subject policy
metadata, full subject-by-revision observation matrix, groups, and initial immutable decision generation. Its
v2 form stores the complete ordered event audit and reconstructs effective decisions through deterministic
policy replay. A diagnostic-only result intentionally has no decision sidecar.

The current safety envelope is 768 MiB of input packages in aggregate (with the OOXML adapter's 256 MiB
per-file cap), 2,000,000 physically stored cells per revision and 10,000,000 audited across the project,
2,000,000 stored cells retained inside useful ranges, 10,000,000 bounded grid cells, and 32 MiB of embedded
Review state. Review import caps the compressed package at 128 MiB and aggregate declared expansion at
768 MiB. The grid limit uses a per-sheet content boundary from the first through last rows and columns containing
a stored value, saved calculated value, or formula. Physically stored blank or formatting-only cells outside it
are reported and released before the
revision snapshot is retained; they count only against the larger audit ceilings and neither expand alignment
nor consume the retained/grid ceilings. Blank storage inside the content rectangle is retained for alignment
safety but is not reported as content.
Unrelated and legacy entries are limited to 64 MiB; the eight canonical, workbook-listed and relationship-bound
projection worksheets may stream up to 512 MiB each, with one unique bound shared-string table up to 256 MiB.
The streaming projection is capped at 8,000,000 cells and 100,000 hyperlinks. Decision sidecars are independently
capped at 16 MiB, 10,000 events, and bounded per-event collections.
Progress, cancellation, and a privacy-safe learned elapsed-time estimate are shown during the longer
phases. Synthetic Pipeline4-shaped alignment regressions and the supplied real R1.0/R1.1 sorter pair are green,
but threshold tuning against a broader real-workbook corpus remains pending. Multi-revision `.xlsb` and CSV/TSV/TXT inputs, arbitrary duplicate
row/column resolution, visual snapshots, complete source-backed project reopen/resume, Excel-sheet synchronized
context menus, the transactional merge writer/UI, and complete archive export/import remain pending. The immutable
ledger persists and replays full event generations, and the task pane now invokes the tested copy-on-write API to
publish one new validated Review revision without modifying its source. Formula/value merge publication remains
deliberately unavailable.
These integrity bindings detect corruption and partial/inconsistent edits. They are not a digital signature and
do not establish authenticity against a coordinated rewrite of the whole workbook.

Phase 4 now also has an internal, Excel-free safety substrate: exact fail-closed formula/value evidence; a bounded
create-new local-NTFS publisher; a canonical four-checkpoint portable merge journal; a separate local create-new
checkpoint store; a deterministic restart matrix; and an immutable evidence-bound merge-plan builder. The plan
freezes explicit current Formula/Value candidate choices, exact typed preconditions/outputs, master targets,
revision byte identities, status counts, and deterministic decision-set/plan hashes while still requiring
all preflight/verification gates. A pure compiler combines compatible aspects into complete cell-content intents
and leaves ambiguous master facts as explicit proof requirements. A closed-construction receipt binds one proof
per required validation to the exact plan/synthesis. A read-only OOXML master preflight retains the exact source
occurrence, owns/hashes its bytes, proves target preconditions, and exposes copy-from-owned-bytes for a future
writer; its portable proofs omit the local path. These components reject stale decision context, non-mergeable
subjects, conflicting/ambiguous targets, path aliases, overwrite races, journal gaps/forks, missing stages, and
output hash mismatches. They are not wired to a command and do not yet write cells, produce the remaining
findings/formula/writer/output proofs, commit project audit, or make this alpha merge-capable.

## Exact comparison preview

**Exact two-file preview** opens a bounded setup window with explicit **Before** and **After** roles. It reads
the selected sources without modifying them, compares stored values, saved calculated values, and exact
formulas, and creates a
macro-free `Review.xlsx` whose visible change rows are integrity-bound to embedded comparison state. Existing
Review workbooks can be validated and reopened from the same ribbon menu, the Hub, or the Comparison workspace.
The workspace keeps the validated report selected and shows its filename/path, project ID, state/projection hash
prefixes, chunk count, and state size before launch. Opening revalidates the report and never activates an
already-open workbook merely because its path matches: the add-in copies the exact validated bytes to a private
random launch directory, holds that copy against writes, deletion, and replacement, forces automation security
off, and asks Excel to open it read-only with link updates disabled. The original Review workbook is not changed;
**Save As** creates an independent working copy.

This separate command is deliberately labelled **Preview**. It supports pairwise `.xlsx`/`.xlsm` comparison
by exact sheet name and positional comparison of CSV/TSV/separated TXT files. It does not use the
multi-revision lineage/alignment engine, create snapshots, merge edits, or parse `.xlsb`; the setup window
repeats these limits before it runs.

## Delimited-file workspace

- **Open files...** loads one or more `.csv`, `.tsv`, or separated-value `.txt` sources into a shared
  workbook, one source per worksheet. **Open folder...** loads the supported files in the selected folder
  non-recursively and in deterministic filename order.
- A generated **Index** sheet contains internal links plus source path, created/modified/opened timestamps,
  byte size, row/column counts, delimiter, encoding/BOM, line endings, warnings, and state.
- File identity is the canonical full path, while runtime sheet identity uses a hidden workspace ID plus
  `Worksheet.CodeName`. Two different paths named `signals.csv` can therefore coexist as `signals` and
  `signals (2)`; reopening the exact same path activates its existing source instead of duplicating it.
- `.tsv` supplies a tab fallback even when the file is empty or single-column; a leading `sep=` directive
  remains authoritative. CSV/TXT delimiter detection continues to consider comma, semicolon, tab, and pipe.
- **Save file** resolves the source from the active worksheet. **Save All** processes every edited source
  sheet. The Index and unrelated sheets are never silently treated as source files.
- A SHA-256 fingerprint of the raw opened/saved bytes detects external edits independently of timestamps.
  Conflicts offer Reload from disk, Save a copy, an explicit overwrite/recreate with `.bak`, or Cancel.
  The disk bytes are checked again immediately before replacement.
- A hidden boundary pair and movable end sentinel preserve blank and trailing-blank records; they are excluded
  from source saves and separated from the visible data region used by Excel data tools. One worksheet row and
  two columns are reserved internally. If Reload would rebuild a sheet containing view-only helper values or
  formulas, the add-in identifies that content and asks before discarding it.
- Source associations live for the current Excel process; reopening the original files rebuilds them.
  The workspace workbook is a view of the files, so closing it never triggers Excel's own "save
  changes?" prompt; unsaved source-sheet values get the add-in's prompt instead, with Save All.
- **Reload from disk** on demand: double-click a file's blue-bordered **Status** cell on the Index and
  its sheet is rebuilt from the current bytes (the same path the external-edit conflict offers),
  asking first when local edits or view-only helper content would be discarded. The Index rescans
  the files whenever it comes into view — on a sheet switch and when its window regains focus — so
  an "edited externally" flag shows without switching sheets.

## IO Pipeline

The IO Pipeline is the plant-commissioning tool family: read-only pulls that compare a plant CSV
against an electrical IO list, and the IO Bridge workspace that pushes corrections back to the real
files. It lives in the Hub section **IO Pipeline**, on its ribbon groups, and under the `IoPipeline`
namespaces (each feature removable as one unit).

### IO-list pulls

Three Hub commands join a plant CSV to an electrical IO list. They target the
current IO-list format only; the pre-2026 layouts (a `PC Board Rack.Slot.Card` column, Siemens
`I310.0` byte.bit addresses, mnemonics spread over `Q*`/`P*`/`ASi` sheets) are no longer read.

**What the IO list must look like**

- One worksheet per card, named `R{rack}.S{slot}.C{card}`. Each number may have any number of digits and
  a label may follow (`R1.S2.C1 - PC2`). **The card number comes from the sheet name** — sheets that do
  not match are ignored, so `COVER`, `Chutes`, `ASI Addressing` and the like need no special handling.
- A slot holds one controller and a controller holds several cards, so **a card number is unique only
  within its slot**: `R1.S2.C1` and `R1.S3.C1` are different cards, and in practice they carry
  overlapping addresses. A plant CSV belongs to one controller and its `Card` column is scoped to it,
  but nothing in the CSV names the slot. `Pull IO-list data` therefore indexes exactly one controller —
  silently when the workbook holds only one, otherwise asking which. `Pull mnemonics` takes the
  controller from the active sheet, and names it in the summary so the chosen CSVs can be checked
  against it.

**How a plant CSV is recognized** (shared by the pulls, the IO Bridge workspace and its apply
engine — `CsvColumnsMap.Locate`): by content, never by header text alone. The key columns are
found by their **names** on whatever row resolves them (any of the first 25, row widths free —
the ISPC export's data rows carry a trailing `;#;` marker column that its header line lacks, and
that must not matter); when no name resolves, the file kind's **legacy fixed positions** take over
(Inputs: Signal ID 2nd, Node 12th, Offset 13th, Bit 14th, Card 16th; Outputs: name 2nd, Node 5th,
Offset 6th, Bit 7th, Board 9th — 1-based), accepted only when the data there actually reads as
signals (text names, numeric addresses) — so picking a wrong CSV (a `BusField.csv`) still refuses
with a plain message instead of silently mis-joining addresses. The pulls' per-file summary notes
when the positional fallback carried a file.
- A header row within the first 20 rows carrying both a `Bit` and a `Mnemonic` column.
- Signal rows put the full address in `Bit` as `I.node.offset.bit`; `Q` is accepted for outputs and
  normalized to `O`. Components compare numerically, so `I.0080.3.0` and `I.80.3.0` are the same address.
- Module rows leave `Bit` empty and carry the node in `ID`.

**What the CSV must look like** — `Node`, `Offset`, `Bit` and a card column. Inputs and Outputs are
exported with different schemas and different delimiters, and both are accepted: the name column may be
`Signal ID`, `Input Name` or `Output Name`; the card column may be `Card` or `Board number`; the type
column may be `Type`, `Input Type` or `Output Type`. Rows starting with `#` (disabled) or `!` (include
directive) are skipped, and direction comes from the file name (`*Inputs*` → `I`, `*Outputs*` → `O`).

**How rows are matched** — every row is looked up **two ways**, and each lookup gets its own column:

- `//Mnemonics by address` — what the other side has at this row's own address (the physical key
  `DIRECTION.node.offset.card.bit`). No address is shown; it is this row's address.
- `//Mnemonics by name` — where this row's own mnemonic turns up *elsewhere*, as `mnemonic@address`,
  because a signal that was re-addressed is still in the list, just somewhere else. Rows the address
  lookup already produced are omitted, so an agreeing row leaves this blank.

| `//Mnemonics matched` | `//Mnemonics by address` | `//Mnemonics by name` | reading |
|---|---|---|---|
| | `LP100301` | | agrees — the signal is where the CSV says |
| `0/1` | `LP200301` | `LP100301@O.90.0.1` | `LP200301` occupies this address; `LP100301` actually sits one bit further |
| `1/2` | `AUTOM/MANUAL` | | one IO-list row carries both mnemonics for that bit |
| `0/4` | `DJ00142/DJ00143/DJ00144/DJ00145` | `DJ00140@I.91.10.6` | none of the four is this signal; it lives one byte away |
| | `CABINETAL00161/ISPC_GROUP_1_NUMLINES` | | a module row (node 236), matched by node |

Addresses are shown normalized — `Q` renders as `O`, and a module row shows `#{node}` rather than a bit
address, so it is never mistaken for one.

Two CSV types are handled apart from the address key:

- **Virtual types (5 VI, 6 VO, 9 AIS, 10 AOS, 13 VAI, 14 VAO)** — not hardware I/O; their address
  fields are frequently stale (every export on file writes AIS/AOS at 0.0.0), so they are excluded
  from matching and counted separately in the summary. The CSV type taxonomy is the controller's own
  `SIGNAL_TYPE_*` defines: 1 DI · 2 DO · 3 AI · 4 AO · 5 VI · 6 VO · 7 DEI · 8 DEO · 9 AIS · 10 AOS ·
  11 BUS · 13 VAI · 14 VAO (12 skipped: inputs are odd, an output is its input + 1). The tooling's
  VDI/VDO spellings mean 5/6.
- **Type 11 (bus diagnostics)** — address a device, not a bit. They join on `card` + `node` alone,
  against the module rows. The IO list prefixes those mnemonics with `BUS_`; the prefix is dropped on
  the way into the CSV, since the CSV never carries it.

**Commands**

| Command | Runs on | Writes |
|---|---|---|
| Build IO //Address | plant CSV sheet | `//Address` (`I.80.3.0`) from Node/Offset/Bit |
| Pull IO-list data | plant CSV sheet | `//Mnemonics matched`, `//Mnemonics by address`, `//Mnemonics by name`, `//ID`, `//Bit`, `//Description`, `//Functional unit`, `//Location`, `//Device`, `//Type`, `//Drawing`, `//Sheet` |
| Pull mnemonics | an `R.S.C` card sheet | `//Mnemonics matched`, `//Mnemonics by address`, `//Mnemonics by name` |

**The matched ratio** — `//Mnemonics matched` leads, next to the column it judges, and is **blank when
the two sides agree**, so only disagreements draw the eye. Otherwise it reads `matched/total`:

- *matched* — how many of the row's own mnemonics appear among the address-matched ones. Comparison is
  by **set**: order is irrelevant, a `BUS_` prefix is ignored on either side, case is ignored, and
  duplicate or empty parts collapse.
- *total* — how many mnemonics the **IO-list side** carries, since the IO list is the reference. On a
  plant CSV that is the pulled `//Mnemonics by address`; on a card sheet it is the row's own `Mnemonic`.

Only the address match is judged — a mnemonic found elsewhere by name is reported in its own column and
never counts here, or a re-addressed signal would score as agreement. An empty address match yields no
verdict at all (blank), since there is nothing to compare.

This makes the two kinds of disagreement legible at a glance: `1/4` means the signal is present but
shares its address with three others (common, usually harmless — one IO-list row legitimately covering
several mnemonics), while `0/1` or `0/4` means it is not there at all and the two sides genuinely
disagree. A conditional format paints the ratio and the `//Mnemonics by address` cells wherever a ratio
was written, keyed on that column, so an Undo restores the values and the highlight follows them;
re-running replaces only this rule and leaves the CSV view styling intact.

`//Description` joins the permanent and variable halves of language 1 with a space; language 2 is not
pulled. The remaining descriptive fields describe the union of both lookups and de-duplicate their
values with `/`. All helper columns are marked `//` in row 1 and are never written back to the source
file. Every run is a single Power-Tools undo step, and each command ends with a summary reporting rows
keyed, matched by address, matched by node, found only by mnemonic elsewhere, skipped as virtual, and
carrying a ratio.

### IO Bridge

A dedicated workspace for the commissioning phase, when the IO list (electrical dept), its mnemonics
(SW engineering) and the plant CSVs (machine controller) are corrected on both sides and IO-list
additions must be integrated into the CSVs. Where the pull commands compare once and read-only, the
IO Bridge **queues edits and applies them in batch to the real files** — CSVs through the byte-faithful
CSV engine, the IO-list .xlsx directly via COM. Commands live on the ribbon group **IO Bridge**,
in the Tools Hub, and (while a session is live) on the cell right-click menu.

**Setup** (*open workspace*) — one IO-list workbook, which of its `R{rack}.S{slot}.C{card}` sheets, one
Inputs + one Outputs CSV, plus per-session tuning: an **IO column limit** (first N source columns per
sheet, 0 = all), a **compare-virtuals regex** (virtual CSV rows — types 5/6/9/10/13/14 — whose mnemonic
matches join the comparison BY NAME, for the CLASSDIAGAL-style virtuals that mirror IO-list entries)
and an **exclude regex** over IO-list MNEMONICS: a matching mnemonic drops out of the comparison
while the rest of its cell still joins — the case is a generator key riding along a real mnemonic
in plain mnemonic syntax (`BUS_CABINETAL00159/SAFETY_NUMDOORS_16`: the key tells another tool to
generate sixteen door blocks at that node), which would otherwise count as a missing name. Such a
row reads `· filtered`; its IO-side mnemonic update re-attaches the filtered keys (and a module
row's `BUS_` prefix), so the list's own conventions are never overwritten. A cell whose every
mnemonic matches is excluded whole: shown, no claim, no verdict, faded. All remembered in the
registry. Builds and refreshes run behind an opaque
"working…" curtain — Excel's SDI window churn is covered, not fought.

**Moved signals (found at a different address)** — a signal row the address pass cannot pair, whose
mnemonic the CSV carries at **another address** (the list says `I.35.1.0`, the
CSV `I.35.4.0`) — on the same card, or on **another card** —, is paired by name instead of reading `missing` beside a `CSV-only` leftover. The row
reads `Signal · moved`, verdict `moved` (never the clean green), and the CSV block shows where the
record really sits. The pass runs *before* the positional fallback, so a shifted block or a
swapped pair is not mis-paired with whichever unrelated row now occupies the old address. A record on another card is taken only after every selected card has claimed its own, so it
never robs the sheet it belongs to; the row then also reads `· card N`, and only **A** is offered:
it moves the record to the IO row's address AND card (the Card field is set too), since an IO-list
address cannot name another card. A name the CSV carries twice is never treated as a move. The existing address actions
align either side: **A — Update address on CSV** (the record moves to the IO row's address; only the
node/offset/bit parts that differ are queued) and **B — Update address on IO list** (the list takes the
record's address). An alignment is **not offered on the side where the destination is already taken**
— another record at the IO address, or another IO row at the CSV record's address (a swap or a
shifted block) — because it would collide; the other side stays available.

**Virtual twins** — a virtual CSV record (type 5/6/9/10/13/14) that carries an IO point's mnemonic
LOOKS like the signal but acquires nothing from the field. Such a record is the point's *twin*:
the join binds it to the IO row flagged — row label `Signal · virtual twin`, verdict `virtual`
(never the clean green), a **red outline** around its CSV block — instead of leaving the row
"missing" (which would offer an integrate the apply then reports as "applied externally", the
record being already there). The row action **T — Change the CSV type** gives the twin a real
type through the same picker Integrate uses (queued as a Type field update, resolved by mnemonic
at apply); the address update then follows as for any row. A real record at the address always
wins over a twin; virtuals the compare-virtuals regex opted in are *expected* and keep joining as
clean matches; a virtual with no IO twin stays `CSV-only · virtual`. Two cells per row carry a
**blue outline** and open the sources on double-click: the IO list at the row (the *Functional
unit* column, or Mnemonic when a list lacks it) and the row's CSV in the PT_CSV workspace at the
record (the *CSV Card* column).

**The view** — one sheet per card, IO-list rows in their own order and **columns in their source
positions** (the workspace reads side by side with the original; source column widths and the
revision-tracking interior colors come along, extracted headlessly from the file bytes). Then the CSV
block (orange header, vs blue for the IO side), then `//Match` / `//Row` status columns; unmatched CSV
records sit in a *CSV-only* section below, and the header row carries a ready AutoFilter. Font colors
tell the state at a glance: clean rows green, unused spares (no descriptions, no device) and excluded
rows grey, anomalies (`missing`, `n/m` ratios, CSV-only) full contrast. **Groups** are outline runs
(+/- buttons) over contiguous same-group columns, configured on the `_Columns` sheet (group +
hidden-by-default per column, registry-persisted, *column config* applies) — Excel merges two groups
that touch, so leave a column ungrouped between them. Sheets are always **rebuilt fresh**; a saved
workspace workbook is a dead copy.

**Joining** — address/node keys as in the pulls, with **name-aware claiming**: when several IO rows
share one address (a DEO struct binding point above the actual DO), each CSV record goes first to the
IO row whose mnemonic carries its name; only unclaimed records fall back positionally. Several CSV
records on one point (AUTOM/MANUAL pairs) all bind: values join with `/`, the verdict compares name
SETS, and those CSV cells are view-only (per-record work goes through the expanded view). A named BUS
module without its type-11 record reads `missing` like any gap.

**Editing** — typing into IO or CSV data cells **queues** the change (the cell keeps the typed value,
painted pending-yellow immediately; typing the old value back unqueues). *Row actions* — with
letter accelerators, on one row or a whole selection, **visible rows only** (a selection drawn
over a filtered sheet spans the hidden rows too; the action lands only on the rows you can see) — queue the linked-pair updates in either
direction (M/N mnemonic, A/B address, C/D condition), description→CSV (E), **integrate** (I;
signals ask the CSV Type once per direction, bus modules go straight in as type 11 with the `BUS_`
prefix dropped; empty-mnemonic spares are never integrated), disable/enable (X/V) and delete (Z).
The CSV **description** is composed exactly as the export tool writes it —
`DESC1A DESC1B +LOCATION-DEVICE [DRW]` (drawing code = after the last `_`) — automatically on every
CSV create and CSV content update, and on demand via E. A queued integrate shows the record it will
create across the row's CSV block.

**Expanded view** — for a multi-mnemonic point: a temporary `_Expand` scratchpad, matched pairs first
(IO mnemonic column split per constituent), then IO-only, then CSV-only lines. Edit it freely, then
*row actions* there offers the reconciliations, reading the sheet AS IT STANDS: **M** merge the
sheet's CSV names into the IO-list mnemonic cell, **U** make the CSV match the sheet (delete vanished
records, append new ones cloned from a sibling), **A** update the IO address from the CSV side.

**Closing** — the workspace workbook is a view; closing it never triggers Excel's save prompt
(every decision lives in the sidecar).

**The queue** — `_Queue` lists actions in order, Include first (double-click ☑/☐, typed x/1 count
too; unchecked = skipped but kept). Repeated edits supersede but keep the **first** old value — the
three-way base. Everything persists in `<io-list>.iobridge.json` (atomic, `.bak`), written on every
change; on rebuild each action **re-anchors**: file holds the old value → Pending (typed value
re-shown), the new value → AppliedExternally, anything else → ConflictStale.

**Apply** (*apply queued batch*) — per-action **three-way check**: the file's current value must still
equal the action's old value or that action flags Conflict and the rest proceed. CSVs are patched
headlessly and written only after `VerifyRoundTrip` passes and the file is confirmed unchanged since
reading (atomic write, `.bak`; a nameless create payload is refused; a CSV open in the delimited
workspace is skipped). The IO list is patched via COM — build-time row re-verified, then
`WorksheetFunction.MATCH`/`COUNTIF` on the exact Bit text (never whole-sheet `Find`, which skips
filter-hidden rows). An IO list already open in Excel is patched live and you are **asked** before
saving; declining leaves the actions pending. Statuses flip only after a real write, the workspace
rebuilds to show the new reality, and *clear history* drops Applied rows.

To remove the whole feature: delete `src/ExcelPowerTools/IoPipeline/IoBridge`,
`src/ExcelPowerTools.Core/IoPipeline/IoBridge`, `tests/ExcelPowerTools.Tests/IoPipeline/IoBridge`,
the pane entries, the ribbon group and the three `AddIn.cs` event lines.

### SWP Bridge

The SWP workbook holds the plant's diagnosis data — states, alarms, warnings — one sheet per
object type, each device a fixed block whose cells are formulas pulling template texts from a
CONFIG_PLANT companion. The legacy VBA DiagnosticTool updated it blind; the SWP Bridge replaces
that with a **reviewable, monodirectional (IO list → SWP) diff**: only cells that would actually
change become pending actions, and a headless batch apply patches the file bytes. Commands live on
the **SWP Bridge** ribbon group and in the Hub.

**Setup** (*open workspace*) — pick the SWP file inside its dated version folder (`NN_dd-MMM-yyyy`);
the system **root** (the parent) and the CONFIG_PLANT name derive from it, and a `swp-bridge.json`
sidecar found at that root pre-fills everything else — returning to a system is one pick + OK.
A system can also be reached from the other end: the **Session** row's *Open…* lists every saved
session (`*.json`, plus `*.json.bak` for recovery) and **replaces** the whole form with the one you
pick — its settings, per-row decisions, dictionaries and queue (the automatic prefill only ever
completes what the SWP pick left blank). That is how a system whose SWP was renamed or moved is
reopened — the dialog then says which workbook the session expects and asks for it — and how a
damaged sidecar is recovered: open the `.bak`, and the next decision writes it back as the live one.
If another session's apply has since created a newer `NN_…` folder holding this session's SWP
(copied along, see *Copy along* below), the session **follows the newest one**, and the line under
the SWP names the folder the session had recorded.

*Save as…* beside it saves the session under a new name, **anywhere** — and from then on the session
**is** that file. It is written at once with everything above; the dialog works on it; a workspace
running on the session the dialog had open writes its next decision there (cancel instead of
building, and the file is rewritten with the settings that workspace actually runs); and Setup opens
on it next time — the running session first, then the last one saved or built, then the last SWP
root. That is how one IO list drives two SWP systems as two named sessions kept together in one
folder, which is where both dialogs start. Picking an SWP never swaps a named session for a root's
default `swp-bridge.json`, and the Home sheet's header line names the session file first. A session
saved at its own root stays **self-locating** (it records no root, so moving the whole system folder
keeps working); only one saved elsewhere records which system it belongs to. Building a session for
a *different* SWP than it was saved for would overwrite its decisions, so the dialog says so first
and points at *Save as…*.

*Save as new…*, the third button, starts a **new** session from the one in hand, saved under a new
name anywhere, carrying **everything that is not history** — every setting and file reference, the
dictionaries, the glossary corrections and the strobe templates — while the queue's history and the
per-row decisions start **empty**. It is the other half of "one IO list, two SWP systems": *Save as*
copies the whole history, which is right for renaming or moving a session, but carried into the
second system it is the first system's decisions riding along on mnemonics both cards share. The
dialog switches to the new file (and Setup opens on it next time); a workspace running on the old
session does **not** move — its queue and decisions stay in the file they came from — and building
from the dialog rebuilds into that same workbook with the fresh session.

A **reference SWP is required** — a known-good workbook of the same plant type. It is the only
thing that can tell a row whose formulas drifted from one that is merely different, and the only
way back for a row whose CONFIG_PLANT template reference was destroyed, so a workspace without one
has no answer when a formula turns out damaged.

Then: plant type (**MSC** sorter / **ISPC** induction / **UPS** upstream / **SAFETY** for the
direct-to-SCADA safety PLC — MSC-shaped mnemonics on a cabinet-only SWP) with the first-line and
group parameters. Those two are **read off the SWP's own tab names** — `IL02_…`/`IL03_…` with
`ILGRP02_CABINET` gives first line 2, group 2 — because the mnemonic table builds its destination
sheets from exactly those numbers, so retyping them is re-deriving what the file already states and
one wrong digit sends every row to a sheet that does not exist. The note under the boxes names the
sheets they came from; the values stay editable. UPS tab names carry the system id but no group
(the conveyor group is a `SysGroupNN` fact of the PRODUCTION tree), and MSC/SAFETY have no
numbered sheets at all — there the typed values stand.

Then the workspace's **second-language code** (a free 2-letter glossary code —
PT, FR, IT, …), the IO-list workbook + its `R.S.C` card
sheets (with the include/exclude regexes that decide which of its mnemonics are diagnostics at
all) — and when the workbook names **no** `R.S.C` sheet, the sheets whose header row carries
**Mnemonic + Bit** are offered instead (the pair the card reader itself needs), the label
reading *Card sheets (by header: no R.S.C names)*: the safety PLC's list names its one card
sheet by network (`NET SAFETY 192.168.20`) and its cover, cabinet table and fill helper carry no
such pair. Names first, content only when the names name nothing, so a sorter list never grows
an extra ticked sheet because a summary tab happens to hold a Mnemonic column — and ONE
optional **Production root** — the realtime `PRODUCTION` folder (or the backup
folder that contains it), resolved live by the real tree layout: `IODriver\BusField.csv` for every
plant; `MSC\DB_Csv\` (Plant_Inputs, Destination_Parameters for the chute remap,
BUS_ExtendedDiagnostic for the AS-I slave bindings) and
`DataDrawSupervisor\DataDraw_PlantData\AlarmLists\` for the sorter; `ISPC\db_csv\Plant_Inputs.csv`
and `ISPC\AlarmLists\` for induction; `ConveyorSys\SysGroupNN\Database\DB_Params\` (Sensor,
StrobeDevice for the generators) and `…\DB_Diag\` for upstream (the configured group, else the only
one); flat copies directly under the root work too. **SAFETY consumes only the SCADA's
`Alarms_*.txt` export** — a Siemens safety PLC talks to SCADA over its own OPC UA server and has
neither an IODriver busfield nor a realtime Plant_Inputs, so neither is listed — probed in an
`Alarms List\` child (where the safety system keeps it, beside the SWP), an `AlarmLists\` child,
or flat under the root; the SWP's own version folder therefore serves as the root. The
dialog lists what resolved (✓ green) and what stays off (✗ red) — only the inputs the plant type
consumes are listed at all; the resolution re-runs at every build and
the resolved paths land in the sidecar (a sidecar from before the root keeps its per-file paths
until a root replaces them).

The write rules are grouped by the QUESTION they answer, and the last two have the same shape
because they are the same kind of question: **one exclusive default, one destination-sheet regex
naming the sheets that take the other mode, and a per-row column that wins over both** (resolved
most-specific-first: tick → sheet rule → default). Every caption names the mode it applies and
flips with the default; every answer is stored as the mode it chose, never as a "flip", so changing
a default later can never silently reverse a recorded decision; and each build reports how many
rows a pattern moved, because a rule moves rows with no visible gesture on them.

**CONFIG_PLANT** — what happens to the template's text, as two radio buttons:

- **Never overwrite CONFIG_PLANT descriptions.** The text SCADA shows is CONFIG_PLANT's template
  *concatenated with* the SWP's description cell (`SUBSTITUTE(CONFIG!template & T{row}, …)`), so
  writing the composed IO description into T stacks a second sentence onto a template that often
  already says the same thing. Chosen, every row keeps CONFIG's wording and the bridge contributes
  only the FLD label — and no description edit at all when there is no mismatch, so an existing
  template-fed cell is never blanked. It keeps CONFIG's **words**, though, and never starves the cell
  a kept formula shows: where the formula already displays `T`/`U` after a template that says nothing
  past its last placeholder — CABINET's `INTRODUCTION %1 - ARMOIRE DE GROUPE %2 - ` followed by
  `&" "&U{row}` — there is no sentence to stack onto and that cell is the alarm's only name, so the
  composed description is written there in both languages while the formula stays untouched (`_Log`
  counts those rows). It is read from what the cell displays — the cached text, the description's
  own text removed, cut after the block label — and a display that cannot be read (no cached value,
  a file patched and not yet recalculated by Excel) keeps the label-only treatment. Withholding the
  description there is what left a French-only SCADA reading the English an older run had put in `U`.
- **Replace the template's trailing text with the row description** — the **Q/R formula rebuild**,
  below.

They are radio buttons because they are two answers to ONE question: a row cannot both keep the
template whole and cut its tail off. (Before, they were independent checkboxes, and the third
state — compose the description *and* leave the formula alone — was the stacking shape above.
A sidecar carrying neither switch migrates to the rebuild and says so in its load notes.) The
per-row column decides outright, and its caption flips: `[Keep CONFIG text]` under the rebuild,
`[Write IO text]` under "never overwrite". Rows a rule or a tick excepted are counted in `_Log`.

A CONFIG template reads
`INDUCTION %1 - BELT %2 - SPARE ALARM.63`: a head naming the place (`%1` = the sheet description
`B6`, `%2` = the block's electrical label) and a tail naming the bit generically — while the SWP's
own `T`/`U` cell already names that bit in the plant's words. Chosen, every joined row's formula is
rebuilt to keep the template **up to its last placeholder** and put `T`/`U` where the tail was, so
SCADA reads `INDUCTION =S1+IN22 - BELT =S1+IN22.05 - FIELD BUS FAILURE - DRIVE [n0115-…]`. The
emitted shape keeps the row's own template reference and truncates it **before** substitution:

```
=LET(t,[1]DB_IL_BELT!$K$122,p,FIND(CHAR(1),SUBSTITUTE(t,"%",CHAR(1),LEN(t)-LEN(SUBSTITUTE(t,"%","")))),
     SUBSTITUTE(SUBSTITUTE(LEFT(t,IFERROR(p,LEN(t))+1),"%1",$B$6),"%2",F679)&" - "&T778)
```

Truncating first is not a detail: once `%1`/`%2` are replaced the markers are gone, and cutting the
rendered text at `" - "` breaks on every template whose own tail contains one (`FIELD BUS FAILURE -
DRIVE`). `CHAR(1)` is the sentinel that finds the last `%`; `IFERROR` keeps a template with no
placeholder at all from poisoning the cell. The file spells it differently from the formula bar:
`LET` is stored as **`_xlfn.LET`** and its parameter names as **`_xlpm.t`** / **`_xlpm.p`**. Each
half matters. A modern function spelled bare comes back as `_xludf.<name>` and shows `#NAME?` —
what the previous wholesale `CONCAT` rewrite did. A LET parameter spelled bare is worse: Excel's
loader **removes the formula on open** ("Removed Records: Formula from …"), keeping only the cached
value — what the first build of this rebuild did. Excel displays the stored form back as
`=LET(t,…)`, as shown above. A cell still holding that first, unloadable shape is recognized and
rebuilt into the loadable one on the next build. A Q/R cell that holds a plain value instead — a
paste-over, or a formula Excel removed while repairing a file, frozen at its last correct text —
has no template to rebuild from and can never show `T`/`U` again; it is **named in `_Log`**, not
skipped, and the reference SWP or the file's `.bak` is what restores it.

Two rows never get the rebuilt shape, because on them the tail is the only thing naming the alarm:
a row **excepted** into "never overwrite" (its `T`/`U` takes the FLD label alone) and a **mismatch
row the IO list gave no description**. Those keep the plain `&" "&T` append. A row whose template
reference is **gone** — what the old wholesale rewrite left behind — cannot be rebuilt from itself:
it is named in `_Log` and left alone, since restoring a lost template is the reference-SWP formula
audit's job. The **Formulas** column on the spine reads `rebuild` or `append` accordingly.

**ELECTRICAL_DESIGNATION** — where the functional unit / location / device land — is the same
shape: the append/overwrite radio pair, its own destination-sheet regex, and the per-row
`[Append FLD]` / `[Overwrite FLD]` column. It is described in full under the FLD rule below.

The exported **Alarms_*.txt
files** feed the third coherence leg:
every checked mnemonic must appear among the exported alarms and the exported texts must carry
the row's Location and Device (description 2 checked only when exported).

**Only the lists this SWP exports are read.** A production folder commonly holds the alarm lists
for a whole plant — or for every system of one type — so they can be collected from one place, and
each controller then picks its own because the indexes are global. Every diagnostic sheet names the
file it exports in **row 5, "Output filename"** (`Alarms_S1_CABINET.txt`), and that set decides:
a file no sheet claims is **skipped with an `_Log` Info line** rather than parsed, because a
sibling controller's list is shaped for different mnemonics and every line of it would otherwise be
reported as malformed (a sorter SWP handed the induct list read `Alarms_S1_IL03.txt line 414:
mnemonic "IL03.BELTWR0630" has the wrong length`, and so on for the whole file). The loaded-count
line names how many files were left out. Two guards keep the filter from hiding real problems: an
SWP whose sheets declare nothing — an older file — reads **every** offered file rather than
silently dropping them all, and a folder where **no** offered file belongs to this SWP is a
Warning, since "nothing to check" must never read as "everything checks out". WHICH objects are
checked and with WHAT index shift comes from the **`_Offsets` table** (below) — the official
values reproduce the v49 "ddsng 9.4.2" semantics (MSC cabinets export +1, ISPC line cabinets
1:1 with groups −1, UPS cabinets −1 with belts/zones 1:1, SAFETY 1:1 zero-based). The Workspace
sheet's **Alarm strings** verdict column reads ok / not exported / text differs, with the details in
`_Log`; a file whose tags repeat warns that the export was appended-to (ExcelLoader never
overwrites — delete the old `Alarms_*.txt` and regenerate). The sorter's fourth leg is
**BUS_ExtendedDiagnostic.csv**, which binds AS-I slave states to their CLASSDIAG mnemonics: a
slave row's Slot (`1M`/`2M`, empty = line 1) + ID (`24A`) compose its identity `1M.24A`, the
gateway it hangs off is the `PA` row above it in the card sheet, and the sub-device number follows
the gateway image layout (line 1 A 1..31 / B 33..63, line 2 A 65..95 / B 97..127 — the IO list's
*ASI Addressing* sheet, kept in code); the **BUS ext. diag** column reads ok / not bound / node
differs / sub-device differs, and such slaves are no longer looked up in the busfield (they are
not Profinet nodes). Double-clicking any of the four verdict cells opens its document at the
matched line — the CSVs in Excel, the `Alarms_*.txt` in Notepad++ (Notepad when it is not
installed). An include/exclude **regex pair**
extends the built-in diagnostic-mnemonic patterns — an include match is an explicit override and
bypasses the length rules too. `.xls` is refused — save as `.xlsx` first. Identity fields
(FUnit, Location, Device, Profinet name/IP, address, ID) are cleaned of known **stray
characters** on read — the leaked Excel apostrophe, typographic quotes, backtick, zero-width
characters, NBSP/tab — a deliberate DENYLIST (unknown characters pass through untouched
rather than being silently deleted), with one `_Log` line per distinct cleaned value.

Running *open workspace* again while one is open **reconfigures in place**: after a confirm, the
new settings rebuild INTO the same workbook — the window, its position and the workbook itself
survive, so flipping a write rule (overwrite ↔ append) no longer closes anything. Decisions live
in the sidecar and are untouched either way.

A **Home** sheet opens the workspace: the six command buttons, the four-step loop (Review ·
Inspect · Steer · Apply) as a strip, a legend painted with the real swatches, a sheet map whose
names are jump links, the double-click gestures as a Where · Cell · What table, and the list of
the edits that stick. Closing the workspace workbook never triggers Excel's save prompt — the
sidecar is the state.

**The Workspace sheet** (the spine) — one flat sheet, one row per diagnostic mnemonic in
IO-list order: the IO side, the
computed SWP coordinate (`DB_TAG.AlarmN>entry>bit`, the legacy address math ported verbatim —
CABINET/CLASSDIAG for MSC with class types 62 CHUTE, 64 SCANNER, 66 CONTROL_ZONE,
69 CHECK_POSITION, 73 CELL_TEST, 112 AIR_BLOWER and 113 ITEM_LOCATOR — SAFETY takes the sorter's
CABINET and EMERGENCY shapes and **nothing else**, a safety SWP being the cabinet sheet alone
(rules §9.2), so a CLASSDIAG mnemonic in a safety list is simply not a diagnostic there; its
signal rows carry Siemens absolute addresses with no node (`I.0.4`, `I0.4`, `%I12.7` — a
distributed input is unique PLC-wide), which SAFETY alone reads and writes to the SWP's I/O
address as the TIA form `I0.4`, outputs skipped as ever, while on any other plant type that
three-part shape stays the "unreadable I/O address" error it is (a sorter row missing its bit) —
CABI/BELT/ZONE →
IL/ILGRP sheets for ISPC, ALLQ/ALLB/ALLZ quads → UP sheets for UPS, EMERGENCY→CABINETAL and
EMERI_→ALLQ_ rewrites included — anchored at the START of the mnemonic, so command signals
like `A_RESET_EMERGENCY` are neither rewritten nor treated as diagnostics),
the row's CURRENT SWP values, the proposed values (blank when nothing changes) and the verdicts.
Headers are colored by SOURCE — blue for the IO list, lavender for the checked production files,
orange for the SWP, olive for the override columns — and every column autofits to its largest
cell. The block identity is one column per field, **FU (IOL / SWP)** and **Loc (IOL / SWP)**
right before Device: the value alone when both sides agree, `io / swp` (tinted) when they differ.
The **Formulas** column is the row's formula status: `append` / `rebuild` for a queued Q/R
description-formula change, `repair H` / `repair I` / `repair H+I` when the cleaner queued a
restoration of that SWP row's FUnit/Location formulas — painted with the action's queue status.
The SWP-related headers wear **light orange** (the IO side stays blue) for a quick spot. Row
states are font colors — green up-to-date, red missing/invalid/duplicate, orange label mismatch;
pending "▸" cells carry the queue highlight (yellow pending, red conflict, blue applied
externally). Two mnemonics resolving to the same SWP row are surfaced as **duplicates and queue
nothing** (the legacy silently let the last one win). An SWP row whose identity cell (A, C or D)
caches an **error value** (`#REF!` after a row deletion) cannot be matched and receives no update
at all — `_Log` lists such rows (ten per sheet, then a count) instead of letting them fail silently. Rows without a bit address take their node
from the ID column — numeric Profinet nodes and **AS-I `nA`/`nB` addresses** both count (the SWP
receives `236.0.0` for a Profinet module and the composed `1M.24A` — line.slave, no `.0.0` — for
an AS-I slave); any other ID leaves the row addressless, its G untouched — this is
how the sorter's virtual CLASSDIAG rows come through.

**Overrides** — the `[bracketed]` columns are the editable exceptions and persist per mnemonic in
the sidecar: *Mismatch OK* (the FU/Location mismatch is deliberate — silences its warning and
verdict; the configured treatment still applies), *DB index* (reroute the device index,
empty = AUTO), *Force IDs* (on a `*BELT` sheet the device cell J is normally left alone — a belt is
a run of conveyor, not one labeled part; ticked, the row writes J like any other: the IO device in
overwrite mode, **blank** in append mode, where the device rides inside the appended label instead),
and the two description overrides.

Two override columns **name the opposite of the workspace default**, because that is what ticking
them does — `[Overwrite FLD]` / `[Append FLD]` flips where one row's identity lands, and
`[Keep CONFIG text]` / `[Write IO text]` flips whether the bridge's composed text is added on top of
CONFIG_PLANT's. Cleared, the row follows the default again. Each tick is stored as the answer it
chose, never as "flip", so changing the workspace default later cannot silently reverse a row's
decision; a row whose stored answer has become the default simply shows unticked. The FLD rule
holds under either FLD choice — the identity always lands somewhere, only *where* changes. Unlike
the workspace setting, an explicit `[Keep CONFIG text]` tick also bypasses that setting's automatic
scope, so it works on any row. Editing one recomputes the diff **in memory** — no file is re-read —
and the sheet's **AutoFilter criteria survive the re-render** (the filters you set, and the
dashboard's filter-to-sheet, are captured before the refill and re-applied after; a filter Excel
cannot express as criteria — by colour or icon — cannot be read back and is dropped). `_Dictionary`
(the description find→replace pairs, per language) and `_StrobeTemplates` (the strobe alarm
blocks) are free-edit sheets saved on every change and applied on the next refresh.

**Descriptions**: language-1/2 parts joined, dictionary-shortened in row order; a bus-module row
(Profinet name column filled) instead composes
`{bus-failure phrase} [{profinet name} - {ip}]`. With a glossary loaded, the **glossary is the
authority for that phrase**: an IO-list description that IS an official entry (the FIELD BUS
family — `FIELD BUS DEVICE FAILURE`, `FIELD BUS FAILURE - DRIVE/COUPLER/GATEWAY/…`) is kept with
its official casing and glossary translation; anything else is replaced by the glossary's
dedicated **`FIELD BUS MODULE FAILURE`** row (both languages) with a `_Log` warning naming the
unofficial text. Per-language fallback constants serve only glossary-less workspaces (an
internal table — the former `_Languages` sheet is gone). **Text landing in Description 2 always
passes the language-2 dictionary last**: a glossary match keeps the IO list's language-2 variable
part whenever the IO list WROTE one — a part the dictionary then empties still counts as written,
so nothing rides in its place (the FVT safety list's `CH1/CH2` markers, deleted by rule on both
sides, used to come back in raw from language 1 — 2026-09-22); only an IO list with no language-2
variable part lets language 1's ride, cleaned by the language-1 dictionary and then the language-2
one, and the `_Log` warning names the text as it lands. A blank language-2 falls back to language 1,
and once a second language other than English is configured that fallback is a `_Log` warning naming
the row: a SCADA offering only that language shows the language-1 text as it stands.
An **empty composed description never blanks an existing T / U** (a stale text still names the
alarm; a blank one names nothing) — the FLD label alone still lands, and the kept text is noted in
`_Log`. On an FUnit/Location mismatch the
mismatch **always warns** (`_Log`, Warning severity) — a deliberate mismatch is silenced per
row with the persistent **[Mismatch OK]** override (generated rows are auto-silenced), which
also clears the orange verdict; the warning's wording follows the row's FU/Location formula
HEALTH: intact formulas mean a real configuration difference, broken ones note that the
restoration is already queued (the formula cleaner) and that the comparison stands on the
restored block values. **The FLD rule** governs what the mismatch does to the writes: the IO
list's Functional unit / Location / Device are the most important information the SWP carries
(the maintainer has SCADA and the T.S. manual, nothing else — a location that only says
`+SAFETYDOORS` for sixteen doors means a service call instead of a repair), so they must ALWAYS
be visible, either in the row's own cells or in its descriptions. One exclusive setup choice:
**append identity to descriptions** (the compacted label lands in T/U, the device blanks, and
template-only Q/R formulas — ones not referencing `T{row}`/`U{row}` — are extended automatically,
since that is the only way the appended identity reaches the displayed text: with the rebuilt shape
where `T`/`U` carries a description, with the plain `&" "&T` append where it carries the label alone) or
**brutal overwrite** (the IO FU / Location / Device replace the row's own H / I / J cells — H and
I become plain values over the block formulas, which the formula cleaner reads as deliberate; the
descriptions are written as-is; the comparison then stands on the row's own cells, so an
overwritten row reads up to date). A row whose H/I are shared formulas cannot be overwritten
headlessly and falls back to the append with a warning. Fieldbus-module rows go through the same
policy — their specialness ends at the description composition.

That choice is the workspace **default**, and one **destination-sheet rule** may except whole
sheets from it: a regex under the two radio buttons, matched (case-insensitively) against the
**SWP sheet a row lands on** — `BELT$`, `^UP0[12]_` — names the sheets that take the OTHER
treatment; every sheet it misses follows the default. One pattern is enough because the treatment
is binary, so the caption beside the box names the mode it applies and flips with the default
("…but append identity to descriptions on SWP sheets matching:" under an overwrite default, and
the reverse under an append one); the answer is stored as the **mode it chose**, never as a flip,
so changing the default later cannot silently reverse what the rule does. It generalizes what was
once hardcoded as "sheets ending in BELT": whole sheets routinely want the other treatment, and
ticking every row on one by hand is not a workflow. The three levels resolve **most specific
first — the per-row `[Append FLD]` / `[Overwrite FLD]` tick, then the sheet rule, then the
default** — and on a rule-governed sheet the tick flips the row away from the RULE's mode, so a
tick always changes something. Each build reports how many rows the pattern moved (`_Log`, Info).
Setup refuses a pattern that does not parse; one that reaches a build anyway is **inert, never
fatal** — every sheet follows the default and `_Log` says so, because losing a build to a typo is
worse than treating the typo as "no exception". The FLD rule itself is untouched either way: a
matched sheet still carries the functional unit, location and device, only the place they land
changes. **An empty IO-side value never
erases what the SWP holds**: a missing Functional unit, Location or Device in the IO list is missing
data, not a blank to write — it neither matches nor mismatches, never lands in H / I / J, and never
blanks J in append mode (the blank is proposed only when a device actually moved into the label);
the SWP's own member stays, and `_Log` names the kept device. A row carrying no identity at all —
every strobe alarm the bridge generates, a photocell whose `PHT_*` row is missing — leaves the SWP
row's identity untouched: silently for a generated row, with a warning for an IO-list row. The
append extension reaches every formula shape that does not already reference the description cell;
a Q/R cell holding a plain VALUE (a pasted-over literal, which would display nothing of the label)
is rewritten as a formula over the same text with the description cell appended, noted in `_Log`.

**The glossary** (optional): setup accepts the standard *Glossario Troubleshooting* workbook and
the project's *Glossario ombra* — any sheet whose header row carries **EN plus language codes**
parses; the shadow file wins on a shared phrase. Matching is whole-phrase only, in two tiers: a
pair applies when the row's **Description permanent part equals its EN phrase** (trimmed,
case-insensitive; the variable part rides along — the IO list's **language-2** variable part when it
has one, because that is its own translation of the part, and the language-1 one only when it has
none, which `_Log` warns about once a second language other than English is configured), or — field
files split compliant
texts across the two parts — when the **trimmed full description** equals it (consumed whole,
the translation stands alone). Partial/substring matching was retired — a spliced phrase
produced wrong sentences. Two modes, chosen in setup: **fill the gaps** (rows
with IO-list language-2 keep it) or
**glossary always wins**; either way, an IO-list language-2 that DISAGREES with the glossary is
flagged in `_Log`. The **`_Glossary` sheet** shows the session-relevant pairs only — those that
actually translate something here — with a ☑/☐ Use toggle (double-click), editable translations,
and room to add custom EN→text pairs; decisions persist in the sidecar by EN phrase.

The glossary also owns the **T.S. Ref cards**: its `Numero scheda` column maps every official
phrase to its troubleshooting card, and when a row's composed description exactly matches an
official phrase, that card is what gets written to column L — an IO list carrying a DIFFERENT
card for the same phrase is flagged in `_Log` and loses ("the glossary wins"). Descriptions the
glossary does not know (electrical specifics) keep their IO-list card, per the official rules
(§12.2: electrical alarms' TS comes from the IO list). Two hard rules on top: an **empty
composed T.S. Ref never blanks an existing SWP value** (the SCADA opens the troubleshooting
manual from this code — a wrong section beats no action), and **IO-list alarms must end up
with one**: a row with no card anywhere (IO list, glossary, SWP) wears **red on its T.S.
columns** with a `_Log` warning naming the two remedies. Double-clicking any T.S. ref cell
(blue-bordered) opens the **`_TSRef` lookup** — one row per TS_nnn card with all its English
descriptors from the glossario alongside, jumped to the clicked card — for picking or
verifying a card without leaving the workspace.

**The `_Offsets` table** — the living copy of the index tables in *SWP and diagnostic
configuration rules* (§4.9 sorter, §5.2 induct line, §7.10 upstream): one row per plant type and
diagnosable object, holding the **mnemonic start index**, the **SWP table-entry start** (their
difference re-addresses the join — e.g. ISPC/UPS mnemonics are 1-based over 0-based entries) and
the **OPC-tag start** (their difference drives the alarm-strings match), plus the per-object
**Alarm check** switch — only cabinet/belt/zone tag families can be verified, so CLASSDIAG-style
objects ship unchecked instead of drowning `_Log` in false "not exported" errors. Edits persist
**globally for every system and session** and recompute live; a row for an
object not listed falls back to the plant family's defaults.

**The mnemonic map** — the table that decides **what a mnemonic is and where it lands**, one rule
set per plant type. Each rule is a **regex over the mnemonic** with named groups, plus the
destination it resolves to: the **sheet** and **tag** (templates — `CABINET`, or
`IL{firstLineId+line-1:00}_BELT` for shapes whose sheet is computed from a line digit and a
workspace setting), the **word** (`auto` derives Alarm/Warning from an `(?<aw>AL|WR)` group and
numbers it `ceil((bit+1)/32)`, or name one outright — `Alarm2`, `State`), and where the **index**
and **bit** come from (`{i}`, a constant, or `map(Destination_Parameters,{i})` for an index the
mnemonic does not carry). First enabled match wins, so a specific shape sits above the general one
it would otherwise be swallowed by. The rule also names its **`_Offsets` object key**, so the
per-object entry shift still applies exactly as before.

Three things follow from the table being data. **Nothing is renamed**: the legacy tool rewrote
`EMERGENCY02000` into `CABINETAL02000` — which worked only because both prefixes are nine
characters, so the renamed string put its index and bit on the offsets the cabinet branch already
read — and that cost the real identity, since the plant CSVs know the signal by its true name.
A rule states the destination instead. **A destination may be empty**, meaning *recognized and
deliberately not written*: the `EMERGENCY_ON`/`_RST` commands, and every member after the first of
an AND/OR group (many digital inputs drive one diagnostic bit; the `_0` member carries it). And the
table **is the filter** — "is this a diagnostic at all" is the same question as "does a rule claim
it", so a shape no rule claims is skipped as quietly as the rest of the non-diagnostic mass. What
replaces the old per-row length errors is a ratio: when the map claims almost nothing across a
real sample, `_Log` warns that the plant type is probably wrong. A rule whose regex does not parse
is **inert with a warning**, never fatal — one typo costs one shape, not the build.

An index that comes from a CSV **requires that CSV**: a chute cannot be placed without
`Destination_Parameters`, and an absent, empty or silent-about-this-key source is an error rather
than a fallback to the mnemonic's own number, which would write to whatever row that number
happened to hit.

**Official-rules checks** — beyond the coherence legs, every build audits against the
configuration-rules document and reports to `_Log` (findings only, capped, nothing queued):
a **mnemonic landing on bit 31/63** (the last bit of a DWORD is forbidden, §3.3), the same
**identity** exported twice within a device block *and category* (§3.3), an
exported row whose **priority E** is empty, zero, non-numeric or stored as text — with a gentler
note when it strays from the alarm=2/warning=3 convention (§3.1) — and **locations containing
`-`** (§3.8), on the IO side and on block headers.

Two of those read narrower than the rule text, because the wider reading produced findings nobody
could act on. **Bit 31/63** is only reported where a mnemonic actually lands: an export tick sitting
on the last bit of a DWORD is ignored by the plant and worth nothing to report, while a signal
written there will never be seen. And the thing that must be unique is the **identity — functional
unit + location + device + description — per block AND per category**, not the sentence: sixteen
safety doors on one block all read "SAFETY DOOR OPEN" and are told apart by their device
(`-B1`, `-B2`, …), and the same identity as an alarm and as a warning is a signal that was
declassed from one to the other. Keying on the description alone reported fifteen non-duplicates
per door block and one per declassed signal, which is how the real duplicates got lost.

The busfield leg accepts BOTH valid
configurations: the modern type-11 layout (§3.7 — reference index, element and bit all left 0, the
linkage living in the type-11 BUS input) and the deprecated encoding where the busfield row carries
element/bit/warning flag itself, mismatching only when neither fits. The **reference index** (the
induction line, a ≥ 10 line mask, or the upstream system) is compared only on a row that carries a
linkage at all: a modern row has no line to compare, and node + card already name the device on its
controller's network — comparing it anyway reported every modern ISPC and upstream module as a
missing node. A node the busfield comments out (`#`) or does not list is still missing. The
columns are found by the header's own names — the older IODriver export has no *Linked Master ID* /
*Linked Slave ID*, so reference index, element and bit sit two fields earlier — and only a file whose
header does not name them is read at the current positions. The leg
also wants **node 0** (the Profinet controller) present (§3.2) and every **AS-I master** carrying the
ignore-error flag (warning-on-fault 2 or 3, §4.8).

**The formula cleaner** (the Cleanup VBA, absorbed): every build checks EVERY `#`-marked
diagnostic sheet — STROBE included, targeted or not — for the structural damage hand edits leave
behind. A bit row's FUnit (H) and Location (I) must be (shared) formulas pointing at their block
header: a pasted value, a deleted cell or a formula dragged astray becomes a queued **repair**
(`=H{header}` / `=I{header}`) reviewed and applied like any other action; a block header whose
Location lost its `+` prefix is rebuilt from the electrical label, and a STROBE sheet's first
header gets its `=MID(F15,4,12)` back. Virtual objects are exempt: a block whose header carries
no FUnit or Location (SORTERSTATE, TRAIN, …) is healthy with those columns empty — no repair,
no noise. Intact shared formulas are never touched — a repair whose
cell healed in the meantime conflicts as moot. Optionally, setup takes a **reference SWP**: the
target's formula structure is audited against the base's canonical block (sheets map by name or
by suffix) and `_Log` reports columns whose formulas were lost and CONFIG_PLANT references
pointing at the wrong template row — findings only, capped per sheet. One exemption: **column L
(the T.S. reference) is not audited for formula loss on a cabinet table** (sheet parameter B4 =
`DB_CABINETS_DIAG`). The cards there are assigned by hand while the reference file happens to
carry a formula, so every hand-filled row read as damage. It is one column on one table family —
a lost Q, R, H or I formula on a cabinet is still reported.

A **block** is one physical unit of 5 DWords — State, Alarm1, Alarm2, Warning1, Warning2 — so 160
bit rows sharing a single header (the sorter's blocks measure 166 rows: the header plus 5 × (1 +
32)). A row folds onto the base by its **distance from its own header**, which the reader already
resolved for every row; no block-length constant is involved, so blocks of unequal length audit
correctly. Inside a block a same-sheet reference is never an absolute address — it names either
**the row itself** (T / U, the description cells) or **the row's block header** (F in Q / R, C in
S, and H / I) — and both anchors are collapsed before comparing. The reference SWP holds one block
per element near the top of the sheet while the target holds many, so the same structural formula
reads `F15` there and `F181` here and the two compare EQUAL; a reference to any **other** row
survives normalization and stays a finding, which is exactly the dragged-formula damage the audit
exists to catch. Normalization is token-aware: an external target such as `[1]DB_CHUTE!$L$47` and
quoted text such as `"AL14"` are never mistaken for a local cell, whatever the row number is.

**The queue** (`_Queue`) is diff-driven: every build recomputes the intended writes and reconciles
them into the standing actions — include choices (☑/☐, double-click) and the first-seen old value
survive, satisfied intents retire as *AppliedExternally*, vanished ones drop, moved files re-base
with the evidence in the note. In append mode a row's **device blank and the label writes it
depends on** (T / U, and the Q / R extension that displays them) form one **FLD group** — the
`Group` column reads `FLD` — and they toggle together and apply together, never half of a decision:
a blank J next to an un-appended label is the device existing nowhere. A group split outside the
gesture (a hand-edited sidecar, a member stuck in conflict) is held back from Apply and counted in
the confirmation. Everything persists in `swp-bridge.json` at the SWP root (atomic,
`.bak`), one file per system, surviving every version-folder hop.

**Double-clickable cells wear a blue border** everywhere: on the Workspace sheet,
Verdict/Coordinate open the preview, **Row opens the IO-list workbook at that card-sheet row**
and **SWP row opens the SWP workbook at the matched row** (reusing an already-open instance;
opening the SWP is for inspection — apply still refuses while it stays open); on `_Queue` the
Include toggle, on `_Glossary` the Use toggle, on `_Preview` the full-block toggle and the
**SWP row** column (opens the SWP workbook at that row).

**The preview** — the `_Preview` sheet **opens with the workspace** (already showing the first
previewable row) and its tab sits **right after Workspace**, not at the far end of the strip. The
three sheets you move between are colour-coded in the tab strip with the palette they use inside:
**Home light blue**, **Workspace orange**, **_Preview yellow**. Double-click a Verdict/Coordinate
cell (or *preview block*) to point it at another row: the scratch
sheet renders the matched device block's 32-bit section as the SWP holds it — row number, tag
type, entry, bit, every written column in the SWP's own order — Export (B), I/O addr (G), the FLD as
neighbours **FUnit (H) · Location (I) · Device (J)**, T.S. ref (L), the descriptions (T / U) — the
target row highlighted, each changing cell showing `current → new`, and the final Q/R texts
predicted from the cached values. Double-clicking the **"show full block"** toggle widens the
view to the whole device — every section, with EVERY pending update of that device shown and
labeled with the mnemonic that feeds it.

**The dashboard** — a small live window on its own thread (it keeps painting while Excel is busy):
the current phase, per-sheet tallies (activate a row to filter the Workspace sheet), the findings tail
(activate to jump to the row), and Refresh / Apply / Cancel. `_Log` keeps the full findings list
with jump links, and a `SwpBridge_Log.txt` lands next to the SWP on apply.

**Apply** — headless, never through Excel: per-action **three-way check** (the cell must still hold
the value the diff saw, else that action flags Conflict and the rest proceed); a cell that is a
**shared formula** refuses any write, whatever the column (a literal over the shared master would
orphan its dependents and Excel would repair the file on open); two actions on one cell let only the
first through; an FLD group with a conflicted member holds its other members back, noted on each.
Then the affected
sheets are stream-rewritten in the file bytes (a 35 MB CHUTE sheet never becomes a DOM; untouched
sheets are not rewritten at all). Writing `X` (the rules-§3.1 format; existing lowercase `x`
never churns) over the Export template formula converts it to a literal; missing cells are inserted in column order; the edit stamp lands in AA; the calculation
chain is dropped and full-recalc-on-load set, so the next Excel open recomputes the descriptions.
The patched workbook is **re-read and verified before anything is written**; the write is atomic
with `.bak`. Two modes, chosen in setup: **in place**, or **version folder** — the next
`NN_dd-MMM-yyyy` sibling is created, CONFIG_PLANT copied alongside (the `[1]` external link is
folder-relative, so it keeps resolving), and the COPY patched; the original is never written.
Statuses flip only after the disk write returns, and the workspace rebuilds to show the new
reality. The Cell-class SWP workbooks (`CELL_1`, `DIGITAL_ANTENNA_1`) are out of scope.

**Copy along.** Beside *new dated version folder* sits one regex — *…and copy along the files
matching* — tested case-insensitively against the file names of the current version folder. When
Apply creates the next folder, every match is copied into it next to the patched SWP and
CONFIG_PLANT, so the new folder is a complete version: with one IO list driving two SWPs,
`^SWP_02_Induction_generated_G1\.xlsx$` in the G2 session brings G1 along — and the G1 session,
reopened, follows it there. Top-level files only; Excel's `~$` lock files and `.bak` copies never,
and the SWP and CONFIG_PLANT are never copied twice. The apply message and `_Log` name every file
copied and every one that could not be — a failed copy is reported, never fatal, since the patched
SWP is what the apply exists for. A pattern that does not parse is refused in setup and copies
nothing if it reaches an apply.

To remove the whole feature: delete `src/ExcelPowerTools/IoPipeline/SwpBridge`,
`src/ExcelPowerTools.Core/IoPipeline/SwpBridge`, `tests/ExcelPowerTools.Tests/IoPipeline/SwpBridge`,
the pane entries, the ribbon group and the three `AddIn.cs` event lines.

## Project layout

```
src/ExcelPowerTools.Core   netstandard2.0 — pure, Excel-free logic (unit-test target)
src/ExcelPowerTools        net48 — the Excel-DNA add-in (COM interop, ribbon, UI, UDFs)
tests/ExcelPowerTools.Tests net8.0 — xUnit tests over the core logic
```

## Build

```
dotnet build -c Release          # builds everything, packs the .xll
dotnet test                      # runs the unit tests
```

Output (Release): `src/ExcelPowerTools/bin/Release/net48/publish/`
- **`ExcelPowerTools-AddIn64-packed.xll`** ← use this (your Excel is 64-bit)
- `ExcelPowerTools-AddIn-packed.xll` (32-bit Excel)
- **`ExcelPowerTools.Portable.exe`** — the no-install session launcher (below)

The `-packed` files are self-contained (all managed code embedded); the loose
`bin/.../net48/ExcelPowerTools-AddIn64.xll` also works but keeps its DLLs alongside.

## Install into Excel

1. Copy `ExcelPowerTools-AddIn64-packed.xll` somewhere permanent (e.g. `%APPDATA%\ExcelPowerTools\`).
2. Excel → **File ▸ Options ▸ Add-ins**; at the bottom, *Manage:* **Excel Add-ins** ▸ **Go…**
3. **Browse…**, pick the `.xll`, tick it, **OK**.
4. First load shows an Excel-DNA security prompt — allow it.
5. A **Power Tools** tab appears on the ribbon. **Tools Hub** opens the docked full catalog; Ctrl+Q
   ("Keep Values Only") works everywhere.

To update: rebuild, replace the `.xll` (Excel must be closed), reopen Excel.

## No-install alternative: the portable session launcher

For users who can't — or won't — go through the add-in install: hand out the `publish/` folder
(the exe plus the packed `.xll` files, produced together by the Release build) and have them
double-click **`ExcelPowerTools.Portable.exe`**. It starts a **dedicated Excel instance** and
loads the add-in into it via `Application.RegisterXLL` — **for that session only**: nothing is
installed, no registry, no admin rights, and their normal Excel setup and any already-open
windows are untouched. Everything works exactly as installed (ribbon, Hub, commands, bridges);
closing that Excel ends the session.

Details it handles by itself: it detects the running Excel's bitness and picks the matching
packed file (with a precise message if only the other one is present); a file carrying the
downloaded-from-internet mark is offered an unblock (the right-click ▸ Properties ▸ Unblock
action) before Excel gets to refuse it; and if the Trust Center still declines the load, it
names the exact settings to check and offers a retry. The one thing it cannot bypass is a
policy that blocks unsigned add-ins outright — that audience needs the pure-automation host
planned in `docs/portable-host-plan.md`.

## Deliberate differences from the old VBA

These were reviewed and kept on purpose (all low-impact for the tool's actual use):

- **Change Field** reads a cell's `.Value` in your Excel locale (as the VBA did). It only edits cells
  that contain the target separator, so date/number cells are left alone in practice.
- **Change Field ▸ +/- (add/subtract)** works on **whole-number** fields. A non-integer field is left
  unchanged (the VBA would coerce-and-round `1.5` +2 into `"004"`, which is lossy). Skipped cells are
  reported in the end-of-run summary.
- **`=PT.INDEXMULTI`** returns `#N/A` for an out-of-range position (like `INDEX`), instead of failing the
  whole formula with `#VALUE!` as the old unguarded UDF did.
- Per-cell error pop-ups became one **summary** at the end; the **backup** command *copies* (keeps the
  original) instead of renaming; dialog settings live in the **registry**, not a hidden sheet.

## Development notes

- All Excel calls are **late-bound** (`dynamic`) so the add-in runs against any installed Excel
  version; only the workbook-save *events* use the typed `Microsoft.Office.Interop.Excel` types
  (embedded, `EmbedInteropTypes`), because a `ref bool Cancel` parameter can't go through `dynamic`.
- Keep new pure logic in `ExcelPowerTools.Core` (and cover it with tests). If you add a Core
  `.cs` file, also add a `<Compile Include .. Link>` line for it in `ExcelPowerTools.csproj`
  so it stays embedded in the single-file `.xll`.
- Requires the **ExcelDna.AddIn** and **ExcelDna.Interop** NuGet packages (restored automatically).
```
