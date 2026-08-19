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

The **Power Tools** ribbon is a compact command strip, grouped by workflow rather than by implementation:

| Ribbon group | Direct actions |
|---|---|
| Power Tools | Tools Hub, Undo, Redo, Timeline |
| Compare & review | Compare revisions, Exact two-file preview, Open Review workbook |
| Delimited workspace | Open data (files/folder split button), Save file, Save all, Index |
| Transform | Keep values, Smart fill, Change field, Regex replace |
| Clean & analyze | Natural sort (ascending/descending split button), Deduplicate, Column stats |

The large coloured icons identify the Hub and the three primary workflows; secondary actions keep familiar
Office icons. Save controls are disabled until their required source sheet or workspace is active, and
Undo/Redo reflect the Power Tools timeline.

**Tools Hub** is a docked, searchable task pane with category filters for Compare, Files, Transform, Data,
and Utilities. It exposes the complete catalog—including the less frequent file operations, workbook split,
experimental IO-list bridge, developer shortcuts, About, and persistent behavior switches—without allowing
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

## Multi-revision comparison (Alpha)

**Compare revisions** is the primary comparison command on the ribbon and is also available in the Hub with
an **ALPHA** badge. It accepts 2–20 distinct local `.xlsx`/`.xlsm` revision files. Inventorying hashes the
read-only sources, proposes a filename/evidence-based order and parent graph, and then opens a setup window in
which the user can reorder revisions, edit labels and parent links, choose the master and optional trusted
baseline, and select a new `Review.xlsx` destination. The displayed order, parent links, master, and baseline
must be explicitly confirmed before the comparison starts.

The engine conservatively infers sheet identity (including supported rename evidence), headers, keys, and
logical row/column identity across adjacent lineage edges. It builds multi-revision formula/value history and
groups related changes only when identity is sufficiently supported. Ambiguities and unsupported coverage do
not become guessed matches: they produce stable, verbose findings and corrective actions on the generated
workbook's **Diagnostics** sheet. A blocked project therefore still yields a useful diagnostic Review workbook,
but no logical history is claimed for the unresolved branch.

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
per-file cap), 2,000,000 physically stored cells, 10,000,000 bounded grid cells, and 32 MiB of embedded Review
state. Review import caps the compressed package at 128 MiB and aggregate declared expansion at 768 MiB.
The grid limit uses a per-sheet content boundary from the first through last formula/value row and column.
Physically stored blank or formatting-only cells outside it are reported and do not expand alignment; blank
evidence inside it remains comparable, and every physical cell still counts toward the stored-cell parsing cap.
Unrelated and legacy entries are limited to 64 MiB; the eight canonical, workbook-listed and relationship-bound
projection worksheets may stream up to 512 MiB each, with one unique bound shared-string table up to 256 MiB.
The streaming projection is capped at 8,000,000 cells and 100,000 hyperlinks. Decision sidecars are independently
capped at 16 MiB, 10,000 events, and bounded per-event collections.
Progress, cancellation, and a privacy-safe learned elapsed-time estimate are shown during the longer
phases. Synthetic Pipeline4-shaped alignment regressions are green, but threshold tuning against the real
Pipeline4 revision corpus remains pending. Multi-revision `.xlsb` and CSV/TSV/TXT inputs, arbitrary duplicate
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
the selected sources without modifying them, compares exact stored formula/value content, and creates a
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
  Closing a workspace with unsaved source-sheet values offers Save All before closing.

## IO-list bridge (experimental)

Three Hub commands under **Experimental** join a plant CSV to an electrical IO list. They target the
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

- **Type 5 / 6 (virtual)** — not hardware I/O; their address fields are frequently stale, so they are
  excluded from matching and counted separately in the summary.
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

## IO Bridge (experimental)

A dedicated workspace for the commissioning phase, when the IO list (electrical dept), its mnemonics
(SW engineering) and the plant CSVs (machine controller) are corrected on both sides and IO-list
additions must be integrated into the CSVs. Where the pull commands compare once and read-only, the
IO Bridge **queues edits and applies them in batch to the real files** — CSVs through the byte-faithful
CSV engine, the IO-list .xlsx directly via COM, formatting preserved.

**The workspace** — *IO Bridge: open workspace* asks for one IO-list workbook, which of its
`R{rack}.S{slot}.C{card}` sheets to include, and one Inputs + one Outputs CSV. It generates a workbook
with one sheet per card: the IO-list rows in their own order (the spine, so electrical's ordering and
additions stay recognizable), each joined to its CSV record by the usual address/node keys; a `//Match`
verdict column; unmatched CSV records of that card in a *CSV-only* section below. Sheets are **rebuilt
fresh from the documents** on every open/refresh — the workbook itself is never the durable state, so
saving it produces a dead copy (reopened copies are ignored by the commands).

**Column groups** — columns are arranged group-by-group under real Excel outline groups (the +/-
buttons). The generated `_Columns` sheet lists every column with its group, the group's order and its
hidden-by-default flag (collapsed); edit it and run *IO Bridge: column config* — the configuration is
saved per user in the registry, keyed by normalized header text, and the workspace regroups.

**Editing** — type directly into IO or CSV data cells: the change is **queued, not applied**; the cell
keeps the typed value with a yellow pending highlight, and typing the original value back unqueues it.
Generated cells (headers, status columns, the queue sheet) revert on edit. *IO Bridge: row actions*
queues row-level operations for the selected row: the four **linked-field pairs** in either direction —
address (CSV `Node`/`Offset`/`Bit` ↔ IO `Bit` column), mnemonic (CSV name ↔ IO `Mnemonic`),
description (IO → CSV), normal condition (CSV `Logic` ↔ IO `Normal condition`) — plus *integrate an
IO-only signal into the CSV* (the record is composed from the row; the CSV Type is asked, since the IO
list cannot supply it), *disable/enable* (the `@`/`#` field-0 marker) and *delete the CSV row*.

**The queue** — the `_Queue` sheet lists every action in queue order, the **Include column first**:
double-click its ☑/☐ glyph to toggle (typed x/1/☑ or blank/0/☐ also count). Unchecked actions are
skipped by Apply and stay queued. Repeated edits of the same field supersede each other but keep the
**first** old value — the three-way base. The queue persists in a JSON sidecar
(`<io-list>.iobridge.json`, next to the IO list, atomically written with a `.bak` on every change), so
it survives Excel sessions and crashes; on rebuild every action **re-anchors**: file still holds the
old value → Pending (typed value re-shown, highlighted); file already holds the new value →
AppliedExternally; anything else → ConflictStale with what the file now holds.

**Apply** — *IO Bridge: apply batch* groups the included pending actions per target and applies each
with a **per-action three-way check**: the file's current value must still equal the action's old
value, or that action flags Conflict and the rest proceed. CSVs are patched headlessly and written
only after `VerifyRoundTrip` passes and the file is confirmed unchanged since reading (atomic write,
`.bak` kept); a CSV open in the delimited-file workspace is skipped ("close it there first"). The
IO-list workbook is patched via COM — rows located by the build-time position re-verified, then
`WorksheetFunction.MATCH` on the exact Bit text with a COUNTIF uniqueness check (never a whole-sheet
`Find`, which skips filter-hidden rows). If the IO list was already open in Excel it is patched live
and you are **asked** before saving; declining leaves the actions pending. Statuses flip only after a
write actually happened, and the workspace rebuilds afterwards so the sheets show the new reality.
*IO Bridge: clear history* drops Applied rows from the queue.

To remove the whole feature: delete `src/ExcelPowerTools/Experimental/IoBridge`,
`src/ExcelPowerTools.Core/Experimental/IoBridge`, `tests/ExcelPowerTools.Tests/Experimental/IoBridge`,
the pane entries and the three `AddIn.cs` event lines.

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
