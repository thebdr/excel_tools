# excel_tools

A personal C# framework for Excel tooling — the modern replacement for the VBA `PERSONAL.XLSB`
macro collection. Everything is an **Excel-DNA add-in** (a single `.xll` you enable once), written
in C# on **.NET Framework 4.8**, so it loads reliably and the code lives in git as plain text.

> **Why net48?** Late-bound (`dynamic`) Office COM — the VBA-identical `ws.Cells(r,c).Value` style —
> works cleanly only on .NET Framework. On modern .NET (8/10) dynamic COM indexers are broken. See
> [docs/add-in-reference.md](docs/add-in-reference.md).

## Layout

```
excel_tools/
├─ ExcelPowerTools.sln
├─ src/
│   ├─ ExcelPowerTools/          the add-in (COM, ribbon, UI, UDFs, dev commands)
│   │   ├─ AddIn.cs              lifecycle + app-event wiring
│   │   ├─ RibbonController.cs   compact "Power Tools" command strip
│   │   ├─ Commands/             one file per feature area (incl. DevCommands.cs)
│   │   ├─ Functions/            worksheet UDFs (=PT.*)
│   │   ├─ Interop/              COM helpers + FrameworkPaths (self-location)
│   │   ├─ UI/                   docked Tools Hub + specialist WinForms dialogs
│   │   └─ Settings/             per-user settings (registry)
│   └─ ExcelPowerTools.Core/     pure, Excel-free logic (unit-test target)
├─ tests/ExcelPowerTools.Tests/  xUnit over Core
└─ docs/add-in-reference.md      feature-by-feature reference + VBA mapping
```

## Build & test

```
dotnet build -c Release      # builds + packs the .xll
dotnet test                  # runs the unit tests
```

The 64-bit add-in to install: `src/ExcelPowerTools/bin/Release/net48/publish/ExcelPowerTools-AddIn64-packed.xll`
(enable via Excel ▸ File ▸ Options ▸ Add-ins ▸ *Manage: Excel Add-ins* ▸ Go ▸ Browse).

## Delimited-file workspace

**Open files...** accepts multiple `.csv`, `.tsv`, and separated-value `.txt` files; **Open folder...**
loads every supported file in one folder. All sources become separate Text-formatted sheets in one shared
workspace workbook. Its generated **Index** sheet links to every source and shows the full path, timestamps,
dimensions, delimiter, encoding, line endings, notes, and current edit/conflict status. Different paths with
the same filename are supported and receive unique sheet names.

Each source sheet keeps its own detected format and raw-byte fingerprint. **Save file** and **Save All** verify
that the disk file still matches the version that was opened or last saved; an external edit offers Reload,
Save a copy, or an explicit backed-up overwrite. Blank and trailing-blank records are retained through a hidden
worksheet boundary. Source associations are runtime state, so closing a workspace with unsaved source changes
prompts before the workbook is allowed to close.

## Ribbon and Tools Hub

The ribbon is intentionally limited to frequent, context-driven actions:

- **Power Tools** — the docked Hub, Undo/Redo, and Timeline.
- **Compare & review** — Compare revisions, Exact two-file preview, and validated Review workbook reopening.
- **Delimited workspace** — Open data, Save file, Save all, and Index.
- **Transform** — Keep values, Smart fill, Change field, and Regex replace.
- **Clean & analyze** — natural sort, Deduplicate, and Column stats.

**Tools Hub** opens a searchable task pane on the right side of the active Excel window. It contains the
full command catalog, including export/backup, workbook utilities, experimental IO-list tools, developer
shortcuts, and the **Track manual edits** / **Intercept native delimited opens** settings. Each Excel window
keeps its own reusable Hub pane, so the ribbon remains useful as more modules are added without turning into
an ever-growing row of buttons. Its **Comparison** workspace is now implemented: a selected `Review.xlsx` is
integrity-validated before the pane shows its project and package summary, and **Open protected copy** launches
a private read-only copy without changing the selected report. When a validated decision sidecar is present,
the pane also opens a virtualized decision ledger below the privacy-safe package summary. Reviewers can filter,
multi-select, and right-click cells or whole groups; choose the exact master or a candidate revision; mark items
equivalent, for investigation, or clear; and attach reviewer/audit-note metadata. The ribbon's **Review workbook
in Hub...** command and the Hub's **Open Review workbook** card both enter this same workflow. Decisions remain in memory until
**Save new Review revision** creates and fully revalidates a newly named workbook. The selected source remains
byte-for-byte unchanged, and the save rechecks its package hash, validated identities, and stable Windows file
occurrence on the same protected handle used to derive the revision.

**Compare revisions** is the multi-revision alpha, available as the ribbon's primary comparison action and in
the Hub. Select 2–20 distinct `.xlsx`/`.xlsm` files, review and correct the inferred order and parent links,
choose the master and optional trusted baseline, and explicitly confirm the lineage before analysis. The
sources remain read-only. Conservative sheet/header/key and row/column alignment builds logical history and
grouped changes only when the evidence is strong enough; ambiguous identity is reported in **Diagnostics**
instead of being guessed. The alignment-correction loop is implemented end to end. When analysis is incomplete,
or confirmed overrides remain active, its review window can confirm exact sheet pairs, header/no-header choices
with an optional separate first-data row, and one-to-four ordered key-column pairs, or clear any choice back to
automatic inference. Each change appends an audit event, and **Rerun analysis** reuses the retained immutable
input snapshots. The user either continues with a complete result or explicitly accepts a diagnostic result;
that accepted result creates one new, validated, macro-free `Review.xlsx` with visible changes and
integrity-checked embedded Phase 2 state, including the active correction generation and its append-only events.
A complete result also freezes its exact revision evidence, changed-subject catalog, observation matrix, and an
empty immutable decision generation in a separately hashed, manifest-bound Custom XML sidecar. Later decision
generations use a versioned append-only event chain and are reconstructed by deterministic policy replay rather
than by trusting a stored effective-state snapshot. Diagnostic-only results do not claim that decision catalog.

Current in-process limits are 768 MiB across the selected inputs, 2,000,000 stored cells, 10,000,000 bounded
grid cells, 32 MiB of embedded Review state, and a separately bounded 16 MiB/10,000-event decision sidecar.
Each sheet's grid is content-bounded by its first and last formula/value rows and columns. Physically stored blank
or formatting-only cells outside that rectangle are reported and ignored for alignment, while blank evidence
inside it is retained; all physical cells still count toward the stored-cell parsing limit.
Review import caps the compressed package at 128 MiB and aggregate declared expansion at 768 MiB. Unrelated and
legacy entries remain limited to 64 MiB; the eight canonical relationship-bound projection worksheets may stream
up to 512 MiB each, with one unique bound shared-string table up to 256 MiB. Streaming validation is limited to
8,000,000 cells and 100,000 hyperlinks.
Synthetic Pipeline4-shaped alignment regressions are green;
tuning against the real workbook corpus remains pending. Multi-revision `.xlsb` and delimited inputs, arbitrary
duplicate row/column resolution, visual snapshots, complete source-backed project resume, Excel-sheet synchronized
context menus,
the user-facing transactional merge writer, and complete archive export remain to be built. The immutable
decision-ledger engine, exact catalog, and
initial root generation are persisted and validated in complete `Review.xlsx` projects. Versioned event history
round-trips by deterministic replay, and the Hub exposes that copy-on-write boundary through a bounded virtual
ledger. A single user action may queue exact per-subject events under one run ID, and one save publishes the ordered
batch into a new, user-named Review workbook. Rich before/after snapshots, expandable per-revision history, and
complete reopening with embedded source artifacts remain pending.

The internal Phase 4 safety substrate is now present even though no merge command is exposed. It includes
fail-closed OOXML formula/value evidence, a bounded create-new publisher for exact local NTFS paths, a canonical
portable `Prepared` → `Validated` → `OutputPublished` → `AuditCommitted` journal, a separate local checkpoint
store, a deterministic restart matrix that blocks gaps, forks, missing files, and hash mismatches, and an
immutable evidence-bound merge-plan builder. The builder freezes only explicit current Formula/Value candidate
choices, their exact master preconditions and selected outputs, master locations, revision byte identities,
review counts, and deterministic decision-set/plan hashes. A pure compiler now combines compatible Formula/Value
aspects into exact per-cell formula, typed constant, or clear-content intents and records unresolved master facts
as proof obligations. A closed-construction validation receipt rejects missing, duplicate, extra, or stale proofs;
a read-only OOXML master preflight now retains the exact source occurrence, owns/hashes its bytes, proves target
preconditions, and requires the future writer to copy those owned bytes rather than reopen a path. Its portable
proofs contain no local path. Project/artifact findings, formula dependency/group, writer-envelope, full
semantic/opaque-part, publication, audit commit, and merge UI work remain. These foundations do not make the
current alpha merge-capable.
The package hashes and generation chain detect corruption and inconsistent edits; `Review.xlsx` is not digitally
signed, so they do not authenticate the file against an attacker who rewrites the entire package coherently.

**Exact two-file preview** remains available for an explicit Before/After pair. It performs positional exact
stored-content comparison for `.xlsx`/`.xlsm` or matching CSV/TSV/separated TXT files and creates the same kind
of integrity-checked macro-free Review workbook, without multi-revision alignment. Opening any Review workbook
through the add-in revalidates it and uses a protected read-only launch copy; the original report stays
untouched and remains the path to share or reopen.

## The dev loop (driven from Excel)

The Hub's developer section lets you jump into the code the way Alt+F11 does in VBA:

- **Open in VS** → opens `ExcelPowerTools.sln` in the **newest** installed Visual Studio (resolved via
  `vswhere`, so it ignores the `.sln` file association, which may point at an older VS).
- **Repo folder** / **Build output** → open those folders in Explorer.
- **About** → shows the loaded `.xll` path and the resolved repo/solution (handy diagnostics).

Typical cycle:
1. In Excel, open **Tools Hub**, search for **Visual Studio**, and choose **Open solution in Visual Studio**.
2. Edit code. To debug, pick the **"Excel (load add-in)"** launch profile and press **F5** — VS starts a
   fresh Excel with the freshly built Debug `.xll` loaded, breakpoints live. (The profile's Excel path
   and Debug `.xll` path are in `src/ExcelPowerTools/Properties/launchSettings.json`.)
3. For a normal rebuild, **close the Excel that has the add-in loaded first** — a running `.xll` is
   file-locked, so the pack step can't overwrite it otherwise. After closing, give `EXCEL.EXE` a few
   seconds to fully exit (it can linger ~15–30s and hold the lock) before rebuilding.

## Adding a new tool

The framework grows by convention — no plumbing:

1. **A command**: add a method to a `Commands/*.cs` class, tagged `[ExcelCommand(Name = "PT_Something")]`
   (add `ShortCut = "^k"` for a hotkey), then add it to `PowerToolsPane.BuildToolCatalog`. Only add a ribbon
   button and callback when the action is frequent and benefits from immediate selection context.
2. **A worksheet function**: add a `[ExcelFunction(Name = "PT.SOMETHING")]` method in `Functions/`.
3. **Pure logic** worth testing: put it in `ExcelPowerTools.Core`, add a matching `<Compile Link>` line
   in `ExcelPowerTools.csproj` (so it stays in the single-file `.xll`), and cover it in `tests/`.
