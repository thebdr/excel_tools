# Excel Power Tools

A rebuild of the personal VBA macro collection (formerly in `PERSONAL.XLSB`) as a proper
**Excel-DNA add-in** written in C#. It loads as a single `.xll` you enable once in Excel —
no more flaky `PERSONAL.XLSB` startup — and moves the code off VBA onto .NET.

## What maps to what

| Original VBA | New home | Notes |
|---|---|---|
| `IoListAux.ChangeField` + `ufIoListAuxParameters` | `Commands/FieldEditorCommands.cs` + `UI/FieldEditorForm.cs` | Core split/rebuild logic lives in `ExcelPowerTools.Core/FieldStringEditor.cs` (unit-tested). Per-cell error `MsgBox` is now a single end-of-run summary. Each separator box has quick-pick buttons to its right — the usual `( ) . , - _ ;` plus whitespace tokens `\s` (space), `\n`, `\r`, `\r\n`, and Clear; each colour box has 6 colour swatches. |
| `PasteMultiSelectionValues.KeepValuesOnlyV2` (Ctrl+Q) | `Commands/SelectionCommands.cs` | Same Ctrl+Q shortcut, registered by the add-in. |
| `csvImportExport.*` | `Commands/CsvCommands.cs` | Split by `;`, save with `;`, backup. |
| `ThisWorkbook` app events | `AddIn.cs` + `CsvCommands.OnWorkbook*` | The "save this .csv with `;`?" prompt on save. |
| `aPlayground.exportSheetsAsNewFile` | `Commands/WorkbookCommands.cs` | Generalised: works on the active workbook and asks how many sheets per file (the VBA had a hard-coded path). |
| `xlReferences` / `xlCells` helpers | `Interop/RangeHelpers.cs` | `LastRow`, `LastColumn`, `LastRowInColumn`, `WorksheetExists`. |
| `xlFormulas.vbaIndexMultiValues` (UDF) | `Functions/WorksheetFunctions.cs` | Worksheet function `=PT.INDEXMULTI(range, positions)`. |

The dialog's saved parameters used to live on a hidden `DataStorage` sheet; an add-in has no
workbook, so they now persist per-user in the registry under
`HKCU\Software\ExcelPowerTools\FieldEditor`.

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
5. A **Power Tools** tab appears on the ribbon. Ctrl+Q ("Keep Values Only") works everywhere.

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
