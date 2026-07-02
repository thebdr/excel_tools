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
│   │   ├─ RibbonController.cs   the "Power Tools" ribbon tab
│   │   ├─ Commands/             one file per feature area (incl. DevCommands.cs)
│   │   ├─ Functions/            worksheet UDFs (=PT.*)
│   │   ├─ Interop/              COM helpers + FrameworkPaths (self-location)
│   │   ├─ UI/                   WinForms dialogs
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

## The dev loop (driven from Excel)

The add-in has a **Framework** ribbon group so you can jump into the code the way Alt+F11 does in VBA:

- **Open in VS** → opens `ExcelPowerTools.sln` in the **newest** installed Visual Studio (resolved via
  `vswhere`, so it ignores the `.sln` file association, which may point at an older VS).
- **Repo folder** / **Build output** → open those folders in Explorer.
- **About** → shows the loaded `.xll` path and the resolved repo/solution (handy diagnostics).

Typical cycle:
1. In Excel, **Framework ▸ Open in VS**.
2. Edit code. To debug, pick the **"Excel (load add-in)"** launch profile and press **F5** — VS starts a
   fresh Excel with the freshly built Debug `.xll` loaded, breakpoints live. (The profile's Excel path
   and Debug `.xll` path are in `src/ExcelPowerTools/Properties/launchSettings.json`.)
3. For a normal rebuild, **close the Excel that has the add-in loaded first** — a running `.xll` is
   file-locked, so the pack step can't overwrite it otherwise. After closing, give `EXCEL.EXE` a few
   seconds to fully exit (it can linger ~15–30s and hold the lock) before rebuilding.

## Adding a new tool

The framework grows by convention — no plumbing:

1. **A command** (menu/ribbon action): add a method to a `Commands/*.cs` class, tagged
   `[ExcelCommand(Name = "PT_Something")]` (add `ShortCut = "^k"` for a hotkey). Add a `<button>` to
   `RibbonController.GetCustomUI` and a one-line `OnSomething` callback.
2. **A worksheet function**: add a `[ExcelFunction(Name = "PT.SOMETHING")]` method in `Functions/`.
3. **Pure logic** worth testing: put it in `ExcelPowerTools.Core`, add a matching `<Compile Link>` line
   in `ExcelPowerTools.csproj` (so it stays in the single-file `.xll`), and cover it in `tests/`.
