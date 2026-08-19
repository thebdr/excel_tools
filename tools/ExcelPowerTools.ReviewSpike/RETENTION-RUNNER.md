# Real-Excel Review.xlsx retention runner

`tools\run-review-excel-retention.ps1` measures whether a generated `Review.xlsx` retains its comparison state and visible projection after a real Microsoft Excel open, save, close, and read-only reopen. The default matrix is 10,000, 100,000, and 250,000 projected change rows.

This is a Windows PowerShell 5.1 harness. It is intentionally separate from normal builds and tests because it starts real Excel processes.

## Prepare and run

Build the headless validation CLI first. The runner uses the compiled DLL directly so compiler output cannot be confused with its JSON protocol.

```powershell
dotnet build tools\ExcelPowerTools.ReviewSpike\ExcelPowerTools.ReviewSpike.csproj -c Release
```

Close every Excel window and confirm that no `EXCEL.EXE` process remains. Choose an output directory that does not exist, then run:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass `
  -File tools\run-review-excel-retention.ps1 `
  -OutputDirectory C:\Temp\ExcelPowerTools-retention-20260717
```

Do not start Excel while the matrix is running. A non-empty subset can be selected for a deliberate investigation:

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass `
  -File tools\run-review-excel-retention.ps1 `
  -OutputDirectory C:\Temp\ExcelPowerTools-retention-10k `
  -Rows 10000
```

`-CliTimeoutMinutes` defaults to 30 and `-ExcelTimeoutMinutes` defaults to 120 per child process. `-ReviewSpikeDll` can select a separately built Release DLL; otherwise the runner uses `tools\ExcelPowerTools.ReviewSpike\bin\Release\net8.0\ExcelPowerTools.ReviewSpike.dll`.

## Safety boundary

- The controller refuses to begin if any `EXCEL` process exists. The worker repeats that guard immediately before COM activation. There is no override switch.
- The source workbook is never passed to Excel. It is validated, copied through `FileMode.CreateNew`, and retained byte-for-byte. Excel receives only `excel-saved.xlsx`.
- External-link updates, events, UI alerts, and VBA automation are disabled before the workbook opens.
- Every Excel tier runs in a separate Windows PowerShell worker. Immediately after COM activation, before any configuration or workbook COM call, the worker maps the Excel application window to its PID and publishes a create-new ownership handshake containing a random nonce, PID, UTC start-time ticks, executable path, and architecture.
- Normal cleanup closes only the workbook COM object and application COM object created by that worker. If the owned Excel process does not exit, the worker terminates only the retained process handle obtained from its own Excel window.
- On every abnormal worker path--timeout, crash/nonzero exit, missing or invalid result, or controller-side validation failure--the controller attempts recovery only through that handshake. It can terminate Excel only when the ownership nonce, PID, process name, start-time ticks, and executable path all still match. Cleanup outcome is attached to the original error; it never uses `Stop-Process -Name EXCEL` or closes an unproven process.
- If an unowned Excel process appears during a run, the matrix stops. That process is logged but not touched.
- The controller, worker, and compiled ReviewSpike DLL are SHA-256 pinned before the run, rechecked before and after every child, and rechecked at finalization. The worker reports its own start/end script hashes, which must match the controller's pin.
- The output directory and every result, baseline, ownership, source, and working artifact use create-new semantics. Existing output paths and reparse-point ancestors are rejected. These paths are trusted boundaries and must not be replaced while the run is active. Failure artifacts are preserved for investigation; the runner does not recursively delete any path.

Because Excel COM calls themselves are synchronous, do not use this harness on a desktop where another person or automation may start Excel concurrently. The ownership checks prevent broad cleanup, but they cannot reserve Excel globally for the duration of a run.

## Validation and result semantics

For each tier, the controller performs this sequence:

1. `ReviewSpike generate` creates and internally validates `source.xlsx`, and writes `baseline.json` with create-new semantics.
2. `ReviewSpike verify --baseline` verifies the source.
3. A create-new byte-exact working copy is made and verified against the same baseline.
4. Only after those steps succeed does the isolated worker open, save, close, and read-only reopen the working copy in Excel.
5. The untouched source is revalidated and its whole-file SHA-256 must remain identical.
6. The Excel-saved working copy is revalidated against the baseline. Project ID, declared and actual row counts, state SHA-256, projection SHA-256, visible-projection semantic SHA-256, and chunk count must match.

Excel legitimately rewrites ZIP/package bytes, so a changed whole-workbook SHA-256 after save is recorded as information rather than treated as data loss. A read-only reopen must not change the post-save package hash.

The production reader and the independent Changes inspector retain the 512 MiB streaming worksheet-part bound. Each tier is independent: any generation or baseline-verification failure is recorded with its CLI output, Excel is not invoked for that tier, and the matrix continues.

## Output layout

```text
<new-output-directory>\
  run-owner.json
  events.jsonl
  result.json
  review-retention-010000\
    source.xlsx
    baseline.json
    excel-saved.xlsx
    excel-process-ownership.json
    excel-worker-result.json
  review-retention-100000\...
  review-retention-250000\...
```

Files that were not reached because a tier failed are absent. `result.json` is the aggregate result; `events.jsonl` is the verbose chronological log. Both use stable schema and event/error codes. Every error records a unique error ID, stage, case ID, PowerShell source line and position, script stack, fully-qualified error ID, exception type/message/HRESULT/source/stack, inner exceptions, and captured child-process JSON/stdout/stderr where applicable. Worker results include Excel version, build, executable architecture, process PID, file hashes and sizes, phase timings, and process memory/CPU snapshots.
