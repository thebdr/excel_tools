#requires -Version 5.1
#requires -PSEdition Desktop

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ResultPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OwnershipPath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{32}$')]
    [string]$OwnershipToken,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-f]{32}$')]
    [string]$RunId,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^review-retention-[0-9]{6}$')]
    [string]$CaseId
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$workerSchema = "excel-power-tools.review-excel-retention.worker.v1"
$ownershipSchema = "excel-power-tools.review-excel-retention.ownership.v1"
$script:stageCode = "EPT.RETENTION.WORKER.PREFLIGHT"
$script:errorSequence = 0

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class ExcelRetentionNativeMethods
{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@

function Get-UtcTimestamp {
    [DateTimeOffset]::UtcNow.ToString("o", [Globalization.CultureInfo]::InvariantCulture)
}

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = [IO.File]::Open(
        $Path,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::Read)
    try {
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try {
            ([BitConverter]::ToString($sha256.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
        }
        finally {
            $sha256.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Write-NewUtf8Text {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $fullPath = [IO.Path]::GetFullPath($Path)
    $parent = [IO.Path]::GetDirectoryName($fullPath)
    if (-not [IO.Directory]::Exists($parent)) {
        throw "Destination directory does not exist: $parent"
    }

    $stream = [IO.File]::Open(
        $fullPath,
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write,
        [IO.FileShare]::None)
    try {
        $writer = New-Object IO.StreamWriter(
            $stream,
            (New-Object Text.UTF8Encoding($false)),
            65536,
            $true)
        try {
            $writer.Write($Text)
            $writer.Flush()
        }
        finally {
            $writer.Dispose()
        }
        $stream.Flush($true)
    }
    finally {
        $stream.Dispose()
    }
}

function Write-NewJson {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Value
    )

    Write-NewUtf8Text -Path $Path -Text ($Value | ConvertTo-Json -Depth 32)
}

function Get-ExceptionDetail {
    param([Exception]$Exception, [int]$Depth = 0)

    if ($null -eq $Exception -or $Depth -ge 8) {
        return $null
    }

    $data = [ordered]@{}
    foreach ($key in $Exception.Data.Keys) {
        $data[[string]$key] = [string]$Exception.Data[$key]
    }

    [ordered]@{
        type = $Exception.GetType().FullName
        message = $Exception.Message
        hresult = ("0x{0:X8}" -f ($Exception.HResult -band 0xffffffffL))
        source = $Exception.Source
        stackTrace = $Exception.StackTrace
        data = $data
        innerException = Get-ExceptionDetail -Exception $Exception.InnerException -Depth ($Depth + 1)
    }
}

function Get-OptionalPropertyValue {
    param(
        $Value,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $Value) { return $null }
    $property = $Value.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    $property.Value
}

function Get-ErrorDetail {
    param(
        [Parameter(Mandatory = $true)][Management.Automation.ErrorRecord]$ErrorRecord,
        [Parameter(Mandatory = $true)][string]$StableCode
    )

    $script:errorSequence++
    $invocation = $ErrorRecord.InvocationInfo
    $command = Get-OptionalPropertyValue -Value $invocation -Name "MyCommand"
    [ordered]@{
        errorId = "$RunId/$CaseId/error-{0:D4}" -f $script:errorSequence
        code = $StableCode
        occurredUtc = Get-UtcTimestamp
        stage = $script:stageCode
        fullyQualifiedErrorId = $ErrorRecord.FullyQualifiedErrorId
        category = [string]$ErrorRecord.CategoryInfo
        targetObject = if ($null -eq $ErrorRecord.TargetObject) { $null } else { [string]$ErrorRecord.TargetObject }
        sourceLocation = [ordered]@{
            scriptName = Get-OptionalPropertyValue -Value $invocation -Name "ScriptName"
            functionName = Get-OptionalPropertyValue -Value $command -Name "Name"
            lineNumber = Get-OptionalPropertyValue -Value $invocation -Name "ScriptLineNumber"
            offsetInLine = Get-OptionalPropertyValue -Value $invocation -Name "OffsetInLine"
            sourceLine = Get-OptionalPropertyValue -Value $invocation -Name "Line"
            positionMessage = Get-OptionalPropertyValue -Value $invocation -Name "PositionMessage"
            scriptStackTrace = Get-OptionalPropertyValue -Value $ErrorRecord -Name "ScriptStackTrace"
        }
        exception = Get-ExceptionDetail -Exception $ErrorRecord.Exception
    }
}

function Throw-WorkerError {
    param(
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message
    )

    $exception = New-Object InvalidOperationException("[$Code] $Message")
    $exception.Data["stableCode"] = $Code
    throw $exception
}

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-ExcelProcesses {
    @(
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
            $path = $null
            $startTimeUtc = $null
            $startTimeUtcTicks = $null
            try { $path = $_.MainModule.FileName } catch { }
            try {
                $startTimeUtc = $_.StartTime.ToUniversalTime().ToString("o", [Globalization.CultureInfo]::InvariantCulture)
                $startTimeUtcTicks = $_.StartTime.ToUniversalTime().Ticks
            }
            catch { }
            [ordered]@{
                processId = $_.Id
                processName = $_.ProcessName
                startTimeUtc = $startTimeUtc
                startTimeUtcTicks = $startTimeUtcTicks
                executablePath = $path
            }
        }
    )
}

function Get-ExecutableArchitecture {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not [IO.File]::Exists($Path)) {
        return "unknown"
    }

    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $reader = New-Object IO.BinaryReader($stream, [Text.Encoding]::UTF8, $true)
        try {
            if ($reader.ReadUInt16() -ne 0x5a4d) { return "unknown" }
            $stream.Position = 0x3c
            $peOffset = $reader.ReadInt32()
            if ($peOffset -lt 0 -or $peOffset -gt ($stream.Length - 6)) { return "unknown" }
            $stream.Position = $peOffset
            if ($reader.ReadUInt32() -ne 0x00004550) { return "unknown" }
            switch ($reader.ReadUInt16()) {
                0x014c { return "x86" }
                0x8664 { return "x64" }
                0xaa64 { return "arm64" }
                default { return "unknown" }
            }
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-ProcessSnapshot {
    param([Diagnostics.Process]$Process)

    if ($null -eq $Process) { return $null }
    try {
        $Process.Refresh()
        [ordered]@{
            capturedUtc = Get-UtcTimestamp
            workingSetBytes = $Process.WorkingSet64
            peakWorkingSetBytes = $Process.PeakWorkingSet64
            privateMemoryBytes = $Process.PrivateMemorySize64
            virtualMemoryBytes = $Process.VirtualMemorySize64
            totalProcessorMs = [Math]::Round($Process.TotalProcessorTime.TotalMilliseconds, 3)
            handleCount = $Process.HandleCount
            threadCount = $Process.Threads.Count
        }
    }
    catch {
        [ordered]@{
            capturedUtc = Get-UtcTimestamp
            unavailable = $true
            error = $_.Exception.Message
        }
    }
}

function Get-ElapsedMilliseconds {
    param([Diagnostics.Stopwatch]$Stopwatch)
    [Math]::Round($Stopwatch.Elapsed.TotalMilliseconds, 3, [MidpointRounding]::AwayFromZero)
}

function Open-ReviewWorkbook {
    param(
        [Parameter(Mandatory = $true)]$Application,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][bool]$ReadOnly
    )

    $workbooks = $null
    try {
        $workbooks = $Application.Workbooks
        $missing = [Type]::Missing
        $workbooks.Open(
            $Path,
            0,
            $ReadOnly,
            $missing,
            $missing,
            $missing,
            $true)
    }
    finally {
        Release-ComObject $workbooks
    }
}

function Assert-OpenedExpectedWorkbook {
    param(
        [Parameter(Mandatory = $true)]$Workbook,
        [Parameter(Mandatory = $true)][string]$ExpectedPath,
        [Parameter(Mandatory = $true)][bool]$ExpectedReadOnly
    )

    $actualPath = [IO.Path]::GetFullPath([string]$Workbook.FullName)
    if (-not [string]::Equals($actualPath, $ExpectedPath, [StringComparison]::OrdinalIgnoreCase)) {
        Throw-WorkerError -Code "EPT.RETENTION.EXCEL.WRONG_WORKBOOK" -Message (
            "Excel opened '$actualPath' instead of '$ExpectedPath'.")
    }
    $readWritePromotionUsed = $false
    if (-not $ExpectedReadOnly -and [bool]$Workbook.ReadOnly) {
        # Some managed Office environments open automation workbooks read-only even when Open received false.
        # ChangeFileAccess is attempted before any workbook mutation and is verified immediately afterwards.
        try {
            $Workbook.ChangeFileAccess(2) | Out-Null
        }
        catch {
            $_.Exception.Data["stableCode"] = "EPT.RETENTION.EXCEL.READWRITE_UNAVAILABLE"
            $_.Exception.Data["correctiveAction"] = (
                "Run the retention matrix in an Excel environment that permits automation workbooks to open read/write; " +
                "review Office licensing, Protected View, and mandatory sensitivity-label policy.")
            throw
        }
        $readWritePromotionUsed = $true
    }
    if ([bool]$Workbook.ReadOnly -ne $ExpectedReadOnly) {
        Throw-WorkerError -Code "EPT.RETENTION.EXCEL.READONLY_MISMATCH" -Message (
            "Workbook ReadOnly=$($Workbook.ReadOnly), expected $ExpectedReadOnly.")
    }
    $readWritePromotionUsed
}

$resolvedInput = [IO.Path]::GetFullPath($InputPath)
$resolvedResult = [IO.Path]::GetFullPath($ResultPath)
$resolvedOwnership = [IO.Path]::GetFullPath($OwnershipPath)
$workerScriptSha256AtStart = Get-Sha256 $PSCommandPath
$startedUtc = Get-UtcTimestamp
$totalStopwatch = [Diagnostics.Stopwatch]::StartNew()
$application = $null
$workbook = $null
$excelProcess = $null
$ownership = $null
$metrics = $null
$workerError = $null
$cleanupErrors = New-Object Collections.Generic.List[object]
$forcedExcelTermination = $false
$exitCode = 1

try {
    if (-not [IO.File]::Exists($resolvedInput)) {
        Throw-WorkerError -Code "EPT.RETENTION.WORKER.INPUT_MISSING" -Message "Input workbook does not exist: $resolvedInput"
    }
    if (-not [string]::Equals([IO.Path]::GetExtension($resolvedInput), ".xlsx", [StringComparison]::OrdinalIgnoreCase)) {
        Throw-WorkerError -Code "EPT.RETENTION.WORKER.INPUT_EXTENSION" -Message "Input workbook must have the .xlsx extension."
    }
    if ([IO.File]::Exists($resolvedResult) -or [IO.File]::Exists($resolvedOwnership)) {
        Throw-WorkerError -Code "EPT.RETENTION.WORKER.OUTPUT_EXISTS" -Message (
            "Worker result and ownership paths must both be new.")
    }

    $existingExcel = @(Get-ExcelProcesses)
    if ($existingExcel.Count -ne 0) {
        $exception = New-Object InvalidOperationException(
            "[EPT.RETENTION.SAFETY.EXCEL_ALREADY_RUNNING] Refusing to create an Excel COM instance because EXCEL is already running.")
        $exception.Data["stableCode"] = "EPT.RETENTION.SAFETY.EXCEL_ALREADY_RUNNING"
        $exception.Data["excelProcesses"] = ($existingExcel | ConvertTo-Json -Depth 8 -Compress)
        throw $exception
    }

    $beforeFile = Get-Item -LiteralPath $resolvedInput
    $beforeSha256 = Get-Sha256 $resolvedInput
    $activationBoundaryUtc = [DateTime]::UtcNow
    $activationStopwatch = [Diagnostics.Stopwatch]::StartNew()

    $script:stageCode = "EPT.RETENTION.EXCEL.ACTIVATE"
    $application = New-Object -ComObject Excel.Application
    $activationStopwatch.Stop()

    # Publish the recovery identity immediately after activation. Every later Excel property access is a
    # synchronous COM call and may hang; the controller must already have enough information to clean up only
    # this exact process if the worker is terminated.
    $script:stageCode = "EPT.RETENTION.EXCEL.OWNERSHIP"
    $excelHwnd = [IntPtr]([long]$application.Hwnd)
    [uint32]$excelPid = 0
    $excelWindowThreadId = [ExcelRetentionNativeMethods]::GetWindowThreadProcessId($excelHwnd, [ref]$excelPid)
    if ($excelWindowThreadId -eq 0 -or $excelPid -eq 0) {
        Throw-WorkerError -Code "EPT.RETENTION.EXCEL.PID_UNAVAILABLE" -Message (
            "Excel returned a window handle that could not be mapped to a process ID.")
    }

    $excelProcess = Get-Process -Id ([int]$excelPid) -ErrorAction Stop
    if (-not [string]::Equals($excelProcess.ProcessName, "EXCEL", [StringComparison]::OrdinalIgnoreCase)) {
        Throw-WorkerError -Code "EPT.RETENTION.EXCEL.PID_WRONG_PROCESS" -Message (
            "Excel HWND mapped to process '$($excelProcess.ProcessName)' ($excelPid).")
    }
    $excelStartedUtc = $excelProcess.StartTime.ToUniversalTime()
    if ($excelStartedUtc -lt $activationBoundaryUtc.AddSeconds(-5)) {
        Throw-WorkerError -Code "EPT.RETENTION.EXCEL.OWNERSHIP_UNPROVEN" -Message (
            "Mapped Excel process predates this worker's COM activation boundary.")
    }

    $excelExecutablePath = $excelProcess.MainModule.FileName
    $ownership = [ordered]@{
        schema = $ownershipSchema
        runId = $RunId
        caseId = $CaseId
        ownershipToken = $OwnershipToken
        publishedUtc = Get-UtcTimestamp
        workerProcessId = $PID
        processId = $excelProcess.Id
        processName = $excelProcess.ProcessName
        processStartTimeUtc = $excelStartedUtc.ToString("o", [Globalization.CultureInfo]::InvariantCulture)
        processStartTimeUtcTicks = $excelStartedUtc.Ticks
        executablePath = $excelExecutablePath
        executableArchitecture = Get-ExecutableArchitecture $excelExecutablePath
        applicationHwnd = $excelHwnd.ToInt64()
    }
    Write-NewJson -Path $resolvedOwnership -Value $ownership

    $script:stageCode = "EPT.RETENTION.EXCEL.CONFIGURE"
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.ScreenUpdating = $false
    $application.EnableEvents = $false
    $application.AskToUpdateLinks = $false
    $application.AutomationSecurity = 3

    $excelMetadata = [ordered]@{
        process = $ownership
        version = [string]$application.Version
        build = [string]$application.Build
        productCode = [string]$application.ProductCode
        operatingSystem = [string]$application.OperatingSystem
        path = [string]$application.Path
        startupPath = [string]$application.StartupPath
        libraryPath = [string]$application.LibraryPath
    }

    $memoryAfterActivation = Get-ProcessSnapshot $excelProcess

    $script:stageCode = "EPT.RETENTION.EXCEL.OPEN"
    $openStopwatch = [Diagnostics.Stopwatch]::StartNew()
    $workbook = Open-ReviewWorkbook -Application $application -Path $resolvedInput -ReadOnly $false
    $openStopwatch.Stop()
    $readWritePromotionUsed = Assert-OpenedExpectedWorkbook `
        -Workbook $workbook -ExpectedPath $resolvedInput -ExpectedReadOnly $false
    $worksheets = $null
    try {
        $worksheets = $workbook.Worksheets
        $worksheetCount = [int]$worksheets.Count
    }
    finally {
        Release-ComObject $worksheets
    }
    $workbookMetadata = [ordered]@{
        name = [string]$workbook.Name
        fullName = [string]$workbook.FullName
        fileFormat = [int]$workbook.FileFormat
        calculationVersion = [int]$workbook.CalculationVersion
        worksheetCount = $worksheetCount
        readWritePromotionUsed = [bool]$readWritePromotionUsed
    }
    $memoryAfterOpen = Get-ProcessSnapshot $excelProcess

    $script:stageCode = "EPT.RETENTION.EXCEL.SAVE"
    $saveStopwatch = [Diagnostics.Stopwatch]::StartNew()
    $workbook.Save()
    $saveStopwatch.Stop()
    if (-not [bool]$workbook.Saved) {
        Throw-WorkerError -Code "EPT.RETENTION.EXCEL.SAVE_NOT_COMMITTED" -Message (
            "Excel returned from Save but still reports the workbook as unsaved.")
    }
    $memoryAfterSave = Get-ProcessSnapshot $excelProcess

    $script:stageCode = "EPT.RETENTION.EXCEL.CLOSE_AFTER_SAVE"
    $closeAfterSaveStopwatch = [Diagnostics.Stopwatch]::StartNew()
    $workbook.Close($false)
    $closeAfterSaveStopwatch.Stop()
    Release-ComObject $workbook
    $workbook = $null

    $afterSaveFile = Get-Item -LiteralPath $resolvedInput
    $afterSaveSha256 = Get-Sha256 $resolvedInput

    $script:stageCode = "EPT.RETENTION.EXCEL.REOPEN"
    $reopenStopwatch = [Diagnostics.Stopwatch]::StartNew()
    $workbook = Open-ReviewWorkbook -Application $application -Path $resolvedInput -ReadOnly $true
    $reopenStopwatch.Stop()
    $null = Assert-OpenedExpectedWorkbook `
        -Workbook $workbook -ExpectedPath $resolvedInput -ExpectedReadOnly $true
    $memoryAfterReopen = Get-ProcessSnapshot $excelProcess

    $script:stageCode = "EPT.RETENTION.EXCEL.CLOSE_AFTER_REOPEN"
    $closeAfterReopenStopwatch = [Diagnostics.Stopwatch]::StartNew()
    $workbook.Close($false)
    $closeAfterReopenStopwatch.Stop()
    Release-ComObject $workbook
    $workbook = $null

    $afterReopenFile = Get-Item -LiteralPath $resolvedInput
    $afterReopenSha256 = Get-Sha256 $resolvedInput
    if (-not [string]::Equals($afterSaveSha256, $afterReopenSha256, [StringComparison]::Ordinal)) {
        Throw-WorkerError -Code "EPT.RETENTION.EXCEL.REOPEN_MUTATED_FILE" -Message (
            "A read-only reopen changed the workbook package hash.")
    }

    $metrics = [ordered]@{
        excel = $excelMetadata
        workbook = $workbookMetadata
        inputBefore = [ordered]@{
            bytes = $beforeFile.Length
            lastWriteUtc = $beforeFile.LastWriteTimeUtc.ToString("o", [Globalization.CultureInfo]::InvariantCulture)
            sha256 = $beforeSha256
        }
        afterSave = [ordered]@{
            bytes = $afterSaveFile.Length
            lastWriteUtc = $afterSaveFile.LastWriteTimeUtc.ToString("o", [Globalization.CultureInfo]::InvariantCulture)
            sha256 = $afterSaveSha256
        }
        afterReadOnlyReopen = [ordered]@{
            bytes = $afterReopenFile.Length
            lastWriteUtc = $afterReopenFile.LastWriteTimeUtc.ToString("o", [Globalization.CultureInfo]::InvariantCulture)
            sha256 = $afterReopenSha256
        }
        packageHashChangedOnSave = -not [string]::Equals($beforeSha256, $afterSaveSha256, [StringComparison]::Ordinal)
        readOnlyReopenPreservedPackageHash = $true
        elapsedMs = [ordered]@{
            activate = Get-ElapsedMilliseconds $activationStopwatch
            open = Get-ElapsedMilliseconds $openStopwatch
            save = Get-ElapsedMilliseconds $saveStopwatch
            closeAfterSave = Get-ElapsedMilliseconds $closeAfterSaveStopwatch
            reopenReadOnly = Get-ElapsedMilliseconds $reopenStopwatch
            closeAfterReopen = Get-ElapsedMilliseconds $closeAfterReopenStopwatch
        }
        processSnapshots = [ordered]@{
            afterActivation = $memoryAfterActivation
            afterOpen = $memoryAfterOpen
            afterSave = $memoryAfterSave
            afterReopen = $memoryAfterReopen
        }
    }
    $exitCode = 0
}
catch {
    $stableCode = $script:stageCode + ".FAILED"
    if ($_.Exception.Data.Contains("stableCode")) {
        $stableCode = [string]$_.Exception.Data["stableCode"]
    }
    $workerError = Get-ErrorDetail -ErrorRecord $_ -StableCode $stableCode
    $exitCode = 1
}
finally {
    $script:stageCode = "EPT.RETENTION.WORKER.CLEANUP"
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {
            $cleanupErrors.Add((Get-ErrorDetail -ErrorRecord $_ -StableCode "EPT.RETENTION.CLEANUP.WORKBOOK_CLOSE_FAILED"))
        }
        finally {
            Release-ComObject $workbook
            $workbook = $null
        }
    }

    if ($null -ne $application) {
        try { $application.Quit() } catch {
            $cleanupErrors.Add((Get-ErrorDetail -ErrorRecord $_ -StableCode "EPT.RETENTION.CLEANUP.EXCEL_QUIT_FAILED"))
        }
        finally {
            Release-ComObject $application
            $application = $null
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($null -ne $excelProcess) {
        try {
            if (-not $excelProcess.WaitForExit(15000)) {
                # This Process object was obtained from the exact HWND owned by this worker.
                # Kill is intentionally scoped to that retained process handle, never a name lookup.
                $excelProcess.Kill()
                $forcedExcelTermination = $true
                if (-not $excelProcess.WaitForExit(10000)) {
                    Throw-WorkerError -Code "EPT.RETENTION.CLEANUP.EXCEL_DID_NOT_EXIT" -Message (
                        "The owned Excel process did not exit after a targeted termination request.")
                }
            }
        }
        catch {
            $cleanupErrors.Add((Get-ErrorDetail -ErrorRecord $_ -StableCode "EPT.RETENTION.CLEANUP.EXCEL_PROCESS_FAILED"))
        }
        finally {
            $excelProcess.Dispose()
            $excelProcess = $null
        }
    }
}

$totalStopwatch.Stop()
if ($cleanupErrors.Count -ne 0) {
    $exitCode = 1
}
$workerScriptSha256AtEnd = Get-Sha256 $PSCommandPath
if (-not [string]::Equals(
        $workerScriptSha256AtStart,
        $workerScriptSha256AtEnd,
        [StringComparison]::OrdinalIgnoreCase)) {
    $exitCode = 1
    $cleanupErrors.Add([ordered]@{
        errorId = "$RunId/$CaseId/error-{0:D4}" -f (++$script:errorSequence)
        code = "EPT.RETENTION.WORKER.SCRIPT_CHANGED"
        occurredUtc = Get-UtcTimestamp
        stage = "EPT.RETENTION.WORKER.COMPLETE"
        expectedSha256 = $workerScriptSha256AtStart
        actualSha256 = $workerScriptSha256AtEnd
        message = "The worker script changed while this worker was running."
    })
}

$result = [ordered]@{
    schema = $workerSchema
    runId = $RunId
    caseId = $CaseId
    success = ($exitCode -eq 0)
    startedUtc = $startedUtc
    completedUtc = Get-UtcTimestamp
    totalElapsedMs = Get-ElapsedMilliseconds $totalStopwatch
    inputPath = $resolvedInput
    resultPath = $resolvedResult
    ownershipPath = $resolvedOwnership
    ownership = $ownership
    forcedExcelTermination = $forcedExcelTermination
    metrics = $metrics
    error = $workerError
    cleanupErrors = $cleanupErrors.ToArray()
    sourceLocation = [ordered]@{
        scriptPath = $PSCommandPath
        scriptSha256 = $workerScriptSha256AtStart
        scriptSha256AtStart = $workerScriptSha256AtStart
        scriptSha256AtEnd = $workerScriptSha256AtEnd
        scriptUnchangedDuringWorker = [string]::Equals(
            $workerScriptSha256AtStart,
            $workerScriptSha256AtEnd,
            [StringComparison]::OrdinalIgnoreCase)
    }
}

try {
    Write-NewJson -Path $resolvedResult -Value $result
}
catch {
    $fallback = [ordered]@{
        schema = $workerSchema
        runId = $RunId
        caseId = $CaseId
        success = $false
        code = "EPT.RETENTION.WORKER.RESULT_WRITE_FAILED"
        resultPath = $resolvedResult
        error = Get-ErrorDetail -ErrorRecord $_ -StableCode "EPT.RETENTION.WORKER.RESULT_WRITE_FAILED"
    } | ConvertTo-Json -Depth 32 -Compress
    [Console]::Error.WriteLine($fallback)
    exit 1
}

$compactResult = $result | ConvertTo-Json -Depth 32 -Compress
if ($exitCode -eq 0) {
    [Console]::Out.WriteLine($compactResult)
}
else {
    [Console]::Error.WriteLine($compactResult)
}
exit $exitCode
