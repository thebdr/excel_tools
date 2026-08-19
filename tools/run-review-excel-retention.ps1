#requires -Version 5.1
#requires -PSEdition Desktop

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory,

    [string]$Workspace,

    [ValidateSet(10000, 100000, 250000)]
    [int[]]$Rows = @(10000, 100000, 250000),

    [string]$ReviewSpikeDll,

    [ValidateRange(1, 240)]
    [int]$CliTimeoutMinutes = 30,

    [ValidateRange(1, 1440)]
    [int]$ExcelTimeoutMinutes = 120
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$resultSchema = "excel-power-tools.review-excel-retention.result.v1"
$eventSchema = "excel-power-tools.review-excel-retention.event.v1"
$ownerSchema = "excel-power-tools.review-excel-retention.run-owner.v1"
$workerSchema = "excel-power-tools.review-excel-retention.worker.v1"
$ownershipSchema = "excel-power-tools.review-excel-retention.ownership.v1"
$runId = [Guid]::NewGuid().ToString("N")
$script:eventSequence = 0
$script:errorSequence = 0
$script:stageCode = "EPT.RETENTION.RUN.PREFLIGHT"
$script:caseId = $null
$script:eventWriter = $null

if ([string]::IsNullOrWhiteSpace($Workspace)) {
    $Workspace = Split-Path -Parent $PSScriptRoot
}

function Get-UtcTimestamp {
    [DateTimeOffset]::UtcNow.ToString("o", [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ElapsedMilliseconds {
    param([Diagnostics.Stopwatch]$Stopwatch)
    [Math]::Round($Stopwatch.Elapsed.TotalMilliseconds, 3, [MidpointRounding]::AwayFromZero)
}

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try {
            ([BitConverter]::ToString($sha256.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
        }
        finally { $sha256.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Write-NewUtf8Text {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $stream = [IO.File]::Open(
        [IO.Path]::GetFullPath($Path),
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
        finally { $writer.Dispose() }
        $stream.Flush($true)
    }
    finally { $stream.Dispose() }
}

function Write-NewJson {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Value
    )
    Write-NewUtf8Text -Path $Path -Text ($Value | ConvertTo-Json -Depth 40)
}

function New-EventLog {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = [IO.File]::Open(
        [IO.Path]::GetFullPath($Path),
        [IO.FileMode]::CreateNew,
        [IO.FileAccess]::Write,
        [IO.FileShare]::Read)
    New-Object IO.StreamWriter(
        $stream,
        (New-Object Text.UTF8Encoding($false)),
        65536,
        $false)
}

function Write-Event {
    param(
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][ValidateSet("debug", "info", "warning", "error")][string]$Level,
        [Parameter(Mandatory = $true)][string]$Message,
        $Data = $null
    )

    if ($null -eq $script:eventWriter) { return }
    $script:eventSequence++
    $event = [ordered]@{
        schema = $eventSchema
        eventId = "$runId/event-{0:D6}" -f $script:eventSequence
        code = $Code
        level = $Level
        occurredUtc = Get-UtcTimestamp
        stage = $script:stageCode
        caseId = $script:caseId
        message = $Message
        data = $Data
        sourceLocation = [ordered]@{
            scriptPath = $PSCommandPath
            functionName = $MyInvocation.MyCommand.Name
        }
    }
    $script:eventWriter.WriteLine(($event | ConvertTo-Json -Depth 40 -Compress))
    $script:eventWriter.Flush()
}

function Get-ExceptionDetail {
    param([Exception]$Exception, [int]$Depth = 0)

    if ($null -eq $Exception -or $Depth -ge 8) { return $null }
    $data = [ordered]@{}
    foreach ($key in $Exception.Data.Keys) { $data[[string]$key] = [string]$Exception.Data[$key] }
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
        errorId = "$runId/error-{0:D4}" -f $script:errorSequence
        code = $StableCode
        occurredUtc = Get-UtcTimestamp
        stage = $script:stageCode
        caseId = $script:caseId
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

function Throw-RunnerError {
    param(
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message,
        [Collections.IDictionary]$Data
    )

    $exception = New-Object InvalidOperationException("[$Code] $Message")
    $exception.Data["stableCode"] = $Code
    if ($null -ne $Data) {
        foreach ($key in $Data.Keys) { $exception.Data[[string]$key] = [string]$Data[$key] }
    }
    throw $exception
}

function Get-PinnedArtifactStatus {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedSha256
    )

    try {
        if (-not [IO.File]::Exists($Path)) {
            return [ordered]@{
                name = $Name
                path = $Path
                expectedSha256 = $ExpectedSha256
                actualSha256 = $null
                matched = $false
                error = "Artifact is missing."
            }
        }
        $actual = Get-Sha256 $Path
        [ordered]@{
            name = $Name
            path = $Path
            expectedSha256 = $ExpectedSha256
            actualSha256 = $actual
            matched = [string]::Equals($ExpectedSha256, $actual, [StringComparison]::OrdinalIgnoreCase)
            error = $null
        }
    }
    catch {
        [ordered]@{
            name = $Name
            path = $Path
            expectedSha256 = $ExpectedSha256
            actualSha256 = $null
            matched = $false
            error = $_.Exception.GetType().FullName + ": " + $_.Exception.Message
        }
    }
}

function Get-PinnedRunArtifactIntegrity {
    $controller = Get-PinnedArtifactStatus -Name "controller" -Path $script:controllerPath `
        -ExpectedSha256 $script:expectedControllerSha256
    $worker = Get-PinnedArtifactStatus -Name "worker" -Path $script:workerPath `
        -ExpectedSha256 $script:expectedWorkerSha256
    $reviewSpike = Get-PinnedArtifactStatus -Name "reviewSpike" -Path $script:reviewSpikeDll `
        -ExpectedSha256 $script:expectedReviewSpikeSha256
    [ordered]@{
        capturedUtc = Get-UtcTimestamp
        allMatched = [bool]($controller.matched -and $worker.matched -and $reviewSpike.matched)
        controller = $controller
        worker = $worker
        reviewSpike = $reviewSpike
    }
}

function Assert-PinnedRunArtifacts {
    param([Parameter(Mandatory = $true)][string]$Code)

    $integrity = Get-PinnedRunArtifactIntegrity
    if (-not $integrity.allMatched) {
        Throw-RunnerError -Code $Code -Message (
            "A pinned controller, worker, or ReviewSpike artifact changed or became unavailable.") -Data @{
                artifactIntegrity = ($integrity | ConvertTo-Json -Depth 12 -Compress)
            }
    }
    $integrity
}

function Assert-NoReparsePointAncestors {
    param(
        [Parameter(Mandatory = $true)][string]$ExistingPath,
        [Parameter(Mandatory = $true)][string]$Code
    )

    $current = Get-Item -LiteralPath ([IO.Path]::GetFullPath($ExistingPath)) -Force
    while ($null -ne $current) {
        if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            Throw-RunnerError -Code $Code -Message (
                "A retention-run artifact path contains a reparse point: $($current.FullName)")
        }
        if ($current -is [IO.DirectoryInfo]) {
            $current = $current.Parent
        }
        elseif ($current -is [IO.FileInfo]) {
            $current = $current.Directory
        }
        else {
            $current = $null
        }
    }
}

function Get-ExcelProcesses {
    @(
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
            $path = $null
            $started = $null
            $ticks = $null
            try { $path = $_.MainModule.FileName } catch { }
            try {
                $started = $_.StartTime.ToUniversalTime().ToString("o", [Globalization.CultureInfo]::InvariantCulture)
                $ticks = $_.StartTime.ToUniversalTime().Ticks
            }
            catch { }
            [ordered]@{
                processId = $_.Id
                processName = $_.ProcessName
                startTimeUtc = $started
                startTimeUtcTicks = $ticks
                executablePath = $path
            }
        }
    )
}

function Assert-NoExcelProcess {
    param([Parameter(Mandatory = $true)][string]$Code)

    $processes = @(Get-ExcelProcesses)
    if ($processes.Count -ne 0) {
        Throw-RunnerError -Code $Code -Message (
            "Refusing to continue because one or more EXCEL processes are running. " +
            "No existing Excel process will be attached, closed, or terminated.") -Data @{
                excelProcesses = ($processes | ConvertTo-Json -Depth 8 -Compress)
            }
    }
}

function ConvertTo-NativeArgument {
    param([AllowEmptyString()][string]$Value)

    if ($Value -notmatch '[\s"]') { return $Value }
    $builder = New-Object Text.StringBuilder
    [void]$builder.Append('"')
    $backslashes = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') {
            $backslashes++
            continue
        }
        if ($character -eq '"') {
            [void]$builder.Append(('\' * (($backslashes * 2) + 1)))
            [void]$builder.Append('"')
            $backslashes = 0
            continue
        }
        if ($backslashes -ne 0) {
            [void]$builder.Append(('\' * $backslashes))
            $backslashes = 0
        }
        [void]$builder.Append($character)
    }
    if ($backslashes -ne 0) { [void]$builder.Append(('\' * ($backslashes * 2))) }
    [void]$builder.Append('"')
    $builder.ToString()
}

function Invoke-OwnedProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$WorkingDirectory,
        [Parameter(Mandatory = $true)][int]$TimeoutMinutes
    )

    $argumentText = (($Arguments | ForEach-Object { ConvertTo-NativeArgument ([string]$_) }) -join " ")
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FileName
    $startInfo.Arguments = $argumentText
    $startInfo.WorkingDirectory = $WorkingDirectory
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object Diagnostics.Process
    $process.StartInfo = $startInfo
    $startedUtc = Get-UtcTimestamp
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    try {
        if (-not $process.Start()) { Throw-RunnerError -Code "EPT.RETENTION.PROCESS.START_FAILED" -Message "Child process did not start." }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $timedOut = -not $process.WaitForExit($TimeoutMinutes * 60 * 1000)
        $forcedTermination = $false
        if ($timedOut) {
            # This exact Process object was created above; no name-based termination is used.
            $process.Kill()
            $forcedTermination = $true
            if (-not $process.WaitForExit(10000)) {
                Throw-RunnerError -Code "EPT.RETENTION.PROCESS.DID_NOT_EXIT" -Message (
                    "The exact child process did not exit after a targeted termination request.")
            }
        }
        else {
            $process.WaitForExit()
        }
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        $stopwatch.Stop()
        [ordered]@{
            fileName = $FileName
            arguments = @($Arguments)
            processId = $process.Id
            startedUtc = $startedUtc
            completedUtc = Get-UtcTimestamp
            elapsedMs = Get-ElapsedMilliseconds $stopwatch
            timedOut = $timedOut
            forcedTermination = $forcedTermination
            exitCode = if ($timedOut) { $null } else { $process.ExitCode }
            stdout = $stdout
            stderr = $stderr
        }
    }
    finally { $process.Dispose() }
}

function ConvertFrom-ChildJson {
    param([Parameter(Mandatory = $true)][string]$Text, [Parameter(Mandatory = $true)][string]$Code)

    $trimmed = $Text.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        Throw-RunnerError -Code $Code -Message "Child process returned no JSON output."
    }
    try { $trimmed | ConvertFrom-Json }
    catch {
        Throw-RunnerError -Code $Code -Message "Child process output was not one JSON document." -Data @{
            parseError = $_.Exception.Message
            output = $trimmed
        }
    }
}

function Invoke-ReviewSpike {
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$OperationCode
    )

    $prelaunchIntegrity = Assert-PinnedRunArtifacts -Code "EPT.RETENTION.ARTIFACT.DRIFT_BEFORE_CHILD"
    Write-Event -Code "$OperationCode.STARTED" -Level info -Message "ReviewSpike command started." -Data @{
        arguments = $Arguments
        artifactIntegrity = $prelaunchIntegrity
    }
    $child = Invoke-OwnedProcess -FileName $script:dotnetPath -Arguments (@($script:reviewSpikeDll) + $Arguments) `
        -WorkingDirectory $script:workspacePath -TimeoutMinutes $CliTimeoutMinutes
    $postChildIntegrity = Assert-PinnedRunArtifacts -Code "EPT.RETENTION.ARTIFACT.DRIFT_AFTER_CHILD"
    $payloadText = if ([string]::IsNullOrWhiteSpace($child.stdout)) { $child.stderr } else { $child.stdout }
    $payload = $null
    if (-not [string]::IsNullOrWhiteSpace($payloadText)) {
        try { $payload = $payloadText.Trim() | ConvertFrom-Json } catch { }
    }
    Write-Event -Code "$OperationCode.COMPLETED" -Level $(if ($child.exitCode -eq 0) { "info" } else { "error" }) `
        -Message "ReviewSpike command completed." -Data @{
            process = $child
            payload = $payload
            artifactIntegrity = $postChildIntegrity
        }
    if ($child.timedOut) {
        Throw-RunnerError -Code "$OperationCode.TIMEOUT" -Message "ReviewSpike exceeded its timeout." -Data @{
            process = ($child | ConvertTo-Json -Depth 20 -Compress)
        }
    }
    if ($child.exitCode -ne 0) {
        Throw-RunnerError -Code "$OperationCode.FAILED" -Message (
            "ReviewSpike exited with code $($child.exitCode).") -Data @{
                process = ($child | ConvertTo-Json -Depth 20 -Compress)
                cliPayload = ($payload | ConvertTo-Json -Depth 20 -Compress)
            }
    }
    if ($null -eq $payload) {
        $payload = ConvertFrom-ChildJson -Text $child.stdout -Code "$OperationCode.INVALID_JSON"
    }
    [ordered]@{ process = $child; metrics = $payload }
}

function Copy-CreateNew {
    param([Parameter(Mandatory = $true)][string]$Source, [Parameter(Mandatory = $true)][string]$Destination)

    $input = [IO.File]::Open($Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $output = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $input.CopyTo($output, 1024 * 1024)
            $output.Flush($true)
        }
        finally { $output.Dispose() }
    }
    finally { $input.Dispose() }
}

function Get-MetricValue {
    param($Metrics, [Parameter(Mandatory = $true)][string]$Name)

    $property = $Metrics.PSObject.Properties[$Name]
    if ($null -eq $property) {
        Throw-RunnerError -Code "EPT.RETENTION.CLI.METRIC_MISSING" -Message "ReviewSpike metric '$Name' is missing."
    }
    $property.Value
}

function Get-SemanticInvariants {
    param([Parameter(Mandatory = $true)]$Metrics)

    [ordered]@{
        projectId = [string](Get-MetricValue $Metrics "projectId")
        rowCount = [long](Get-MetricValue $Metrics "rowCount")
        declaredRowCount = [long](Get-MetricValue $Metrics "declaredRowCount")
        stateSha256 = [string](Get-MetricValue $Metrics "stateSha256")
        projectionSha256 = [string](Get-MetricValue $Metrics "projectionSha256")
        visibleProjectionSemanticSha256 = [string](Get-MetricValue $Metrics "visibleProjectionSemanticSha256")
        chunkCount = [int](Get-MetricValue $Metrics "chunkCount")
    }
}

function Assert-SemanticInvariantsEqual {
    param($Before, $After)

    foreach ($name in @(
        "projectId",
        "rowCount",
        "declaredRowCount",
        "stateSha256",
        "projectionSha256",
        "visibleProjectionSemanticSha256",
        "chunkCount")) {
        if (-not [object]::Equals($Before[$name], $After[$name])) {
            Throw-RunnerError -Code "EPT.RETENTION.VALIDATION.SEMANTIC_MISMATCH" -Message (
                "Semantic invariant '$name' changed from '$($Before[$name])' to '$($After[$name])'.")
        }
    }
}

function Stop-OwnedExcelFromHandshake {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedToken,
        [Parameter(Mandatory = $true)][string]$ExpectedCaseId
    )

    if (-not [IO.File]::Exists($Path)) {
        return [ordered]@{ ownershipPublished = $false; processFound = $false; terminated = $false }
    }
    $ownership = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    if ($ownership.schema -ne $ownershipSchema -or $ownership.runId -ne $runId -or
        $ownership.caseId -ne $ExpectedCaseId -or $ownership.ownershipToken -ne $ExpectedToken) {
        Throw-RunnerError -Code "EPT.RETENTION.CLEANUP.OWNERSHIP_MISMATCH" -Message (
            "Refusing targeted cleanup because the Excel ownership handshake is not authentic for this case.")
    }

    $process = Get-Process -Id ([int]$ownership.processId) -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        return [ordered]@{ ownershipPublished = $true; processFound = $false; terminated = $false }
    }
    try {
        $actualPath = $process.MainModule.FileName
        $actualTicks = $process.StartTime.ToUniversalTime().Ticks
        if ($process.ProcessName -ne "EXCEL" -or $actualTicks -ne [long]$ownership.processStartTimeUtcTicks -or
            -not [string]::Equals($actualPath, [string]$ownership.executablePath, [StringComparison]::OrdinalIgnoreCase)) {
            Throw-RunnerError -Code "EPT.RETENTION.CLEANUP.PROCESS_IDENTITY_MISMATCH" -Message (
                "Refusing targeted cleanup because PID ownership no longer matches the handshake.")
        }
        $process.Kill()
        if (-not $process.WaitForExit(10000)) {
            Throw-RunnerError -Code "EPT.RETENTION.CLEANUP.EXCEL_DID_NOT_EXIT" -Message (
                "The proven-owned Excel process did not exit after targeted termination.")
        }
        [ordered]@{ ownershipPublished = $true; processFound = $true; terminated = $true; processId = $process.Id }
    }
    finally { $process.Dispose() }
}

function Invoke-VerifiedExcelFailureCleanup {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedToken,
        [Parameter(Mandatory = $true)][string]$ExpectedCaseId
    )

    try {
        [ordered]@{
            attemptedUtc = Get-UtcTimestamp
            succeeded = $true
            result = Stop-OwnedExcelFromHandshake -Path $Path -ExpectedToken $ExpectedToken `
                -ExpectedCaseId $ExpectedCaseId
            error = $null
        }
    }
    catch {
        [ordered]@{
            attemptedUtc = Get-UtcTimestamp
            succeeded = $false
            result = $null
            error = Get-ErrorDetail -ErrorRecord $_ -StableCode "EPT.RETENTION.CLEANUP.VERIFIED_EXCEL_FAILED"
        }
    }
}

function Invoke-ExcelWorker {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$WorkerResultPath,
        [Parameter(Mandatory = $true)][string]$OwnershipPath,
        [Parameter(Mandatory = $true)][string]$OwnershipToken
    )

    $arguments = @(
        "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
        "-File", $script:workerPath,
        "-InputPath", $WorkbookPath,
        "-ResultPath", $WorkerResultPath,
        "-OwnershipPath", $OwnershipPath,
        "-OwnershipToken", $OwnershipToken,
        "-RunId", $runId,
        "-CaseId", $script:caseId)

    $child = $null
    try {
        $prelaunchIntegrity = Assert-PinnedRunArtifacts -Code "EPT.RETENTION.ARTIFACT.DRIFT_BEFORE_WORKER"
        Write-Event -Code "EPT.RETENTION.EXCEL.WORKER_STARTED" -Level info -Message (
            "Isolated Excel worker started.") -Data @{ artifactIntegrity = $prelaunchIntegrity }
        $child = Invoke-OwnedProcess -FileName $script:windowsPowerShellPath -Arguments $arguments `
            -WorkingDirectory $script:workspacePath -TimeoutMinutes $ExcelTimeoutMinutes
        $postChildIntegrity = Assert-PinnedRunArtifacts -Code "EPT.RETENTION.ARTIFACT.DRIFT_AFTER_WORKER"
        Write-Event -Code "EPT.RETENTION.EXCEL.WORKER_COMPLETED" `
            -Level $(if ($child.exitCode -eq 0) { "info" } else { "error" }) `
            -Message "Isolated Excel worker completed." -Data @{
                process = $child
                artifactIntegrity = $postChildIntegrity
            }
        if ($child.timedOut) {
            Throw-RunnerError -Code "EPT.RETENTION.EXCEL.WORKER_TIMEOUT" -Message (
                "Excel worker exceeded its timeout.") -Data @{
                    process = ($child | ConvertTo-Json -Depth 20 -Compress)
                }
        }
        if (-not [IO.File]::Exists($WorkerResultPath)) {
            Throw-RunnerError -Code "EPT.RETENTION.EXCEL.WORKER_RESULT_MISSING" -Message (
                "Excel worker did not create its result with create-new semantics.") -Data @{
                    process = ($child | ConvertTo-Json -Depth 20 -Compress)
                }
        }
        $workerResult = Get-Content -LiteralPath $WorkerResultPath -Raw | ConvertFrom-Json
        if ($workerResult.schema -ne $workerSchema -or $workerResult.runId -ne $runId -or
            $workerResult.caseId -ne $script:caseId) {
            Throw-RunnerError -Code "EPT.RETENTION.EXCEL.WORKER_RESULT_IDENTITY" -Message (
                "Worker result identity is invalid.")
        }
        if ($null -eq $workerResult.ownership -or $workerResult.ownership.ownershipToken -ne $OwnershipToken -or
            [int]$workerResult.ownership.workerProcessId -ne [int]$child.processId) {
            Throw-RunnerError -Code "EPT.RETENTION.EXCEL.WORKER_OWNERSHIP_IDENTITY" -Message (
                "Worker result does not contain the expected ownership token and worker process identity.")
        }

        $sourceProperty = $workerResult.PSObject.Properties["sourceLocation"]
        $reportedSource = if ($null -eq $sourceProperty) { $null } else { $sourceProperty.Value }
        $reportedPath = if ($null -eq $reportedSource) { "" } else { [string]$reportedSource.scriptPath }
        $reportedStartHash = if ($null -eq $reportedSource) { "" } else {
            [string]$reportedSource.scriptSha256AtStart
        }
        $reportedEndHash = if ($null -eq $reportedSource) { "" } else {
            [string]$reportedSource.scriptSha256AtEnd
        }
        if (-not [string]::Equals($reportedPath, $script:workerPath, [StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals(
                $reportedStartHash,
                $script:expectedWorkerSha256,
                [StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals(
                $reportedEndHash,
                $script:expectedWorkerSha256,
                [StringComparison]::OrdinalIgnoreCase)) {
            Throw-RunnerError -Code "EPT.RETENTION.EXCEL.WORKER_SCRIPT_IDENTITY" -Message (
                "Worker-reported script path or hash differs from the pinned worker artifact.") -Data @{
                    reportedPath = $reportedPath
                    reportedStartSha256 = $reportedStartHash
                    reportedEndSha256 = $reportedEndHash
                    expectedPath = $script:workerPath
                    expectedSha256 = $script:expectedWorkerSha256
                }
        }
        if ($child.exitCode -ne 0 -or -not [bool]$workerResult.success) {
            $workerError = Get-OptionalPropertyValue -Value $workerResult -Name "error"
            $workerCode = [string](Get-OptionalPropertyValue -Value $workerError -Name "code")
            if ($workerCode -notmatch '^EPT\.RETENTION\.[A-Z0-9_.-]{1,120}$') {
                $workerCode = "EPT.RETENTION.EXCEL.WORKER_FAILED"
            }
            Throw-RunnerError -Code $workerCode -Message (
                "Excel worker reported a verified structured failure.") -Data @{
                    process = ($child | ConvertTo-Json -Depth 20 -Compress)
                    workerResult = ($workerResult | ConvertTo-Json -Depth 30 -Compress)
                }
        }

        $postWorkerCleanupCheck = Stop-OwnedExcelFromHandshake -Path $OwnershipPath `
            -ExpectedToken $OwnershipToken -ExpectedCaseId $script:caseId
        if ($postWorkerCleanupCheck.processFound) {
            Throw-RunnerError -Code "EPT.RETENTION.EXCEL.WORKER_LEFT_OWNED_PROCESS" -Message (
                "The worker reported success but left its proven-owned Excel process running; it was terminated.") `
                -Data @{ cleanup = ($postWorkerCleanupCheck | ConvertTo-Json -Depth 12 -Compress) }
        }
        [ordered]@{
            process = $child
            result = $workerResult
            postWorkerCleanupCheck = $postWorkerCleanupCheck
        }
    }
    catch {
        $originalError = $_
        $failureCleanup = Invoke-VerifiedExcelFailureCleanup -Path $OwnershipPath `
            -ExpectedToken $OwnershipToken -ExpectedCaseId $script:caseId
        try {
            $originalError.Exception.Data["verifiedExcelFailureCleanup"] =
                ($failureCleanup | ConvertTo-Json -Depth 20 -Compress)
        }
        catch { }
        try {
            Write-Event -Code "EPT.RETENTION.EXCEL.FAILURE_CLEANUP" `
                -Level $(if ($failureCleanup.succeeded) { "warning" } else { "error" }) `
                -Message "Verified Excel cleanup was attempted after an abnormal worker path." `
                -Data $failureCleanup
        }
        catch {
            try { $originalError.Exception.Data["failureCleanupLogError"] = $_.Exception.Message } catch { }
        }
        throw
    }
}

$controllerPath = [IO.Path]::GetFullPath($PSCommandPath)
$workspacePath = [IO.Path]::GetFullPath($Workspace)
$outputPath = [IO.Path]::GetFullPath($OutputDirectory)
$outputParent = [IO.Path]::GetDirectoryName($outputPath.TrimEnd([IO.Path]::DirectorySeparatorChar))
$workerPath = Join-Path $workspacePath "tools\ExcelPowerTools.ReviewSpike\ReviewExcelRetentionWorker.ps1"
if ([string]::IsNullOrWhiteSpace($ReviewSpikeDll)) {
    $ReviewSpikeDll = Join-Path $workspacePath "tools\ExcelPowerTools.ReviewSpike\bin\Release\net8.0\ExcelPowerTools.ReviewSpike.dll"
}
$reviewSpikeDll = [IO.Path]::GetFullPath($ReviewSpikeDll)
$dotnetCommand = Get-Command dotnet -CommandType Application -ErrorAction Stop
$dotnetPath = $dotnetCommand.Source
$windowsPowerShellPath = Join-Path $PSHOME "powershell.exe"
$script:workspacePath = $workspacePath
$script:controllerPath = $controllerPath
$script:workerPath = $workerPath
$script:reviewSpikeDll = $reviewSpikeDll
$script:dotnetPath = $dotnetPath
$script:windowsPowerShellPath = $windowsPowerShellPath

if (-not [IO.Directory]::Exists($workspacePath)) { throw "Workspace does not exist: $workspacePath" }
if (-not [IO.Directory]::Exists($outputParent)) { throw "Output parent directory must already exist: $outputParent" }
if ([IO.Directory]::Exists($outputPath) -or [IO.File]::Exists($outputPath)) { throw "Output path must not exist: $outputPath" }
if (-not [IO.File]::Exists($workerPath)) { throw "Excel retention worker is missing: $workerPath" }
if (-not [IO.File]::Exists($reviewSpikeDll)) {
    throw "ReviewSpike Release DLL is missing. Build it first: dotnet build tools\ExcelPowerTools.ReviewSpike -c Release"
}
if (-not [IO.File]::Exists($windowsPowerShellPath)) { throw "Windows PowerShell is missing: $windowsPowerShellPath" }
if (@($Rows).Count -eq 0) { throw "Rows must contain at least one scale tier." }
if (@($Rows | Select-Object -Unique).Count -ne $Rows.Count) { throw "Rows must not contain duplicates." }

Assert-NoReparsePointAncestors -ExistingPath $outputParent `
    -Code "EPT.RETENTION.PATH.OUTPUT_REPARSE_POINT"
Assert-NoReparsePointAncestors -ExistingPath $controllerPath `
    -Code "EPT.RETENTION.PATH.CONTROLLER_REPARSE_POINT"
Assert-NoReparsePointAncestors -ExistingPath $workerPath `
    -Code "EPT.RETENTION.PATH.WORKER_REPARSE_POINT"
Assert-NoReparsePointAncestors -ExistingPath $reviewSpikeDll `
    -Code "EPT.RETENTION.PATH.REVIEWSPIKE_REPARSE_POINT"

$script:expectedControllerSha256 = Get-Sha256 $controllerPath
$script:expectedWorkerSha256 = Get-Sha256 $workerPath
$script:expectedReviewSpikeSha256 = Get-Sha256 $reviewSpikeDll

[void][IO.Directory]::CreateDirectory($outputPath)
$runOwnerPath = Join-Path $outputPath "run-owner.json"
$eventsPath = Join-Path $outputPath "events.jsonl"
$resultPath = Join-Path $outputPath "result.json"
$script:eventWriter = New-EventLog $eventsPath
$cases = New-Object Collections.Generic.List[object]
$errors = New-Object Collections.Generic.List[object]
$startedUtc = Get-UtcTimestamp
$totalStopwatch = [Diagnostics.Stopwatch]::StartNew()
$runExitCode = 0
$safetyAbort = $false

$runOwner = [ordered]@{
    schema = $ownerSchema
    runId = $runId
    createdUtc = $startedUtc
    controllerProcessId = $PID
    workspacePath = $workspacePath
    outputDirectory = $outputPath
    rows = @($Rows)
    script = [ordered]@{ path = $controllerPath; sha256 = $script:expectedControllerSha256 }
    worker = [ordered]@{ path = $workerPath; sha256 = $script:expectedWorkerSha256 }
    reviewSpike = [ordered]@{ path = $reviewSpikeDll; sha256 = $script:expectedReviewSpikeSha256 }
}
Write-NewJson -Path $runOwnerPath -Value $runOwner

try {
    Write-Event -Code "EPT.RETENTION.RUN.STARTED" -Level info -Message "Excel retention matrix started." -Data $runOwner
    Assert-NoExcelProcess -Code "EPT.RETENTION.SAFETY.EXCEL_ALREADY_RUNNING"

    foreach ($rowCount in $Rows) {
        $script:caseId = "review-retention-{0:D6}" -f $rowCount
        $caseStartedUtc = Get-UtcTimestamp
        $caseStopwatch = [Diagnostics.Stopwatch]::StartNew()
        $script:stageCode = "EPT.RETENTION.CASE.START"
        $caseDirectory = Join-Path $outputPath $script:caseId
        [void][IO.Directory]::CreateDirectory($caseDirectory)
        $sourcePath = Join-Path $caseDirectory "source.xlsx"
        $baselinePath = Join-Path $caseDirectory "baseline.json"
        $excelPath = Join-Path $caseDirectory "excel-saved.xlsx"
        $workerResultPath = Join-Path $caseDirectory "excel-worker-result.json"
        $ownershipPath = Join-Path $caseDirectory "excel-process-ownership.json"
        $ownershipToken = [Guid]::NewGuid().ToString("N")
        $generate = $null
        $sourceBefore = $null
        $workingBefore = $null
        $worker = $null
        $workingAfter = $null
        $sourceAfter = $null
        $sourceShaBefore = $null
        $sourceShaAfter = $null
        $excelWorkerStarted = $false

        Write-Event -Code "EPT.RETENTION.CASE.STARTED" -Level info -Message "Retention tier started." -Data @{ rowCount = $rowCount }
        try {
            Assert-NoExcelProcess -Code "EPT.RETENTION.SAFETY.EXCEL_PRESENT_BEFORE_TIER"

            $script:stageCode = "EPT.RETENTION.CLI.GENERATE"
            $generate = Invoke-ReviewSpike -Arguments @(
                "generate", "--output", $sourcePath, "--rows", [string]$rowCount,
                "--project-id", $script:caseId, "--baseline-output", $baselinePath) `
                -OperationCode "EPT.RETENTION.CLI.GENERATE"

            $script:stageCode = "EPT.RETENTION.CLI.SOURCE_VERIFY_BEFORE"
            $sourceBefore = Invoke-ReviewSpike -Arguments @(
                "verify", "--input", $sourcePath, "--baseline", $baselinePath) `
                -OperationCode "EPT.RETENTION.CLI.SOURCE_VERIFY_BEFORE"
            $sourceShaBefore = Get-Sha256 $sourcePath

            $script:stageCode = "EPT.RETENTION.COPY"
            $copyStopwatch = [Diagnostics.Stopwatch]::StartNew()
            Copy-CreateNew -Source $sourcePath -Destination $excelPath
            $copyStopwatch.Stop()
            if (-not [string]::Equals($sourceShaBefore, (Get-Sha256 $excelPath), [StringComparison]::Ordinal)) {
                Throw-RunnerError -Code "EPT.RETENTION.COPY.HASH_MISMATCH" -Message "Create-new working copy differs from source."
            }

            $script:stageCode = "EPT.RETENTION.CLI.WORKING_VERIFY_BEFORE"
            $workingBefore = Invoke-ReviewSpike -Arguments @(
                "verify", "--input", $excelPath, "--baseline", $baselinePath) `
                -OperationCode "EPT.RETENTION.CLI.WORKING_VERIFY_BEFORE"
            $beforeInvariants = Get-SemanticInvariants $workingBefore.metrics

            # Excel is never invoked unless generation and both pre-Excel validations succeeded.
            $script:stageCode = "EPT.RETENTION.EXCEL.WORKER"
            Assert-NoExcelProcess -Code "EPT.RETENTION.SAFETY.EXCEL_PRESENT_BEFORE_WORKER"
            $excelWorkerStarted = $true
            $worker = Invoke-ExcelWorker -WorkbookPath $excelPath -WorkerResultPath $workerResultPath `
                -OwnershipPath $ownershipPath -OwnershipToken $ownershipToken
            Assert-NoExcelProcess -Code "EPT.RETENTION.SAFETY.EXCEL_PRESENT_AFTER_WORKER"

            $script:stageCode = "EPT.RETENTION.CLI.SOURCE_VERIFY_AFTER"
            $sourceAfter = Invoke-ReviewSpike -Arguments @(
                "verify", "--input", $sourcePath, "--baseline", $baselinePath) `
                -OperationCode "EPT.RETENTION.CLI.SOURCE_VERIFY_AFTER"
            $sourceShaAfter = Get-Sha256 $sourcePath
            if (-not [string]::Equals($sourceShaBefore, $sourceShaAfter, [StringComparison]::Ordinal)) {
                Throw-RunnerError -Code "EPT.RETENTION.SOURCE.MUTATED" -Message "The source artifact changed during the tier."
            }

            $script:stageCode = "EPT.RETENTION.CLI.WORKING_VERIFY_AFTER"
            $workingAfter = Invoke-ReviewSpike -Arguments @(
                "verify", "--input", $excelPath, "--baseline", $baselinePath) `
                -OperationCode "EPT.RETENTION.CLI.WORKING_VERIFY_AFTER"
            $afterInvariants = Get-SemanticInvariants $workingAfter.metrics
            Assert-SemanticInvariantsEqual -Before $beforeInvariants -After $afterInvariants

            $caseStopwatch.Stop()
            $caseResult = [ordered]@{
                caseId = $script:caseId
                rowCount = $rowCount
                status = "succeeded"
                startedUtc = $caseStartedUtc
                completedUtc = Get-UtcTimestamp
                elapsedMs = Get-ElapsedMilliseconds $caseStopwatch
                artifacts = [ordered]@{
                    source = $sourcePath
                    baseline = $baselinePath
                    excelSaved = $excelPath
                    workerResult = $workerResultPath
                    processOwnership = $ownershipPath
                }
                sourcePreservedByteForByte = $true
                semanticRetention = $true
                wholePackageHashChangeIsInformational = $true
                semanticInvariants = $afterInvariants
                copyElapsedMs = Get-ElapsedMilliseconds $copyStopwatch
                generate = $generate
                sourceBefore = $sourceBefore
                workingBefore = $workingBefore
                excelWorker = $worker
                workingAfter = $workingAfter
                sourceAfter = $sourceAfter
                sourceSha256Before = $sourceShaBefore
                sourceSha256After = $sourceShaAfter
            }
            $cases.Add($caseResult)
            Write-Event -Code "EPT.RETENTION.CASE.SUCCEEDED" -Level info -Message "Retention tier succeeded." -Data $caseResult
        }
        catch {
            $caseStopwatch.Stop()
            $stableCode = $script:stageCode + ".FAILED"
            if ($_.Exception.Data.Contains("stableCode")) { $stableCode = [string]$_.Exception.Data["stableCode"] }
            $detail = Get-ErrorDetail -ErrorRecord $_ -StableCode $stableCode
            $errors.Add($detail)
            $caseResult = [ordered]@{
                caseId = $script:caseId
                rowCount = $rowCount
                status = "failed"
                startedUtc = $caseStartedUtc
                completedUtc = Get-UtcTimestamp
                elapsedMs = Get-ElapsedMilliseconds $caseStopwatch
                failedStage = $script:stageCode
                error = $detail
                artifacts = @(
                    Get-ChildItem -LiteralPath $caseDirectory -File -ErrorAction SilentlyContinue | ForEach-Object {
                        [ordered]@{ path = $_.FullName; bytes = $_.Length; lastWriteUtc = $_.LastWriteTimeUtc.ToString("o") }
                    })
                excelWasInvoked = $excelWorkerStarted
            }
            $cases.Add($caseResult)
            Write-Event -Code $stableCode -Level error -Message "Retention tier failed; its artifacts were preserved." -Data $caseResult
            $runExitCode = 1

            try { Assert-NoExcelProcess -Code "EPT.RETENTION.SAFETY.EXCEL_REMAINS_AFTER_FAILURE" }
            catch {
                $safetyCode = "EPT.RETENTION.SAFETY.EXCEL_REMAINS_AFTER_FAILURE"
                if ($_.Exception.Data.Contains("stableCode")) { $safetyCode = [string]$_.Exception.Data["stableCode"] }
                $safetyError = Get-ErrorDetail -ErrorRecord $_ -StableCode $safetyCode
                $errors.Add($safetyError)
                Write-Event -Code $safetyCode -Level error -Message (
                    "Matrix aborted because an Excel process exists and cannot be treated as owned.") -Data $safetyError
                $safetyAbort = $true
            }
        }
        if ($safetyAbort) { break }
    }
}
catch {
    $stableCode = $script:stageCode + ".FAILED"
    if ($_.Exception.Data.Contains("stableCode")) { $stableCode = [string]$_.Exception.Data["stableCode"] }
    $detail = Get-ErrorDetail -ErrorRecord $_ -StableCode $stableCode
    $errors.Add($detail)
    Write-Event -Code $stableCode -Level error -Message "Retention matrix aborted." -Data $detail
    $runExitCode = 1
}
finally {
    $script:stageCode = "EPT.RETENTION.RUN.COMPLETE"
    $script:caseId = $null
    $totalStopwatch.Stop()
    $finalArtifactIntegrity = Get-PinnedRunArtifactIntegrity
    if (-not $finalArtifactIntegrity.allMatched) {
        try {
            Throw-RunnerError -Code "EPT.RETENTION.ARTIFACT.DRIFT_AT_FINALIZATION" -Message (
                "A pinned run artifact changed or became unavailable before finalization.") -Data @{
                    artifactIntegrity = ($finalArtifactIntegrity | ConvertTo-Json -Depth 12 -Compress)
                }
        }
        catch {
            $integrityError = Get-ErrorDetail -ErrorRecord $_ `
                -StableCode "EPT.RETENTION.ARTIFACT.DRIFT_AT_FINALIZATION"
            $errors.Add($integrityError)
            Write-Event -Code "EPT.RETENTION.ARTIFACT.DRIFT_AT_FINALIZATION" -Level error -Message (
                "Pinned run-artifact integrity failed at finalization.") -Data $integrityError
        }
        $runExitCode = 1
    }
    $result = [ordered]@{
        schema = $resultSchema
        runId = $runId
        success = ($runExitCode -eq 0)
        startedUtc = $startedUtc
        completedUtc = Get-UtcTimestamp
        elapsedMs = Get-ElapsedMilliseconds $totalStopwatch
        outputDirectory = $outputPath
        requestedRows = @($Rows)
        completedTierCount = @($cases | Where-Object { $_.status -eq "succeeded" }).Count
        failedTierCount = @($cases | Where-Object { $_.status -eq "failed" }).Count
        safetyAbort = $safetyAbort
        cases = $cases.ToArray()
        errors = $errors.ToArray()
        runOwnerPath = $runOwnerPath
        eventLogPath = $eventsPath
        artifactIntegrityAtFinalization = $finalArtifactIntegrity
        sourceLocation = [ordered]@{
            scriptPath = $controllerPath
            pinnedScriptSha256 = $script:expectedControllerSha256
            finalObservedScriptSha256 = $finalArtifactIntegrity.controller.actualSha256
        }
    }
    Write-Event -Code "EPT.RETENTION.RUN.COMPLETED" -Level $(if ($runExitCode -eq 0) { "info" } else { "error" }) `
        -Message "Excel retention matrix completed." -Data @{
            success = $result.success
            completedTierCount = $result.completedTierCount
            failedTierCount = $result.failedTierCount
            resultPath = $resultPath
        }
    if ($null -ne $script:eventWriter) {
        $script:eventWriter.Dispose()
        $script:eventWriter = $null
    }
    Write-NewJson -Path $resultPath -Value $result
    $compactResult = $result | ConvertTo-Json -Depth 40 -Compress
    if ($runExitCode -eq 0) { [Console]::Out.WriteLine($compactResult) }
    else { [Console]::Error.WriteLine($compactResult) }
}

exit $runExitCode
