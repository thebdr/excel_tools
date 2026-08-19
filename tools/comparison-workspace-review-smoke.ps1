param(
    [Parameter(Mandatory = $true)]
    [string]$Workspace,
    [string]$OutputDirectory
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$workspacePath = [IO.Path]::GetFullPath($Workspace)
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $env:TEMP ("ExcelPowerTools-workspace-review-smoke-" + [Guid]::NewGuid().ToString("N"))
}
$outputPath = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($outputPath) | Out-Null

$addinPath = Join-Path $workspacePath "src\ExcelPowerTools\obj\Debug\net48\ExcelPowerTools.dll"
$binPath = Join-Path $workspacePath "src\ExcelPowerTools\bin\Debug\net48"
$excelDnaPath = Join-Path $env:USERPROFILE ".nuget\packages\exceldna.addin\1.9.0\tools\net462\ExcelDna.Integration.dll"
[Reflection.Assembly]::LoadFrom($excelDnaPath) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path $binPath "DocumentFormat.OpenXml.Framework.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path $binPath "DocumentFormat.OpenXml.dll")) | Out-Null
$assembly = [Reflection.Assembly]::LoadFrom($addinPath)
$flags = [Reflection.BindingFlags]"Instance,Static,Public,NonPublic"

$reviewPath = Join-Path $outputPath "Validated Review.xlsx"
$legacyReviewPath = Join-Path $outputPath "Legacy Review.xlsx"
$projectId = "a" * 64
$stateType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Persistence.ReviewWorkbookState", $true)
$reviewFileType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Persistence.ReviewWorkbookFile", $true)
$create = @($reviewFileType.GetMethods($flags) | Where-Object {
    $_.IsStatic -and $_.Name -eq "CreateNew" -and $_.GetParameters().Count -eq 3
})[0]

# Build a genuine immutable decision sidecar. The Hub must never trust summary flags that are not backed by
# a validated catalog, so this fixture intentionally uses the same evidence -> catalog -> generation path as
# a real comparison project.
$revisionIdentityType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Decisions.ReviewEvidenceRevisionIdentity",
    $true)
$evidenceType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Decisions.ReviewEvidenceGeneration",
    $true)
$subjectType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Decisions.ReviewDecisionSubject",
    $true)
$groupType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Decisions.ReviewDecisionGroup",
    $true)
$catalogType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Decisions.ReviewDecisionCatalog",
    $true)
$decisionEngineType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Decisions.ReviewDecisionEngine",
    $true)
$observationType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Model.Observation",
    $true)
$scalarType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Model.CanonicalScalar",
    $true)
$provenanceType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Model.ObservationProvenance",
    $true)
$aspectType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Model.ComparisonAspect", $true)
$subjectPolicyType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Decisions.ReviewDecisionSubjectPolicy",
    $true)
$exactTypeType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Model.ObservationExactType", $true)
$dateSystemType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Canonical.ExcelDateSystem", $true)
$cachedTrustType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Model.CachedResultTrust", $true)
$comparisonIdsType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Identifiers.ComparisonIds",
    $true)

$revisionListType = [System.Collections.Generic.List``1].MakeGenericType($revisionIdentityType)
$revisions = [Activator]::CreateInstance($revisionListType)
$revisions.Add([Activator]::CreateInstance(
    $revisionIdentityType,
    [object[]]@("candidate", [int]0, ("1" * 64))))
$revisions.Add([Activator]::CreateInstance(
    $revisionIdentityType,
    [object[]]@("master", [int]1, ("2" * 64))))
$evidence = [Activator]::CreateInstance(
    $evidenceType,
    [object[]]@($projectId, "mapping-workspace-smoke", $revisions))

$valueAspect = [Enum]::Parse($aspectType, "Value")
$selectableAndMergeable = [Enum]::Parse($subjectPolicyType, "SelectableAndMergeable")
$subjectId = "subject-value"
$subjectListType = [System.Collections.Generic.List``1].MakeGenericType($subjectType)
$subjects = [Activator]::CreateInstance($subjectListType)
$subjects.Add([Activator]::CreateInstance(
    $subjectType,
    [object[]]@($subjectId, "logical-sheet:value:A1", $valueAspect, $selectableAndMergeable)))

$stringExactType = [Enum]::Parse($exactTypeType, "String")
$notApplicable = [Enum]::Parse($dateSystemType, "NotApplicable")
$notPresent = [Enum]::Parse($cachedTrustType, "NotPresent")
$observationIdMethod = @($comparisonIdsType.GetMethods($flags) | Where-Object {
    $_.IsStatic -and $_.Name -eq "Observation" -and $_.GetParameters().Count -eq 4
})[0]
$observationListType = [System.Collections.Generic.List``1].MakeGenericType($observationType)
$observations = [Activator]::CreateInstance($observationListType)

foreach ($observationInput in @(
    [pscustomobject]@{
        RevisionId = "candidate"; Hash = ("3" * 64); Value = "candidate value";
        Source = "candidate.xlsx"
    },
    [pscustomobject]@{
        RevisionId = "master"; Hash = ("4" * 64); Value = "master value";
        Source = "master.xlsx"
    })) {
    $observationId = $observationIdMethod.Invoke(
        $null,
        [object[]]@($subjectId, $observationInput.RevisionId, $stringExactType, $observationInput.Hash))
    $scalar = [Activator]::CreateInstance(
        $scalarType,
        [object[]]@($stringExactType, $observationInput.Value, $notApplicable))
    $provenance = [Activator]::CreateInstance(
        $provenanceType,
        [object[]]@($observationInput.Source, "Input", "A1", [int]1, [int]1, "logical-sheet"))
    $observation = [Activator]::CreateInstance(
        $observationType,
        [object[]]@(
            $observationId,
            $subjectId,
            $observationInput.RevisionId,
            $stringExactType,
            $observationInput.Hash,
            $null,
            $scalar,
            $null,
            $notPresent,
            $observationInput.Value,
            $provenance))
    $observations.Add($observation)
}

$groupListType = [System.Collections.Generic.List``1].MakeGenericType($groupType)
$groups = [Activator]::CreateInstance($groupListType)
$catalog = [Activator]::CreateInstance(
    $catalogType,
    [object[]]@($evidence, "master", $subjects, $observations, $groups))
$createRoot = @($decisionEngineType.GetMethods($flags) | Where-Object {
    $_.IsStatic -and $_.Name -eq "Create" -and $_.GetParameters().Count -eq 1
})[0]
$rootGeneration = $createRoot.Invoke($null, [object[]]@($catalog))

$state = [Activator]::CreateInstance($stateType, [object[]]@(
    $projectId,
    '{"schema":"workspace-smoke","sourcePath":"C:\\must-not-appear\\input.xlsx"}'))
$state.DecisionGeneration = $rootGeneration
$state.DecisionEvidenceGenerationId = $rootGeneration.BaseGenerationId
$state.DecisionMappingGenerationId = $rootGeneration.MappingGenerationId
$create.Invoke($null, [object[]]@([string]$reviewPath, $state, $null))

$legacyState = [Activator]::CreateInstance($stateType, [object[]]@(
    $projectId,
    '{"schema":"workspace-smoke","kind":"legacy"}'))
$create.Invoke($null, [object[]]@([string]$legacyReviewPath, $legacyState, $null))

$hydrationService = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Persistence.ReviewWorkbookHydrationService",
    $true)
$hydrate = $hydrationService.GetMethod("Hydrate", $flags)
$legacySession = $hydrate.Invoke($null, [object[]]@([string]$legacyReviewPath))
$session = $hydrate.Invoke($null, [object[]]@([string]$reviewPath))
$summary = $session.Summary
$sessionType = $session.GetType()
$sessionConstructor = @($sessionType.GetConstructors($flags) | Where-Object {
    $_.GetParameters().Count -eq 4
})[0]

# A retained session must reject a changed path occurrence before any UI command can use it.
$stalePath = Join-Path $outputPath "Stale occurrence Review.xlsx"
$replacementPath = Join-Path $outputPath "Replacement Review.xlsx"
$staleState = [Activator]::CreateInstance($stateType, [object[]]@(
    ("b" * 64),
    '{"schema":"workspace-smoke","occurrence":"before"}'))
$replacementState = [Activator]::CreateInstance($stateType, [object[]]@(
    ("c" * 64),
    '{"schema":"workspace-smoke","occurrence":"after"}'))
$create.Invoke($null, [object[]]@([string]$stalePath, $staleState, $null))
$create.Invoke($null, [object[]]@([string]$replacementPath, $replacementState, $null))
$staleSession = $hydrate.Invoke($null, [object[]]@([string]$stalePath))
[IO.File]::Copy($replacementPath, $stalePath, $true)
$exactOccurrenceRejected = $false
try { $staleSession.RevalidateCurrentOccurrence() | Out-Null }
catch { $exactOccurrenceRejected = $true }
if (-not $exactOccurrenceRejected) {
    throw "A changed Review.xlsx occurrence was accepted by the retained hydration session."
}

# Bind the future protected launch copy to the hydrated package without starting or automating Excel.
$launchType = $assembly.GetType(
    "ExcelPowerTools.Core.Comparison.Persistence.ReviewWorkbookLaunchCopy",
    $true)
$createLaunch = @($launchType.GetMethods($flags) | Where-Object {
    $_.IsStatic -and $_.Name -eq "Create" -and $_.GetParameters().Count -eq 2
})[0]
$commandsType = $assembly.GetType("ExcelPowerTools.Commands.ComparisonCommands", $true)
$ensureLaunch = $commandsType.GetMethod("EnsureLaunchMatchesHydratedSession", $flags)
$launchRoot = Join-Path $outputPath "protected-launch-smoke"
$launch = $createLaunch.Invoke($null, [object[]]@([string]$reviewPath, [string]$launchRoot))
$launchMismatchRejected = $false
try {
    $ensureLaunch.Invoke($null, [object[]]@($launch, $session))
    $wrongPackageSession = $sessionConstructor.Invoke([object[]]@(
        ("d" * 64),
        $session.FileOccurrenceToken,
        $session.ReadResult,
        $session.Summary))
    try { $ensureLaunch.Invoke($null, [object[]]@($launch, $wrongPackageSession)) }
    catch { $launchMismatchRejected = $true }
}
finally { $launch.Dispose() }
if (-not $launchMismatchRejected) {
    throw "A protected launch copy was not bound to the retained hydrated package identity."
}

$pageType = $assembly.GetType("ExcelPowerTools.UI.ComparisonWorkspacePage", $true)
$pageConstructor = @($pageType.GetConstructors($flags) | Where-Object {
    $_.GetParameters().Count -eq 0
})[0]
$showHydrated = $pageType.GetMethod("ShowHydratedReview", $flags)
$clearLoaded = $pageType.GetMethod("ClearLoadedReview", $flags)
$screenshots = @()
$eventSmokeComplete = $false
$decisionActionSmokeComplete = $false

foreach ($width in @(280, 420, 580)) {
    $page = $pageConstructor.Invoke(@())
    $form = New-Object Windows.Forms.Form
    $form.Text = "Comparison workspace review smoke $width"
    $form.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $form.Location = New-Object Drawing.Point(30, 30)
    $form.ClientSize = New-Object Drawing.Size($width, 900)
    $form.Controls.Add($page)
    try {
        $form.Show()
        [Windows.Forms.Application]::DoEvents()
        $showHydrated.Invoke($page, [object[]]@($legacySession))
        [Windows.Forms.Application]::DoEvents()
        $ledger = $pageType.GetField("_loadedDecisionLedger", $flags).GetValue($page)
        if ($ledger.Visible) {
            throw "A legacy review displayed a decision-ledger card at width $width."
        }

        $showHydrated.Invoke($page, [object[]]@($session))
        [Windows.Forms.Application]::DoEvents()

        $loadedSession = $pageType.GetProperty("LoadedSession", $flags).GetValue($page, $null)
        $loaded = $pageType.GetProperty("LoadedReview", $flags).GetValue($page, $null)
        if (-not [object]::ReferenceEquals($loadedSession, $session) -or
            $loaded.ProjectId -ne $projectId -or
            $loadedSession.ReadResult.StateJson -notmatch "must-not-appear") {
            throw "The complete hydrated session was not retained behind the privacy-safe summary."
        }
        $card = $pageType.GetField("_loadedReview", $flags).GetValue($page)
        if (-not $card.Visible -or $card.Width -le 0 -or $card.Height -le 0) {
            throw "Loaded-review card is not visible at width $width."
        }
        $fileLabel = $pageType.GetField("_loadedFileName", $flags).GetValue($page)
        $pathLabel = $pageType.GetField("_loadedPath", $flags).GetValue($page)
        $limitations = $pageType.GetField("_loadedLimitations", $flags).GetValue($page)
        if ($fileLabel.Text -ne "Validated Review.xlsx" -or
            $pathLabel.Text -ne [IO.Path]::GetFullPath($reviewPath)) {
            throw "Loaded-review identity labels are incorrect at width $width."
        }
        if ($limitations.Text -notmatch "source stays immutable" -or
            $limitations.Text -notmatch "new Review.xlsx" -or
            $limitations.Text -notmatch "revalidates") {
            throw "The loaded card does not explain copy-on-write review behavior."
        }
        if (-not $ledger.Visible -or $ledger.Left -lt 0 -or $ledger.Top -lt 0 -or
            $ledger.Right -gt $card.ClientSize.Width -or $ledger.Bottom -gt $card.ClientSize.Height) {
            throw "The decision-ledger card is hidden or outside the validated-review card at width $width."
        }
        $ledgerTitle = $pageType.GetField("_loadedDecisionLedgerTitle", $flags).GetValue($page)
        $ledgerMetadata = $pageType.GetField("_loadedDecisionLedgerMetadata", $flags).GetValue($page)
        if ($ledgerTitle.Text -notmatch "Decision ledger initialized" -or
            $ledgerMetadata.Text -notmatch "2 evidence revisions" -or
            $ledgerMetadata.Text -notmatch "1 subject" -or
            $ledgerMetadata.Text -notmatch "2 observations" -or
            $ledgerMetadata.Text -notmatch "0 active decisions" -or
            $ledgerMetadata.Text -notmatch [Regex]::Escape($summary.DecisionCatalogIdPrefix) -or
            $ledgerMetadata.Text -notmatch [Regex]::Escape($summary.DecisionGenerationIdPrefix)) {
            throw "The decision-ledger summary is incomplete at width $width."
        }
        $decisionReview = $pageType.GetField("_decisionReview", $flags).GetValue($page)
        $decisionReviewType = $decisionReview.GetType()
        $decisionList = $decisionReviewType.GetField("_ledger", $flags).GetValue($decisionReview)
        if (-not $decisionReview.Visible -or $decisionList.VirtualListSize -ne 1) {
            throw "The editable virtualized decision ledger did not load its exact subject row at width $width."
        }
        $allText = [string]::Join(" ", @(
            $card.Controls | ForEach-Object {
                $_.Text
                $_.Controls | ForEach-Object Text
            }))
        if ($allText.Contains("must-not-appear")) {
            throw "The loaded card exposed embedded untrusted JSON content."
        }

        $openButton = $pageType.GetField("_openLoadedReview", $flags).GetValue($page)
        $chooseButton = $pageType.GetField("_chooseAnotherReview", $flags).GetValue($page)
        foreach ($button in @($openButton, $chooseButton)) {
            if ($button.Left -lt 0 -or $button.Top -lt 0 -or
                $button.Right -gt $card.ClientSize.Width -or
                $button.Bottom -gt $card.ClientSize.Height) {
                throw "Button '$($button.Text)' is outside the loaded card at width $width " +
                      "(button=$($button.Bounds), card=$($card.ClientRectangle))."
            }
        }

        if (-not $eventSmokeComplete) {
            $script:commitRaised = $false
            $script:commitRequestCount = 0
            $script:commitExpectedGenerationId = ""
            $commitEvent = $decisionReviewType.GetEvent("CommitRequested", $flags)
            $commitHandlerScript = {
                param($sender, $eventArgs)
                $script:commitRaised = $true
                $eventFlags = [Reflection.BindingFlags]"Instance,Public,NonPublic"
                $requestsProperty = $eventArgs.GetType().GetProperty("Requests", $eventFlags)
                $generationProperty = $eventArgs.GetType().GetProperty(
                    "ExpectedFinalGenerationId",
                    $eventFlags)
                $script:commitRequestCount = $requestsProperty.GetValue($eventArgs, $null).Count
                $script:commitExpectedGenerationId = $generationProperty.GetValue($eventArgs, $null)
            }
            $commitHandler = [Management.Automation.LanguagePrimitives]::ConvertTo(
                $commitHandlerScript,
                $commitEvent.EventHandlerType)
            $commitEvent.GetAddMethod($true).Invoke(
                $decisionReview,
                [object[]]@($commitHandler)) | Out-Null
            $queueStatus = @($decisionReviewType.GetMethods($flags) | Where-Object {
                $_.Name -eq "QueueStatus" -and $_.GetParameters().Count -eq 2
            })[0]
            $statusType = $assembly.GetType(
                "ExcelPowerTools.Core.Comparison.Model.ReviewDecisionStatus",
                $true)
            $equivalent = [Enum]::Parse($statusType, "Equivalent")
            $investigate = [Enum]::Parse($statusType, "NeedsInvestigation")
            $visibleRowsField = $decisionReviewType.GetField("_visibleRows", $flags)
            $queueStatus.Invoke(
                $decisionReview,
                [object[]]@($visibleRowsField.GetValue($decisionReview), $equivalent))
            $queueStatus.Invoke(
                $decisionReview,
                [object[]]@($visibleRowsField.GetValue($decisionReview), $investigate))
            $pendingCount = $decisionReviewType.GetProperty(
                "PendingDecisionCount",
                $flags).GetValue($decisionReview, $null)
            if ($pendingCount -ne 2) {
                throw "The editable ledger did not retain two chained in-memory requests."
            }
            $saveButton = $decisionReviewType.GetField("_save", $flags).GetValue($decisionReview)
            $saveButton.PerformClick()
            if (-not $script:commitRaised -or $script:commitRequestCount -ne 2 -or
                [string]::IsNullOrWhiteSpace($script:commitExpectedGenerationId)) {
                throw "The editable ledger did not emit its complete pending decision batch."
            }
            $discardButton = $decisionReviewType.GetField("_discard", $flags).GetValue($decisionReview)
            $discardButton.PerformClick()
            $pendingCount = $decisionReviewType.GetProperty(
                "PendingDecisionCount",
                $flags).GetValue($decisionReview, $null)
            if ($pendingCount -ne 0) {
                throw "Discard pending did not restore the loaded decision generation."
            }
            $decisionActionSmokeComplete = $true

            $script:chooseRaised = $false
            $chooseEvent = $pageType.GetEvent("ChooseReviewRequested", $flags)
            $chooseHandler = [System.EventHandler]{ param($sender, $args) $script:chooseRaised = $true }
            $chooseEvent.GetAddMethod($true).Invoke($page, [object[]]@($chooseHandler)) | Out-Null
            $chooseButton.PerformClick()
            if (-not $script:chooseRaised) { throw "Choose-another navigation event did not fire." }

            $script:openRaised = $false
            $script:openedSession = $null
            $openEvent = $pageType.GetEvent("OpenValidatedReviewRequested", $flags)
            $handlerScript = {
                param($sender, $eventArgs)
                $script:openRaised = $true
                $eventFlags = [Reflection.BindingFlags]"Instance,Public,NonPublic"
                $sessionProperty = $eventArgs.GetType().GetProperty("Session", $eventFlags)
                $script:openedSession = $sessionProperty.GetValue($eventArgs, $null)
            }
            $openHandler = [Management.Automation.LanguagePrimitives]::ConvertTo(
                $handlerScript,
                $openEvent.EventHandlerType)
            $openEvent.GetAddMethod($true).Invoke($page, [object[]]@($openHandler)) | Out-Null
            $openButton.PerformClick()
            if (-not $script:openRaised -or
                -not [object]::ReferenceEquals($script:openedSession, $session)) {
                throw "Protected-open navigation did not retain the exact hydrated session."
            }

            $wrongClear = $clearLoaded.Invoke($page, [object[]]@($legacySession))
            if ($wrongClear -or $null -eq $pageType.GetProperty("LoadedSession", $flags).GetValue($page, $null)) {
                throw "A stale interaction cleared a newer hydrated selection."
            }
            $rightClear = $clearLoaded.Invoke($page, [object[]]@($session))
            if (-not $rightClear -or
                $null -ne $pageType.GetProperty("LoadedSession", $flags).GetValue($page, $null) -or
                $card.Visible) {
                throw "The hydrated session was not completely discarded after a failed selection boundary."
            }
            $showHydrated.Invoke($page, [object[]]@($session))
            [Windows.Forms.Application]::DoEvents()
            $eventSmokeComplete = $true
        }

        $content = $pageType.GetField("_content", $flags).GetValue($page)
        $content.ScrollControlIntoView($card)
        [Windows.Forms.Application]::DoEvents()
        $imagePath = Join-Path $outputPath ("comparison-workspace-loaded-$width.png")
        $bitmap = New-Object Drawing.Bitmap($form.ClientSize.Width, $form.ClientSize.Height)
        try {
            $form.DrawToBitmap(
                $bitmap,
                (New-Object Drawing.Rectangle(0, 0, $bitmap.Width, $bitmap.Height)))
            $bitmap.Save($imagePath, [Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $bitmap.Dispose() }
        $screenshots += $imagePath

        $content.ScrollControlIntoView($decisionReview)
        [Windows.Forms.Application]::DoEvents()
        $decisionImagePath = Join-Path $outputPath ("comparison-workspace-decisions-$width.png")
        $decisionBitmap = New-Object Drawing.Bitmap($form.ClientSize.Width, $form.ClientSize.Height)
        try {
            $form.DrawToBitmap(
                $decisionBitmap,
                (New-Object Drawing.Rectangle(0, 0, $decisionBitmap.Width, $decisionBitmap.Height)))
            $decisionBitmap.Save($decisionImagePath, [Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $decisionBitmap.Dispose() }
        $screenshots += $decisionImagePath
    }
    finally {
        $form.Close()
        $form.Dispose()
    }
}

[pscustomobject]@{
    Status = "SMOKE_OK"
    ProjectId = $summary.ProjectId
    HasDecisionSidecar = $summary.HasDecisionSidecar
    DecisionSubjects = $summary.DecisionSubjectCount
    DecisionObservations = $summary.DecisionObservationCount
    ChunkCount = $summary.StateChunkCount
    StateBytes = $summary.StateByteCount
    EventNavigation = $eventSmokeComplete
    DecisionActions = $decisionActionSmokeComplete
    ExactOccurrenceRejected = $exactOccurrenceRejected
    LaunchMismatchRejected = $launchMismatchRejected
    Screenshots = $screenshots
    OutputDirectory = $outputPath
} | ConvertTo-Json -Compress
