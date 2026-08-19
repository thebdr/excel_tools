param(
    [Parameter(Mandatory = $true)]
    [string]$Workspace
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$workspacePath = [IO.Path]::GetFullPath($Workspace)
$addinPath = Join-Path $workspacePath "src\ExcelPowerTools\obj\Debug\net48\ExcelPowerTools.dll"
$binPath = Join-Path $workspacePath "src\ExcelPowerTools\bin\Debug\net48"
$excelDnaPath = Join-Path $env:USERPROFILE ".nuget\packages\exceldna.addin\1.9.0\tools\net462\ExcelDna.Integration.dll"

[Reflection.Assembly]::LoadFrom($excelDnaPath) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path $binPath "DocumentFormat.OpenXml.Framework.dll")) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path $binPath "DocumentFormat.OpenXml.dll")) | Out-Null
$assembly = [Reflection.Assembly]::LoadFrom($addinPath)
$flags = [Reflection.BindingFlags]"Instance,Static,Public,NonPublic"

function Add-ZipTextEntry {
    param($Archive, [string]$Name, [string]$Text)
    $entry = $Archive.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
        try { $writer.Write($Text) } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function New-SmokeWorkbook {
    param([string]$Path, [string]$Note, [string]$ChangedValue)
    $file = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        $archive = New-Object IO.Compression.ZipArchive($file, [IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            Add-ZipTextEntry $archive "[Content_Types].xml" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
'@
            Add-ZipTextEntry $archive "_rels/.rels" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
            Add-ZipTextEntry $archive "xl/workbook.xml" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>
'@
            Add-ZipTextEntry $archive "xl/_rels/workbook.xml.rels" @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
'@
            $sheet = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>$Note</t></is></c><c r="B1" t="inlineStr"><is><t>metadata</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>ID</t></is></c><c r="B2" t="inlineStr"><is><t>Name</t></is></c><c r="C2" t="inlineStr"><is><t>Status</t></is></c><c r="D2" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>A</t></is></c><c r="B3" t="inlineStr"><is><t>Alpha</t></is></c><c r="C3" t="inlineStr"><is><t>Open</t></is></c><c r="D3"><v>10</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>B</t></is></c><c r="B4" t="inlineStr"><is><t>Beta</t></is></c><c r="C4" t="inlineStr"><is><t>Closed</t></is></c><c r="D4"><v>$ChangedValue</v></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>C</t></is></c><c r="B5" t="inlineStr"><is><t>Gamma</t></is></c><c r="C5" t="inlineStr"><is><t>Open</t></is></c><c r="D5"><v>30</v></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>D</t></is></c><c r="B6" t="inlineStr"><is><t>Delta</t></is></c><c r="C6" t="inlineStr"><is><t>Closed</t></is></c><c r="D6"><v>40</v></c></row>
  </sheetData>
</worksheet>
"@
            Add-ZipTextEntry $archive "xl/worksheets/sheet1.xml" $sheet
        }
        finally { $archive.Dispose() }
    }
    finally { $file.Dispose() }
}

function Get-HiddenProperty {
    param($Object, [string]$Name, [ref]$Result)
    $value = $Object.GetType().GetProperty($Name, $flags).GetValue($Object, $null)
    if ($value -is [Management.Automation.PSObject]) {
        $value = $value.BaseObject
    }
    $Result.Value = $value
}

function Unwrap-ReflectedResult {
    param($Value, [ref]$Result)
    if ($Value -is [Management.Automation.PSObject]) {
        $Value = $Value.BaseObject
    }
    $Result.Value = $Value
}

function Get-StaticMethod {
    param([Type]$Type, [string]$Name, [int]$ParameterCount)
    $matches = @($Type.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq $Name -and $_.GetParameters().Count -eq $ParameterCount
    })
    if ($matches.Count -ne 1) { throw "Expected one $($Type.FullName).$Name/$ParameterCount overload; found $($matches.Count)." }
    return $matches[0]
}

function Get-DescendantControls {
    param([Windows.Forms.Control]$Root)
    foreach ($control in $Root.Controls) {
        $control
        Get-DescendantControls $control
    }
}

$smokeRoot = Join-Path $env:TEMP ("ExcelPowerTools-correction-smoke-" + [Guid]::NewGuid().ToString("N"))
[IO.Directory]::CreateDirectory($smokeRoot) | Out-Null
try {
    $beforePath = Join-Path $smokeRoot "project_r1.xlsx"
    $afterPath = Join-Path $smokeRoot "project_r2.xlsx"
    $outputPath = Join-Path $smokeRoot "Review.xlsx"
    $screenshotPath = Join-Path $smokeRoot "alignment-review.png"
    $completeScreenshotPath = Join-Path $smokeRoot "alignment-review-complete.png"
    New-SmokeWorkbook $beforePath "draft note" "20"
    New-SmokeWorkbook $afterPath "reviewed note" "25"

    $prepType = $assembly.GetType("ExcelPowerTools.Commands.ComparisonProjectPreparationService", $true)
    $prepare = Get-StaticMethod $prepType "PrepareAsync" 3
    [string[]]$paths = @($beforePath, $afterPath)
    $prepareArguments = New-Object object[] 3
    $prepareArguments[0] = $paths
    $prepareArguments[1] = $null
    $prepareArguments[2] = [Threading.CancellationToken]::None
    $prepareTask = $prepare.Invoke($null, $prepareArguments)
    $preparation = $prepareTask.GetAwaiter().GetResult()
    $inference = $null
    Get-HiddenProperty $preparation "Inference" ([ref]$inference)
    $artifacts = $null
    Get-HiddenProperty $preparation "OrderedArtifacts" ([ref]$artifacts)

    $lineageEngine = $assembly.GetType("ExcelPowerTools.Core.Comparison.Lineage.LineageCorrectionEngine", $true)
    $lineageRequest = $assembly.GetType("ExcelPowerTools.Core.Comparison.Lineage.LineageCorrectionRequest", $true)
    $reviewerType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Model.ReviewerSnapshot", $true)
    $reviewer = $reviewerType.GetProperty("Empty", $flags).GetValue($null, $null)
    $mapping = $null
    Unwrap-ReflectedResult ((Get-StaticMethod $lineageEngine "FromInference" 1).Invoke($null, [object[]]@($inference))) ([ref]$mapping)
    $orderedIds = @($mapping.OrderedRevisionIds)
    $when = [DateTimeOffset]::UtcNow
    $setParent = $lineageRequest.GetMethod("SetParent", $flags)
    $setMaster = $lineageRequest.GetMethod("SetMaster", $flags)
    $applyLineage = Get-StaticMethod $lineageEngine "Apply" 2
    if ($mapping.GetParent($orderedIds[1]) -ne $orderedIds[0]) {
        $request = $null
        Unwrap-ReflectedResult ($setParent.Invoke($null, [object[]]@(
            $orderedIds[1], $orderedIds[0], $reviewer, $when, "smoke parent"))) ([ref]$request)
        Unwrap-ReflectedResult ($applyLineage.Invoke($null, [object[]]@($mapping, $request))) ([ref]$mapping)
    }
    if ($mapping.MasterRevisionId -ne $orderedIds[1]) {
        $request = $null
        Unwrap-ReflectedResult ($setMaster.Invoke($null, [object[]]@(
            $orderedIds[1], $reviewer, $when, "smoke master"))) ([ref]$request)
        Unwrap-ReflectedResult ($applyLineage.Invoke($null, [object[]]@($mapping, $request))) ([ref]$mapping)
    }

    $setupType = $assembly.GetType("ExcelPowerTools.UI.ComparisonProjectSetupResult", $true)
    $setupCtor = @($setupType.GetConstructors($flags))[0]
    $setupArguments = New-Object object[] 3
    $setupArguments[0] = $artifacts
    $setupArguments[1] = $mapping
    $setupArguments[2] = [string]$outputPath
    $setup = $null
    Unwrap-ReflectedResult ($setupCtor.Invoke($setupArguments)) ([ref]$setup)

    $commandsType = $assembly.GetType("ExcelPowerTools.Commands.ComparisonProjectCommands", $true)
    if (@($commandsType.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq "BuildAsync"
    }).Count -ne 0) {
        throw "Unsafe all-in-one BuildAsync publication wrapper is present. Incomplete analysis must require the explicit diagnostic-review path."
    }
    $loadInputs = Get-StaticMethod $commandsType "LoadInputsAsync" 4
    $loadTask = $loadInputs.Invoke($null, [object[]]@($preparation, $setup, $null, [Threading.CancellationToken]::None))
    $session = $loadTask.GetAwaiter().GetResult()

    $correctionEngine = $assembly.GetType("ExcelPowerTools.Core.Comparison.Alignment.AlignmentCorrectionEngine", $true)
    $correctionRequest = $assembly.GetType("ExcelPowerTools.Core.Comparison.Alignment.AlignmentCorrectionRequest", $true)
    $createCorrections = Get-StaticMethod $correctionEngine "Create" 1
    $applyCorrection = Get-StaticMethod $correctionEngine "Apply" 2
    $corrections = $createCorrections.Invoke($null, [object[]]@($mapping.GenerationId))
    $analyze = @($commandsType.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq "Analyze" -and $_.GetParameters().Count -eq 4
    })[0]
    $initial = $analyze.Invoke($null, [object[]]@($session, $corrections, $null, [Threading.CancellationToken]::None))
    $initialProject = $null
    Get-HiddenProperty $initial "ProjectResult" ([ref]$initialProject)
    if ($initialProject.IsComplete) { throw "The non-leading-header fixture unexpectedly completed without confirmation." }
    if (-not @($initialProject.Findings | Where-Object { $_.Code -eq "CMP-PROJECT-NONLEADING-HEADER" }).Count) {
        throw "The expected non-leading-header finding was not emitted."
    }

    $formType = $assembly.GetType("ExcelPowerTools.UI.ComparisonAlignmentReviewForm", $true)
    $formCtor = @($formType.GetConstructors($flags) | Where-Object { $_.GetParameters().Count -eq 1 })[0]
    $form = $formCtor.Invoke([object[]]@($initial))
    try {
        $form.Show()
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 250
        $bitmap = New-Object Drawing.Bitmap($form.ClientSize.Width, $form.ClientSize.Height)
        try {
            $form.DrawToBitmap($bitmap, (New-Object Drawing.Rectangle(0, 0, $bitmap.Width, $bitmap.Height)))
            $bitmap.Save($screenshotPath, [Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $bitmap.Dispose() }
        $footerBounds = @()
        foreach ($button in @(Get-DescendantControls $form | Where-Object {
            $_ -is [Windows.Forms.Button] -and
            $_.Text -in @("Rerun analysis", "Build diagnostic Review", "Continue to Review", "Cancel")
        })) {
            $origin = $form.PointToClient($button.Parent.PointToScreen($button.Location))
            $bounds = New-Object Drawing.Rectangle($origin, $button.Size)
            if ($bounds.Left -lt 0 -or $bounds.Top -lt 0 -or
                $bounds.Right -gt $form.ClientSize.Width -or $bounds.Bottom -gt $form.ClientSize.Height) {
                throw "Footer button '$($button.Text)' is outside the client area: $bounds in $($form.ClientRectangle)."
            }
            $footerBounds += "$($button.Text)=$($bounds.X),$($bounds.Y),$($bounds.Width),$($bounds.Height)"
        }
    }
    finally {
        $form.Close()
        $form.Dispose()
    }

    $edge = $initialProject.Edges[0]
    $alignment = $edge.SheetAlignments[0]
    $confirmHeader = @($correctionRequest.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq "ConfirmHeaderRow" -and $_.GetParameters().Count -eq 7
    })[0]
    $beforeRequest = $confirmHeader.Invoke($null, [object[]]@(
        $edge.ParentRevisionId, $alignment.BeforeSheetKey, 1, 2, $reviewer, [DateTimeOffset]::UtcNow, "smoke before header"))
    $afterRequest = $confirmHeader.Invoke($null, [object[]]@(
        $edge.ChildRevisionId, $alignment.AfterSheetKey, 1, 2, $reviewer, [DateTimeOffset]::UtcNow, "smoke after header"))
    $corrections = $applyCorrection.Invoke($null, [object[]]@($corrections, $beforeRequest))
    $corrections = $applyCorrection.Invoke($null, [object[]]@($corrections, $afterRequest))
    $confirmSheetPair = @($correctionRequest.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq "ConfirmSheetPair" -and $_.GetParameters().Count -eq 7
    })[0]
    $sheetPairRequest = $confirmSheetPair.Invoke($null, [object[]]@(
        $edge.ParentRevisionId, $edge.ChildRevisionId,
        $alignment.BeforeSheetKey, $alignment.AfterSheetKey,
        $reviewer, [DateTimeOffset]::UtcNow, "smoke sheet pair"))
    $corrections = $applyCorrection.Invoke($null, [object[]]@($corrections, $sheetPairRequest))

    $gridReferenceType = $assembly.GetType("ExcelPowerTools.Core.Comparison.Alignment.AlignmentGridIndexReference", $true)
    $gridReferenceCtor = @($gridReferenceType.GetConstructors($flags))[0]
    $beforeKeyColumns = [Array]::CreateInstance($gridReferenceType, 1)
    $afterKeyColumns = [Array]::CreateInstance($gridReferenceType, 1)
    $beforeKeyColumns.SetValue($gridReferenceCtor.Invoke([object[]]@(0, 1)), 0)
    $afterKeyColumns.SetValue($gridReferenceCtor.Invoke([object[]]@(0, 1)), 0)
    $confirmKey = @($correctionRequest.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq "ConfirmKeyColumns" -and $_.GetParameters().Count -eq 9
    })[0]
    $keyRequest = $confirmKey.Invoke($null, [object[]]@(
        $edge.ParentRevisionId, $edge.ChildRevisionId,
        $alignment.BeforeSheetKey, $alignment.AfterSheetKey,
        $beforeKeyColumns, $afterKeyColumns,
        $reviewer, [DateTimeOffset]::UtcNow, "smoke key"))
    $corrections = $applyCorrection.Invoke($null, [object[]]@($corrections, $keyRequest))
    $corrected = $analyze.Invoke($null, [object[]]@($session, $corrections, $null, [Threading.CancellationToken]::None))
    $correctedProject = $null
    Get-HiddenProperty $corrected "ProjectResult" ([ref]$correctedProject)
    if (-not $correctedProject.IsComplete) {
        throw "The confirmed equal non-leading headers did not complete: " +
            (($correctedProject.Findings | ForEach-Object Code) -join ", ")
    }
    if (@($correctedProject.Findings | Where-Object {
        $_.Code -eq "CMP-PROJECT-NONLEADING-HEADER"
    }).Count -ne 0) {
        throw "The corrected project still exposes the original header finding; the active-scope smoke is invalid."
    }

    $correctedForm = $formCtor.Invoke([object[]]@($corrected))
    try {
        $correctedForm.Show()
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 150
        $activeIssueCount = 0
        $activeHeaderIssueCount = 0
        $activeSheetIssueCount = 0
        $activeKeyIssueCount = 0
        $activeSheetIssue = $null
        $issues = $formType.GetField("_issues", $flags).GetValue($correctedForm)
        foreach ($issue in $issues) {
            $issueType = $issue.GetType()
            $synthetic = $issueType.GetProperty("IsSynthetic", $flags).GetValue($issue, $null)
            if (-not $synthetic) { continue }
            $activeIssueCount++
            $kind = $issueType.GetProperty("EditorKind", $flags).GetValue($issue, $null).ToString()
            if ($kind -eq "Header") { $activeHeaderIssueCount++ }
            if ($kind -eq "SheetPair") {
                $activeSheetIssueCount++
                $activeSheetIssue = $issue
            }
            if ($kind -eq "Key") { $activeKeyIssueCount++ }
        }
        if ($activeHeaderIssueCount -ne 1) {
            throw "Expected one synthetic paired-header issue after the original finding resolved; found $activeHeaderIssueCount."
        }
        if ($activeSheetIssueCount -ne 1 -or $activeKeyIssueCount -ne 1) {
            throw "Expected one synthetic sheet issue and one synthetic key issue; found sheet=$activeSheetIssueCount key=$activeKeyIssueCount."
        }
        $continueButtons = @(Get-DescendantControls $correctedForm | Where-Object {
            $_ -is [Windows.Forms.Button] -and $_.Text -eq "Continue to Review"
        })
        if ($continueButtons.Count -ne 1) {
            throw "A complete corrected analysis must expose exactly one Continue to Review action."
        }
        $issueList = $formType.GetField("_issueList", $flags).GetValue($correctedForm)
        $sheetItem = $null
        foreach ($item in $issueList.Items) { $item.Selected = $false }
        foreach ($item in $issueList.Items) {
            if ([object]::ReferenceEquals($item.Tag, $activeSheetIssue)) {
                $sheetItem = $item
                break
            }
        }
        if ($null -eq $sheetItem) { throw "The synthetic sheet-resolution issue is missing from the finding list." }
        $sheetItem.Selected = $true
        $sheetItem.Focused = $true
        [Windows.Forms.Application]::DoEvents()

        $deletedNewButtons = @(Get-DescendantControls $correctedForm | Where-Object {
            $_ -is [Windows.Forms.Button] -and $_.Text -eq "Treat as deleted + new"
        })
        if ($deletedNewButtons.Count -ne 1) {
            throw "The sheet-resolution editor must expose exactly one explicit deleted + new action."
        }
        $toolTips = $formType.GetField("_toolTips", $flags).GetValue($correctedForm)
        $interactive = @(Get-DescendantControls $correctedForm | Where-Object {
            $_ -is [Windows.Forms.Button] -or
            $_ -is [Windows.Forms.TextBox] -or
            $_ -is [Windows.Forms.ComboBox] -or
            $_ -is [Windows.Forms.RadioButton] -or
            $_ -is [Windows.Forms.NumericUpDown] -or
            $_ -is [Windows.Forms.CheckBox] -or
            $_ -is [Windows.Forms.ListView]
        })
        $missingToolTips = @($interactive | Where-Object {
            [string]::IsNullOrWhiteSpace($toolTips.GetToolTip($_))
        })
        if ($missingToolTips.Count -ne 0) {
            throw "Every interactive alignment-review control must have a tooltip; missing: " +
                (($missingToolTips | ForEach-Object { $_.GetType().Name + ":" + $_.Text }) -join ", ")
        }
        foreach ($menuItem in $issueList.ContextMenuStrip.Items) {
            if ($menuItem -is [Windows.Forms.ToolStripMenuItem] -and
                [string]::IsNullOrWhiteSpace($menuItem.ToolTipText)) {
                throw "Every alignment-review context action must have a tooltip: $($menuItem.Text)"
            }
        }
        $toolTipControlCount = $interactive.Count
        $completeBitmap = New-Object Drawing.Bitmap($correctedForm.ClientSize.Width, $correctedForm.ClientSize.Height)
        try {
            $correctedForm.DrawToBitmap(
                $completeBitmap,
                (New-Object Drawing.Rectangle(0, 0, $completeBitmap.Width, $completeBitmap.Height)))
            $completeBitmap.Save($completeScreenshotPath, [Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $completeBitmap.Dispose() }
    }
    finally {
        $correctedForm.Close()
        $correctedForm.Dispose()
    }

    $clearHeader = @($correctionRequest.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq "ClearHeaderRow" -and $_.GetParameters().Count -eq 5
    })[0]
    $cleared = $corrections
    $headerScopes = New-Object object[] 2
    $headerScopes[0] = [object[]]@($edge.ParentRevisionId, $alignment.BeforeSheetKey)
    $headerScopes[1] = [object[]]@($edge.ChildRevisionId, $alignment.AfterSheetKey)
    foreach ($headerScope in $headerScopes) {
        $clearRequest = $clearHeader.Invoke($null, [object[]]@(
            $headerScope[0], $headerScope[1], $reviewer, [DateTimeOffset]::UtcNow, "smoke return header to auto"))
        $cleared = $applyCorrection.Invoke($null, [object[]]@($cleared, $clearRequest))
    }
    $clearedAnalysis = $analyze.Invoke($null, [object[]]@(
        $session, $cleared, $null, [Threading.CancellationToken]::None))
    $clearedProject = $null
    Get-HiddenProperty $clearedAnalysis "ProjectResult" ([ref]$clearedProject)
    if ($clearedProject.IsComplete -or -not @($clearedProject.Findings | Where-Object {
        $_.Code -eq "CMP-PROJECT-NONLEADING-HEADER"
    }).Count) {
        throw "Clearing both confirmed headers did not safely return the project to the original incomplete auto-detection state."
    }

    $publish = Get-StaticMethod $commandsType "Publish" 3
    $publishResult = $publish.Invoke($null, [object[]]@($corrected, $null, [Threading.CancellationToken]::None))
    if (-not [IO.File]::Exists($outputPath)) { throw "Publish returned without creating Review.xlsx." }

    $reviewFile = $assembly.GetType("ExcelPowerTools.Core.Comparison.Persistence.ReviewWorkbookFile", $true)
    $readMethod = @($reviewFile.GetMethods($flags) | Where-Object {
        $_.IsStatic -and $_.Name -eq "ReadAndValidate" -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType -eq [string]
    })[0]
    $read = $readMethod.Invoke($null, [object[]]@([string]$outputPath))
    $state = $read.StateJson | ConvertFrom-Json
    if (-not $state.complete) { throw "Persisted comparison state is not complete." }
    if ($state.alignmentCorrections.events.Count -ne 4) { throw "Expected four persisted correction events." }
    if ($state.alignmentCorrections.activeHeaderConfirmations.Count -ne 2) {
        throw "Expected two active persisted header confirmations."
    }
    if ($state.alignmentCorrections.activeSheetPairConfirmations.Count -ne 1 -or
        $state.alignmentCorrections.activeKeyColumnConfirmations.Count -ne 1) {
        throw "Expected one active persisted sheet pair and key confirmation."
    }
    if ($state.alignmentCorrectionGenerationId -ne $corrections.GenerationId) {
        throw "Persisted correction generation ID does not match the analyzed generation."
    }
    if (-not $read.HasDecisionState) {
        throw "A complete comparison project did not publish its immutable decision root."
    }
    $decision = $read.DecisionGeneration
    if ($decision.Events.Count -ne 0 -or $decision.ActiveDecisions.Count -ne 0) {
        throw "A newly published comparison must start with an empty decision ledger."
    }
    if ($null -eq $decision.Catalog.EvidenceGeneration) {
        throw "The persisted decision catalog lost its exact evidence generation."
    }
    if ($decision.Catalog.Subjects.Count -ne $state.history.subjectCount) {
        throw "The persisted decision catalog subject count differs from logical history."
    }
    $expectedObservationCount = $decision.Catalog.Subjects.Count *
        $decision.Catalog.EvidenceGeneration.Revisions.Count
    if ($decision.Catalog.Observations.Count -ne $expectedObservationCount) {
        throw "The persisted decision catalog does not contain the full subject-by-revision matrix."
    }

    [pscustomobject]@{
        Status = "SMOKE_OK"
        Complete = $correctedProject.IsComplete
        FindingsBefore = $initialProject.Findings.Count
        CorrectionEvents = $state.alignmentCorrections.events.Count
        ActiveHeaders = $state.alignmentCorrections.activeHeaderConfirmations.Count
        ActiveSyntheticIssues = $activeIssueCount
        DeletedNewAction = $true
        ToolTipControls = $toolTipControlCount
        DecisionSubjects = $decision.Catalog.Subjects.Count
        DecisionObservations = $decision.Catalog.Observations.Count
        FooterBounds = ($footerBounds -join "; ")
        ReviewBytes = (Get-Item $outputPath).Length
        ReviewPath = $outputPath
        ScreenshotPath = $screenshotPath
        CompleteScreenshotPath = $completeScreenshotPath
    } | ConvertTo-Json -Compress
}
catch {
    throw
}
