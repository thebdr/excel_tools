using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Persistence;

namespace ExcelPowerTools.ReviewSpike;

internal static class Program
{
    private const string MetricsSchema = "excel-power-tools.review-spike.metrics.v1";
    private const string BaselineSchema = "excel-power-tools.review-spike.baseline.v1";
    private const string StateSchema = "excel-power-tools.review-spike.state.v1";
    private const string GeneratorName = "ExcelPowerTools.ReviewSpike";
    private const string GeneratorVersion = "2";
    private const int BaselineMaximumBytes = 128 * 1024;
    private const int MaximumArgumentCount = 32;
    private const int MaximumOptionValueCharacters = 32767;
    private const int MaximumProjectIdCharacters = 256;
    private const int MaximumPrefixCharacters = 64;
    private const int SnapshotGroupSize = 100;
    private const int SnapshotRows = 5;
    private const int SnapshotColumns = 4;
    private const int RetentionMismatchExitCode = 3;
    private const int MatrixIncompleteExitCode = 4;

    private static readonly int[] MatrixRowCounts = { 10_000, 100_000, 250_000 };
    private static readonly int[] MatrixStateBytes = { 2 * 1024 * 1024, 8 * 1024 * 1024, 24 * 1024 * 1024 };
    private static readonly string[] ChangesHeader =
    {
        "Change ID",
        "Sheet",
        "Address",
        "Aspect",
        "Before",
        "After",
        "Group ID",
        "Flags",
        "Revisions",
        "Distinct states",
        "Alignment confidence",
    };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private static readonly JsonSerializerOptions BaselineJsonOptions = new(JsonOptions)
    {
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
    };

    public static int Main(string[] args)
    {
        if (args.Length == 0 || IsHelp(args[0]))
        {
            WriteUsage();
            return args.Length == 0 ? 2 : 0;
        }

        string command = args[0].ToLowerInvariant();
        try
        {
            Dictionary<string, string> options = ParseOptions(args.Skip(1).ToArray());
            CommandExecution execution = command switch
            {
                "generate" => Generate(options),
                "matrix" => GenerateMatrix(options),
                "verify" => Verify(options),
                _ => throw new ArgumentException("Unknown command: " + args[0]),
            };

            Console.Out.WriteLine(JsonSerializer.Serialize(execution.Metrics, JsonOptions));
            return execution.ExitCode;
        }
        catch (Exception exception)
        {
            var failure = new
            {
                Schema = MetricsSchema,
                Command = command,
                Success = false,
                ErrorType = exception.GetType().FullName,
                Error = exception.Message,
            };
            Console.Error.WriteLine(JsonSerializer.Serialize(failure, JsonOptions));
            return 1;
        }
    }

    private static CommandExecution Generate(IReadOnlyDictionary<string, string> options)
    {
        RequireOnly(
            options,
            "output",
            "rows",
            "project-id",
            "chunk-size",
            "state-bytes",
            "baseline-output");
        string outputPath = GetWorkbookOutputPath(options, "output");
        int rowCount = GetInteger(options, "rows", 1000, 0, ReviewWorkbookFile.MaxProjectedChangeRows);
        int chunkSize = GetChunkSize(options);
        int requestedStateBytes = GetInteger(
            options,
            "state-bytes",
            0,
            0,
            ReviewWorkbookFile.MaxEmbeddedStateBytes);
        string projectId = ValidateProjectId(
            GetOptional(options, "project-id") ??
            "review-spike-" + rowCount.ToString(CultureInfo.InvariantCulture));
        string? baselinePath = GetOptionalOutputPath(options, "baseline-output", ".json");
        if (baselinePath != null) EnsureDistinctPaths(outputPath, baselinePath);

        GeneratedArtifact artifact = GenerateArtifact(
            outputPath,
            projectId,
            rowCount,
            chunkSize,
            requestedStateBytes,
            baselinePath);
        return new CommandExecution(CreateGenerateMetrics(artifact), 0);
    }

    private static CommandExecution GenerateMatrix(IReadOnlyDictionary<string, string> options)
    {
        RequireOnly(options, "output-directory", "prefix", "chunk-size", "state-bytes");
        string outputDirectory = Path.GetFullPath(GetRequired(options, "output-directory"));
        if (!Directory.Exists(outputDirectory))
            throw new DirectoryNotFoundException("The matrix output directory does not exist: " + outputDirectory);

        string prefix = GetOptional(options, "prefix") ?? "review-spike";
        ValidateFilePrefix(prefix);
        int chunkSize = GetChunkSize(options);
        int? requestedStateBytesOverride = options.ContainsKey("state-bytes")
            ? GetInteger(
                options,
                "state-bytes",
                0,
                1,
                ReviewWorkbookFile.MaxEmbeddedStateBytes)
            : null;

        var plans = new List<MatrixPlan>(MatrixRowCounts.Length);
        for (int index = 0; index < MatrixRowCounts.Length; index++)
        {
            int rowCount = MatrixRowCounts[index];
            string stem = prefix + "-" + rowCount.ToString(CultureInfo.InvariantCulture);
            string outputPath = Path.Combine(outputDirectory, stem + ".xlsx");
            string baselinePath = Path.Combine(outputDirectory, stem + ".baseline.json");
            plans.Add(new MatrixPlan(
                outputPath,
                baselinePath,
                "review-spike-" + rowCount.ToString(CultureInfo.InvariantCulture),
                rowCount,
                requestedStateBytesOverride ?? MatrixStateBytes[index]));
        }

        foreach (MatrixPlan plan in plans)
        {
            EnsureNewOutputPath(plan.OutputPath, ".xlsx");
            EnsureNewOutputPath(plan.BaselinePath, ".json");
        }

        var total = Stopwatch.StartNew();
        var artifactMetrics = new List<object>(plans.Count);
        bool matrixSucceeded = true;
        foreach (MatrixPlan plan in plans)
        {
            try
            {
                GeneratedArtifact artifact = GenerateArtifact(
                    plan.OutputPath,
                    plan.ProjectId,
                    plan.RowCount,
                    chunkSize,
                    plan.RequestedStateBytes,
                    plan.BaselinePath);
                artifactMetrics.Add(CreateGenerateMetrics(artifact));
            }
            catch (Exception exception)
            {
                matrixSucceeded = false;
                artifactMetrics.Add(new MatrixArtifactFailure
                {
                    Schema = MetricsSchema,
                    Command = "generate",
                    Success = false,
                    OutputPath = plan.OutputPath,
                    BaselineOutputPath = plan.BaselinePath,
                    ProjectId = plan.ProjectId,
                    RowCount = plan.RowCount,
                    RequestedStateUtf8Bytes = plan.RequestedStateBytes,
                    ChunkSizeBytes = chunkSize,
                    ErrorType = exception.GetType().FullName ?? exception.GetType().Name,
                    Error = exception.Message,
                    OutputPathExistsAfterFailure = File.Exists(plan.OutputPath) || Directory.Exists(plan.OutputPath),
                    BaselinePathExistsAfterFailure =
                        File.Exists(plan.BaselinePath) || Directory.Exists(plan.BaselinePath),
                });
            }
        }
        total.Stop();

        var metrics = new MatrixMetrics
        {
            Schema = MetricsSchema,
            Command = "matrix",
            Success = matrixSucceeded,
            OutputDirectory = outputDirectory,
            Prefix = prefix,
            RowCounts = MatrixRowCounts,
            ChunkSizeBytes = chunkSize,
            RequestedStateUtf8Bytes = plans.Select(plan => plan.RequestedStateBytes).ToArray(),
            TotalElapsedMs = Milliseconds(total),
            Timings = new TimingMetrics { TotalElapsedMs = Milliseconds(total) },
            Artifacts = artifactMetrics.ToArray(),
        };
        return new CommandExecution(metrics, matrixSucceeded ? 0 : MatrixIncompleteExitCode);
    }

    private static CommandExecution Verify(IReadOnlyDictionary<string, string> options)
    {
        long allocatedBytesAtStart = GC.GetTotalAllocatedBytes(precise: false);
        RequireOnly(options, "input", "baseline", "baseline-output");
        string inputPath = GetInputPath(options, "input", ".xlsx");
        string? baselineInputPath = GetOptionalInputPath(options, "baseline", ".json");
        string? baselineOutputPath = GetOptionalOutputPath(options, "baseline-output", ".json");
        if (baselineInputPath != null && baselineOutputPath != null)
            throw new ArgumentException("--baseline and --baseline-output cannot be used together.");
        if (baselineOutputPath != null) EnsureDistinctPaths(inputPath, baselineOutputPath);

        var total = Stopwatch.StartNew();
        var baselineRead = new Stopwatch();
        RetentionBaseline? baseline = null;
        if (baselineInputPath != null)
        {
            baselineRead.Start();
            baseline = ReadBaseline(baselineInputPath);
            baselineRead.Stop();
        }

        Measurement measurement = MeasureExisting(inputPath);
        BaselineComparison? comparison = baseline == null ? null : CompareBaseline(baseline, measurement);

        var baselineWrite = new Stopwatch();
        if (baselineOutputPath != null)
        {
            baselineWrite.Start();
            WriteBaselineCreateNew(baselineOutputPath, CreateBaseline(measurement));
            baselineWrite.Stop();
        }
        total.Stop();

        bool success = comparison?.SemanticRetentionMatched ?? true;
        var metrics = CreateVerifyMetrics(
            measurement,
            baselineInputPath,
            baselineOutputPath,
            comparison,
            baselineRead,
            baselineWrite,
            total,
            success,
            Math.Max(0, GC.GetTotalAllocatedBytes(precise: false) - allocatedBytesAtStart),
            Process.GetCurrentProcess().PeakWorkingSet64);
        return new CommandExecution(metrics, success ? 0 : RetentionMismatchExitCode);
    }

    private static GeneratedArtifact GenerateArtifact(
        string outputPath,
        string projectId,
        int rowCount,
        int chunkSize,
        int requestedStateBytes,
        string? baselinePath)
    {
        long allocatedBytesAtStart = GC.GetTotalAllocatedBytes(precise: false);
        EnsureNewOutputPath(outputPath, ".xlsx");
        if (baselinePath != null) EnsureNewOutputPath(baselinePath, ".json");

        var total = Stopwatch.StartNew();
        var prepare = Stopwatch.StartNew();
        string stateJson = CreateStateJson(projectId, rowCount, requestedStateBytes, out int snapshotCount);
        int actualStateBytes = Encoding.UTF8.GetByteCount(stateJson);
        int expectedChunkCount = Math.Max(1, checked((actualStateBytes + chunkSize - 1) / chunkSize));
        var state = new ReviewWorkbookState(projectId, stateJson)
        {
            DetailBefore = "Review spike before value",
            DetailAfter = "Review spike after value",
        };

        for (int index = 0; index < rowCount; index++)
        {
            int excelRow = index + 2;
            int groupIndex = index / SnapshotGroupSize;
            string ordinal = (index + 1).ToString("D8", CultureInfo.InvariantCulture);
            string groupOrdinal = (groupIndex + 1).ToString("D6", CultureInfo.InvariantCulture);
            state.ChangeRows.Add(new ReviewChangeProjectionRow(
                "change-" + ordinal,
                "Data",
                "A" + excelRow.ToString(CultureInfo.InvariantCulture),
                index % 4 == 0 ? "Formula" : "Value",
                "before-" + ordinal,
                "after-" + ordinal,
                "group-" + groupOrdinal,
                index % 997 == 0 ? "review" : "",
                5,
                2 + index % 4,
                index % 20 == 0 ? "weak" : "strong"));
        }
        prepare.Stop();

        var write = Stopwatch.StartNew();
        ReviewWorkbookFile.CreateNew(
            outputPath,
            state,
            new ReviewWorkbookWriteOptions { StateChunkSizeBytes = chunkSize });
        write.Stop();
        bool baselinePublished = false;
        try
        {
            Measurement measurement = MeasureExisting(outputPath);
            if (!string.Equals(measurement.Read.ProjectId, projectId, StringComparison.Ordinal) ||
                !string.Equals(measurement.Read.StateJson, stateJson, StringComparison.Ordinal))
                throw new InvalidDataException("The reopened workbook state differs from the generated state.");
            if (measurement.RowCount != rowCount)
                throw new InvalidDataException(
                    "The reopened workbook contains " +
                    measurement.RowCount.ToString(CultureInfo.InvariantCulture) +
                    " projected rows; expected " + rowCount.ToString(CultureInfo.InvariantCulture) + ".");
            if (measurement.SnapshotCount != snapshotCount)
                throw new InvalidDataException("The reopened workbook snapshot count differs from the generated state.");
            if (measurement.Read.ChunkCount != expectedChunkCount)
                throw new InvalidDataException(
                    "The reopened workbook contains " +
                    measurement.Read.ChunkCount.ToString(CultureInfo.InvariantCulture) +
                    " state chunks; expected " + expectedChunkCount.ToString(CultureInfo.InvariantCulture) + ".");

            var baselineWrite = new Stopwatch();
            if (baselinePath != null)
            {
                baselineWrite.Start();
                WriteBaselineCreateNew(baselinePath, CreateBaseline(measurement));
                baselineWrite.Stop();
                baselinePublished = true;
            }
            total.Stop();

            return new GeneratedArtifact(
                measurement,
                baselinePath,
                requestedStateBytes,
                expectedChunkCount,
                chunkSize,
                prepare,
                write,
                baselineWrite,
                total,
                Math.Max(0, GC.GetTotalAllocatedBytes(precise: false) - allocatedBytesAtStart),
                Process.GetCurrentProcess().PeakWorkingSet64);
        }
        catch
        {
            if (baselinePublished && baselinePath != null) TryDelete(baselinePath);
            TryDelete(outputPath);
            throw;
        }
    }

    private static string CreateStateJson(
        string projectId,
        int rowCount,
        int requestedStateBytes,
        out int snapshotCount)
    {
        List<SpikeSnapshot> snapshots = CreateSnapshots(rowCount);
        snapshotCount = snapshots.Count;
        var emptyState = new SpikeState(
            StateSchema,
            GeneratorName,
            GeneratorVersion,
            projectId,
            rowCount,
            SnapshotGroupSize,
            SnapshotRows,
            SnapshotColumns,
            snapshotCount,
            snapshots,
            "");
        string stateJson = JsonSerializer.Serialize(emptyState, JsonOptions);
        int naturalBytes = Encoding.UTF8.GetByteCount(stateJson);
        if (requestedStateBytes <= naturalBytes) return stateJson;

        int paddingLength = requestedStateBytes - naturalBytes;
        string padding = CreateDeterministicPadding(paddingLength, projectId, rowCount);
        stateJson = JsonSerializer.Serialize(emptyState with { Padding = padding }, JsonOptions);
        int actualBytes = Encoding.UTF8.GetByteCount(stateJson);
        if (actualBytes != requestedStateBytes)
            throw new InvalidOperationException("The deterministic state payload did not reach its requested size.");
        if (actualBytes > ReviewWorkbookFile.MaxEmbeddedStateBytes)
            throw new InvalidDataException("The generated state exceeds the safe review-workbook limit.");
        return stateJson;
    }

    private static List<SpikeSnapshot> CreateSnapshots(int rowCount)
    {
        int count = checked((rowCount + SnapshotGroupSize - 1) / SnapshotGroupSize);
        var snapshots = new List<SpikeSnapshot>(count);
        for (int groupIndex = 0; groupIndex < count; groupIndex++)
        {
            int groupOrdinal = groupIndex + 1;
            int firstChangeOrdinal = checked(groupIndex * SnapshotGroupSize + 1);
            var before = new string[SnapshotRows * SnapshotColumns];
            var after = new string[before.Length];
            for (int row = 0; row < SnapshotRows; row++)
            {
                for (int column = 0; column < SnapshotColumns; column++)
                {
                    int position = row * SnapshotColumns + column;
                    int relativeRow = row - SnapshotRows / 2;
                    string coordinate =
                        ":g" + groupOrdinal.ToString("D6", CultureInfo.InvariantCulture) +
                        ":r" + relativeRow.ToString("+00;-00;00", CultureInfo.InvariantCulture) +
                        ":c" + (column + 1).ToString("D2", CultureInfo.InvariantCulture) +
                        ":v" + Math.Max(1, firstChangeOrdinal + relativeRow).ToString("D8", CultureInfo.InvariantCulture);
                    before[position] = "before" + coordinate;
                    after[position] = "after" + coordinate;
                }
            }
            snapshots.Add(new SpikeSnapshot(
                "group-" + groupOrdinal.ToString("D6", CultureInfo.InvariantCulture),
                firstChangeOrdinal,
                before,
                after));
        }
        return snapshots;
    }

    private static string CreateDeterministicPadding(int characterCount, string projectId, int rowCount)
    {
        if (characterCount == 0) return "";
        var result = new StringBuilder(characterCount);
        byte[] seed = Encoding.UTF8.GetBytes(
            GeneratorName + "\n" + GeneratorVersion + "\n" + projectId + "\n" +
            rowCount.ToString(CultureInfo.InvariantCulture));
        Span<byte> counterBytes = stackalloc byte[sizeof(int)];
        int counter = 0;
        while (result.Length < characterCount)
        {
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(counterBytes, counter++);
            byte[] material = new byte[seed.Length + counterBytes.Length];
            seed.CopyTo(material, 0);
            counterBytes.CopyTo(material.AsSpan(seed.Length));
            string block = Convert.ToHexString(SHA256.HashData(material)).ToLowerInvariant();
            int remaining = characterCount - result.Length;
            result.Append(block, 0, Math.Min(block.Length, remaining));
        }
        return result.ToString();
    }

    private static void ValidateDeterministicPadding(
        string padding,
        string projectId,
        int rowCount)
    {
        byte[] seed = Encoding.UTF8.GetBytes(
            GeneratorName + "\n" + GeneratorVersion + "\n" + projectId + "\n" +
            rowCount.ToString(CultureInfo.InvariantCulture));
        Span<byte> counterBytes = stackalloc byte[sizeof(int)];
        int counter = 0;
        int offset = 0;
        while (offset < padding.Length)
        {
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(counterBytes, counter++);
            byte[] material = new byte[seed.Length + counterBytes.Length];
            seed.CopyTo(material, 0);
            counterBytes.CopyTo(material.AsSpan(seed.Length));
            string block = Convert.ToHexString(SHA256.HashData(material)).ToLowerInvariant();
            int count = Math.Min(block.Length, padding.Length - offset);
            if (!padding.AsSpan(offset, count).SequenceEqual(block.AsSpan(0, count)))
                throw new InvalidDataException("The spike-state deterministic padding is invalid.");
            offset += count;
        }
    }

    private static Measurement MeasureExisting(string workbookPath)
    {
        var total = Stopwatch.StartNew();
        using var protectionStream = new FileStream(
            workbookPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            1024 * 1024,
            FileOptions.RandomAccess);
        long originalLength = protectionStream.Length;

        var validate = Stopwatch.StartNew();
        ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(protectionStream);
        validate.Stop();

        var inspect = Stopwatch.StartNew();
        SpikeStateInfo stateInfo = ReadSpikeStateInfo(read.StateJson, read.ProjectId);
        ChangeInspection changes = InspectProjectedChanges(
            protectionStream,
            stateInfo.RowCount ?? throw new InvalidDataException("The spike state has no row count."));
        PackageSizeInfo package = InspectPackageSize(protectionStream);
        inspect.Stop();

        if (stateInfo.RowCount.HasValue && stateInfo.RowCount.Value != changes.RowCount)
            throw new InvalidDataException(
                "The spike state declares " + stateInfo.RowCount.Value.ToString(CultureInfo.InvariantCulture) +
                " projected rows, but the workbook contains " +
                changes.RowCount.ToString(CultureInfo.InvariantCulture) + ".");

        var hash = Stopwatch.StartNew();
        string workbookSha256 = ComputeFileSha256(protectionStream);
        hash.Stop();
        if (protectionStream.Length != originalLength)
            throw new IOException("The review workbook changed while it was being measured.");
        total.Stop();

        return new Measurement(
            workbookPath,
            read,
            changes.RowCount,
            stateInfo.RowCount,
            stateInfo.SnapshotCount,
            stateInfo.GeneratorVersion,
            Encoding.UTF8.GetByteCount(read.StateJson),
            originalLength,
            workbookSha256,
            package.EntryCount,
            package.DeclaredUncompressedBytes,
            package.LargestEntryName,
            package.LargestEntryUncompressedBytes,
            changes.WorksheetUncompressedBytes,
            changes.SemanticSha256,
            validate,
            inspect,
            hash,
            total);
    }

    private static ChangeInspection InspectProjectedChanges(Stream workbookStream, int expectedRowCount)
    {
        long originalPosition = workbookStream.Position;
        try
        {
            workbookStream.Position = 0;
            var settings = new OpenSettings
            {
                AutoSave = false,
                MaxCharactersInPart = ReviewWorkbookFile.MaxStreamingWorksheetCharacters,
            };
            using SpreadsheetDocument document = SpreadsheetDocument.Open(workbookStream, false, settings);
            WorkbookPart workbookPart = document.WorkbookPart ??
                throw new InvalidDataException("The review workbook has no workbook part.");
            Sheet changes = workbookPart.Workbook?.Sheets?.Elements<Sheet>()
                .SingleOrDefault(sheet => string.Equals(sheet.Name?.Value, "Changes", StringComparison.Ordinal)) ??
                throw new InvalidDataException("The review workbook has no unique Changes sheet.");
            string relationshipId = changes.Id?.Value ??
                throw new InvalidDataException("The Changes sheet has no relationship ID.");
            WorksheetPart worksheetPart = workbookPart.GetPartById(relationshipId) as WorksheetPart ??
                throw new InvalidDataException("The Changes sheet does not target a worksheet part.");

            IReadOnlyList<string> sharedStrings = workbookPart.SharedStringTablePart?
                .SharedStringTable?
                .Elements<SharedStringItem>()
                .Select(item => item.InnerText)
                .ToArray() ?? Array.Empty<string>();
            long worksheetBytes;
            using (Stream worksheetStream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
                worksheetBytes = worksheetStream.Length;

            int worksheetRows = 0;
            using IncrementalHash semanticHash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            AppendSemanticValue(semanticHash, "excel-power-tools.review-spike.visible-changes.v1");
            using OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            while (reader.Read())
            {
                if (!reader.IsStartElement || reader.ElementType != typeof(Row)) continue;
                if (reader.LoadCurrentElement() is not Row row)
                    throw new InvalidDataException("The Changes sheet row could not be read.");
                int expectedExcelRow = worksheetRows + 1;
                if (row.RowIndex?.Value != (uint)expectedExcelRow)
                    throw new InvalidDataException(
                        "The Changes sheet row index is invalid at logical row " +
                        expectedExcelRow.ToString(CultureInfo.InvariantCulture) + ".");
                Cell[] cells = row.Elements<Cell>().ToArray();
                if (cells.Length != ChangesHeader.Length)
                    throw new InvalidDataException(
                        "The Changes sheet row " + expectedExcelRow.ToString(CultureInfo.InvariantCulture) +
                        " does not contain exactly A:K.");
                for (int column = 0; column < cells.Length; column++)
                {
                    string expectedReference = ColumnName(column + 1) +
                        expectedExcelRow.ToString(CultureInfo.InvariantCulture);
                    if (!string.Equals(cells[column].CellReference?.Value, expectedReference, StringComparison.Ordinal))
                        throw new InvalidDataException(
                            "The Changes sheet cell reference differs from " + expectedReference + ".");
                    if (cells[column].CellFormula != null)
                        throw new InvalidDataException("The Changes sheet contains an unexpected formula.");
                    string actualValue = ReadCellText(cells[column], sharedStrings);
                    string expectedValue = worksheetRows == 0
                        ? ChangesHeader[column]
                        : ExpectedChangeValue(worksheetRows - 1, column);
                    if (!string.Equals(actualValue, expectedValue, StringComparison.Ordinal))
                        throw new InvalidDataException(
                            "The Changes sheet value differs at " + expectedReference + ".");
                    AppendSemanticValue(semanticHash, expectedReference);
                    AppendSemanticValue(semanticHash, actualValue);
                }
                worksheetRows++;
            }

            if (worksheetRows == 0)
                throw new InvalidDataException("The Changes sheet has no header row.");
            int projectedRows = worksheetRows - 1;
            if (projectedRows != expectedRowCount)
                throw new InvalidDataException(
                    "The Changes sheet contains " + projectedRows.ToString(CultureInfo.InvariantCulture) +
                    " projected rows; expected " + expectedRowCount.ToString(CultureInfo.InvariantCulture) + ".");
            return new ChangeInspection(
                projectedRows,
                worksheetBytes,
                Convert.ToHexString(semanticHash.GetHashAndReset()).ToLowerInvariant());
        }
        finally
        {
            if (workbookStream.CanSeek) workbookStream.Position = originalPosition;
        }
    }

    private static string ReadCellText(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (!int.TryParse(cell.CellValue?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out int index) ||
                index < 0 || index >= sharedStrings.Count)
                throw new InvalidDataException("The Changes sheet contains an invalid shared-string reference.");
            return sharedStrings[index];
        }
        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? "";
        if (cell.DataType?.Value == CellValues.String)
            return cell.CellValue?.Text ?? "";
        throw new InvalidDataException("The Changes sheet contains a non-text cell.");
    }

    private static string ExpectedChangeValue(int zeroBasedIndex, int column)
    {
        int excelRow = zeroBasedIndex + 2;
        string ordinal = (zeroBasedIndex + 1).ToString("D8", CultureInfo.InvariantCulture);
        string groupOrdinal = (zeroBasedIndex / SnapshotGroupSize + 1)
            .ToString("D6", CultureInfo.InvariantCulture);
        return column switch
        {
            0 => "change-" + ordinal,
            1 => "Data",
            2 => "A" + excelRow.ToString(CultureInfo.InvariantCulture),
            3 => zeroBasedIndex % 4 == 0 ? "Formula" : "Value",
            4 => "before-" + ordinal,
            5 => "after-" + ordinal,
            6 => "group-" + groupOrdinal,
            7 => zeroBasedIndex % 997 == 0 ? "review" : "",
            8 => "5",
            9 => (2 + zeroBasedIndex % 4).ToString(CultureInfo.InvariantCulture),
            10 => zeroBasedIndex % 20 == 0 ? "weak" : "strong",
            _ => throw new ArgumentOutOfRangeException(nameof(column)),
        };
    }

    private static void AppendSemanticValue(IncrementalHash hash, string value)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        Span<byte> length = stackalloc byte[sizeof(int)];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(length, bytes.Length);
        hash.AppendData(length);
        hash.AppendData(bytes);
    }

    private static string ColumnName(int oneBasedColumn)
    {
        var result = new StringBuilder();
        int value = oneBasedColumn;
        while (value > 0)
        {
            value--;
            result.Insert(0, (char)('A' + value % 26));
            value /= 26;
        }
        return result.ToString();
    }

    private static PackageSizeInfo InspectPackageSize(Stream workbookStream)
    {
        long originalPosition = workbookStream.Position;
        try
        {
            workbookStream.Position = 0;
            using var archive = new ZipArchive(workbookStream, ZipArchiveMode.Read, leaveOpen: true);
            long uncompressedBytes = 0;
            long largestEntryBytes = -1;
            string largestEntryName = "";
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                uncompressedBytes = checked(uncompressedBytes + entry.Length);
                if (entry.Length > largestEntryBytes)
                {
                    largestEntryBytes = entry.Length;
                    largestEntryName = entry.FullName;
                }
            }
            return new PackageSizeInfo(
                archive.Entries.Count,
                uncompressedBytes,
                largestEntryName,
                Math.Max(0, largestEntryBytes));
        }
        finally
        {
            if (workbookStream.CanSeek) workbookStream.Position = originalPosition;
        }
    }

    private static SpikeStateInfo ReadSpikeStateInfo(string stateJson, string manifestProjectId)
    {
        using JsonDocument document = JsonDocument.Parse(
            stateJson,
            new JsonDocumentOptions { MaxDepth = 32, CommentHandling = JsonCommentHandling.Disallow });
        JsonElement root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object ||
            !root.TryGetProperty("schema", out JsonElement schema) ||
            schema.ValueKind != JsonValueKind.String ||
            !string.Equals(schema.GetString(), StateSchema, StringComparison.Ordinal))
            throw new InvalidDataException("The embedded state is not a ReviewSpike state.");

        if (!root.TryGetProperty("generator", out JsonElement generator) ||
            generator.ValueKind != JsonValueKind.String ||
            !string.Equals(generator.GetString(), GeneratorName, StringComparison.Ordinal))
            throw new InvalidDataException("The spike state generator is invalid.");
        if (!root.TryGetProperty("generatorVersion", out JsonElement version) ||
            version.ValueKind != JsonValueKind.String ||
            !string.Equals(version.GetString(), GeneratorVersion, StringComparison.Ordinal))
            throw new InvalidDataException("The spike state generator version is unsupported.");
        if (!root.TryGetProperty("projectId", out JsonElement projectId) ||
            projectId.ValueKind != JsonValueKind.String ||
            !string.Equals(projectId.GetString(), manifestProjectId, StringComparison.Ordinal))
            throw new InvalidDataException("The spike-state project ID differs from the manifest.");

        if (!root.TryGetProperty("rowCount", out JsonElement rowCount) ||
            !rowCount.TryGetInt32(out int rows) ||
            rows < 0 || rows > ReviewWorkbookFile.MaxProjectedChangeRows)
            throw new InvalidDataException("The spike state rowCount is missing or invalid.");

        if (!root.TryGetProperty("snapshotGroupSize", out JsonElement snapshotGroupSize) ||
            !snapshotGroupSize.TryGetInt32(out int groupSize) || groupSize != SnapshotGroupSize ||
            !root.TryGetProperty("snapshotRows", out JsonElement snapshotRows) ||
            !snapshotRows.TryGetInt32(out int snapshotRowCount) || snapshotRowCount != SnapshotRows ||
            !root.TryGetProperty("snapshotColumns", out JsonElement snapshotColumns) ||
            !snapshotColumns.TryGetInt32(out int snapshotColumnCount) || snapshotColumnCount != SnapshotColumns)
            throw new InvalidDataException("The spike-state snapshot dimensions are invalid.");

        int? snapshotCount = null;
        if (root.TryGetProperty("snapshotCount", out JsonElement declaredSnapshots))
        {
            if (!declaredSnapshots.TryGetInt32(out int count) || count < 0 || count > 100_000)
                throw new InvalidDataException("The spike state snapshotCount is invalid.");
            if (!root.TryGetProperty("snapshots", out JsonElement snapshots) ||
                snapshots.ValueKind != JsonValueKind.Array ||
                snapshots.GetArrayLength() != count)
                throw new InvalidDataException("The spike state snapshot array differs from snapshotCount.");
            int expectedCount = checked((rows + SnapshotGroupSize - 1) / SnapshotGroupSize);
            if (count != expectedCount)
                throw new InvalidDataException("The spike-state snapshot count is not deterministic for its row count.");
            ValidateSnapshots(snapshots, count);
            snapshotCount = count;
        }
        else
        {
            throw new InvalidDataException("The spike state snapshotCount is missing.");
        }

        if (!root.TryGetProperty("padding", out JsonElement paddingElement) ||
            paddingElement.ValueKind != JsonValueKind.String)
            throw new InvalidDataException("The spike-state padding is missing or invalid.");
        ValidateDeterministicPadding(paddingElement.GetString() ?? "", manifestProjectId, rows);
        return new SpikeStateInfo(rows, snapshotCount, GeneratorVersion);
    }

    private static void ValidateSnapshots(JsonElement snapshots, int count)
    {
        int index = 0;
        foreach (JsonElement snapshot in snapshots.EnumerateArray())
        {
            int groupOrdinal = index + 1;
            int firstChangeOrdinal = checked(index * SnapshotGroupSize + 1);
            string expectedGroup = "group-" + groupOrdinal.ToString("D6", CultureInfo.InvariantCulture);
            if (snapshot.ValueKind != JsonValueKind.Object ||
                !snapshot.TryGetProperty("groupId", out JsonElement groupId) ||
                groupId.ValueKind != JsonValueKind.String ||
                !string.Equals(groupId.GetString(), expectedGroup, StringComparison.Ordinal) ||
                !snapshot.TryGetProperty("firstChangeOrdinal", out JsonElement first) ||
                !first.TryGetInt32(out int actualFirst) || actualFirst != firstChangeOrdinal ||
                !snapshot.TryGetProperty("before", out JsonElement before) ||
                !snapshot.TryGetProperty("after", out JsonElement after))
                throw new InvalidDataException("A spike-state snapshot identity is invalid.");
            ValidateSnapshotValues(before, groupOrdinal, firstChangeOrdinal, "before");
            ValidateSnapshotValues(after, groupOrdinal, firstChangeOrdinal, "after");
            index++;
        }
        if (index != count) throw new InvalidDataException("The spike-state snapshot count changed during validation.");
    }

    private static void ValidateSnapshotValues(
        JsonElement values,
        int groupOrdinal,
        int firstChangeOrdinal,
        string prefix)
    {
        if (values.ValueKind != JsonValueKind.Array ||
            values.GetArrayLength() != SnapshotRows * SnapshotColumns)
            throw new InvalidDataException("A spike-state snapshot grid is invalid.");
        int position = 0;
        foreach (JsonElement value in values.EnumerateArray())
        {
            int row = position / SnapshotColumns;
            int column = position % SnapshotColumns;
            int relativeRow = row - SnapshotRows / 2;
            string expected = prefix +
                ":g" + groupOrdinal.ToString("D6", CultureInfo.InvariantCulture) +
                ":r" + relativeRow.ToString("+00;-00;00", CultureInfo.InvariantCulture) +
                ":c" + (column + 1).ToString("D2", CultureInfo.InvariantCulture) +
                ":v" + Math.Max(1, firstChangeOrdinal + relativeRow)
                    .ToString("D8", CultureInfo.InvariantCulture);
            if (value.ValueKind != JsonValueKind.String ||
                !string.Equals(value.GetString(), expected, StringComparison.Ordinal))
                throw new InvalidDataException("A spike-state snapshot value is invalid.");
            position++;
        }
    }

    private static string ComputeFileSha256(Stream input)
    {
        long originalPosition = input.Position;
        try
        {
            input.Position = 0;
            return Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant();
        }
        finally
        {
            if (input.CanSeek) input.Position = originalPosition;
        }
    }

    private static RetentionBaseline CreateBaseline(Measurement measurement) => new()
    {
        Schema = BaselineSchema,
        Generator = GeneratorName,
        GeneratorVersion = measurement.GeneratorVersion ?? "",
        FileName = Path.GetFileName(measurement.Path),
        ProjectId = measurement.Read.ProjectId,
        RowCount = measurement.RowCount,
        DeclaredRowCount = measurement.DeclaredRowCount,
        SnapshotCount = measurement.SnapshotCount,
        StateUtf8Bytes = measurement.StateUtf8Bytes,
        WorkbookBytes = measurement.WorkbookBytes,
        WorkbookSha256 = measurement.WorkbookSha256,
        StateSha256 = measurement.Read.StateSha256,
        ProjectionSha256 = measurement.Read.ProjectionSha256,
        VisibleProjectionSemanticSha256 = measurement.VisibleProjectionSemanticSha256,
        ManifestItemId = measurement.Read.ManifestItemId,
        ChunkCount = measurement.Read.ChunkCount,
        ChunkItemIds = measurement.Read.ChunkItemIds.ToArray(),
        HasDecisionState = measurement.Read.HasDecisionState,
        DecisionItemId = measurement.Read.DecisionItemId,
        DecisionSha256 = measurement.Read.DecisionSha256,
    };

    private static void WriteBaselineCreateNew(string destinationPath, RetentionBaseline baseline)
    {
        ValidateBaseline(baseline);
        EnsureNewOutputPath(destinationPath, ".json");
        byte[] payload = JsonSerializer.SerializeToUtf8Bytes(baseline, BaselineJsonOptions);
        if (payload.Length > BaselineMaximumBytes)
            throw new InvalidDataException("The retention baseline exceeds its safe size limit.");

        string directory = Path.GetDirectoryName(destinationPath)!;
        string temporary = Path.Combine(
            directory,
            "." + Path.GetFileName(destinationPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            using (var output = new FileStream(
                temporary,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                16 * 1024,
                FileOptions.WriteThrough))
            {
                output.Write(payload);
                output.Flush(true);
            }
            File.Move(temporary, destinationPath);
        }
        catch
        {
            TryDelete(temporary);
            throw;
        }
    }

    private static RetentionBaseline ReadBaseline(string path)
    {
        using var input = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            16 * 1024,
            FileOptions.SequentialScan);
        if (input.Length > BaselineMaximumBytes)
            throw new InvalidDataException("The retention baseline exceeds its safe size limit.");
        byte[] payload = new byte[checked((int)input.Length)];
        input.ReadExactly(payload);
        RetentionBaseline baseline = JsonSerializer.Deserialize<RetentionBaseline>(payload, BaselineJsonOptions) ??
            throw new InvalidDataException("The retention baseline is empty.");
        ValidateBaseline(baseline);
        return baseline;
    }

    private static void ValidateBaseline(RetentionBaseline baseline)
    {
        if (!string.Equals(baseline.Schema, BaselineSchema, StringComparison.Ordinal))
            throw new InvalidDataException("The retention baseline schema is unsupported.");
        if (!string.Equals(baseline.Generator, GeneratorName, StringComparison.Ordinal))
            throw new InvalidDataException("The retention baseline generator is unsupported.");
        if (baseline.GeneratorVersion.Length > 32 ||
            baseline.FileName.Length > 512 ||
            string.IsNullOrWhiteSpace(baseline.ProjectId) ||
            baseline.ProjectId.Length > MaximumProjectIdCharacters)
            throw new InvalidDataException("The retention baseline identity fields are invalid.");
        if (baseline.RowCount < 0 || baseline.RowCount > ReviewWorkbookFile.MaxProjectedChangeRows ||
            (baseline.DeclaredRowCount.HasValue &&
             (baseline.DeclaredRowCount.Value < 0 ||
              baseline.DeclaredRowCount.Value > ReviewWorkbookFile.MaxProjectedChangeRows)) ||
            (baseline.SnapshotCount.HasValue &&
             (baseline.SnapshotCount.Value < 0 || baseline.SnapshotCount.Value > 100_000)) ||
            baseline.StateUtf8Bytes < 0 || baseline.StateUtf8Bytes > ReviewWorkbookFile.MaxEmbeddedStateBytes ||
            baseline.WorkbookBytes < 0 || baseline.WorkbookBytes > ReviewWorkbookFile.MaxReviewPackageBytes ||
            baseline.ChunkCount < 1 || baseline.ChunkCount > ReviewWorkbookFile.MaxStateChunkCount)
            throw new InvalidDataException("The retention baseline numeric fields are invalid.");
        if (baseline.ChunkItemIds == null || baseline.ChunkItemIds.Length != baseline.ChunkCount)
            throw new InvalidDataException("The retention baseline chunk identifiers are invalid.");
        ValidateHash(baseline.WorkbookSha256, "workbookSha256");
        ValidateHash(baseline.StateSha256, "stateSha256");
        ValidateHash(baseline.ProjectionSha256, "projectionSha256");
        ValidateHash(baseline.VisibleProjectionSemanticSha256, "visibleProjectionSemanticSha256");
        if (baseline.HasDecisionState)
        {
            ValidateHash(baseline.DecisionSha256, "decisionSha256");
            if (string.IsNullOrWhiteSpace(baseline.DecisionItemId))
                throw new InvalidDataException("The retention baseline decision item ID is missing.");
        }
        else if (!string.IsNullOrEmpty(baseline.DecisionItemId) || !string.IsNullOrEmpty(baseline.DecisionSha256))
        {
            throw new InvalidDataException("The retention baseline has unexpected decision fields.");
        }
        if (string.IsNullOrWhiteSpace(baseline.ManifestItemId) ||
            baseline.ChunkItemIds.Any(string.IsNullOrWhiteSpace) ||
            baseline.ChunkItemIds.Distinct(StringComparer.Ordinal).Count() != baseline.ChunkItemIds.Length)
            throw new InvalidDataException("The retention baseline Custom XML item identifiers are invalid.");
    }

    private static void ValidateHash(string value, string name)
    {
        if (value.Length != 64 || value.Any(character => !Uri.IsHexDigit(character)))
            throw new InvalidDataException("The retention baseline " + name + " is not a SHA-256 digest.");
    }

    private static BaselineComparison CompareBaseline(RetentionBaseline baseline, Measurement measurement)
    {
        bool projectIdMatches = string.Equals(baseline.ProjectId, measurement.Read.ProjectId, StringComparison.Ordinal);
        bool rowCountMatches = baseline.RowCount == measurement.RowCount;
        bool declaredRowCountMatches = baseline.DeclaredRowCount == measurement.DeclaredRowCount;
        bool snapshotCountMatches = baseline.SnapshotCount == measurement.SnapshotCount;
        bool stateSizeMatches = baseline.StateUtf8Bytes == measurement.StateUtf8Bytes;
        bool stateShaMatches = FixedHashEquals(baseline.StateSha256, measurement.Read.StateSha256);
        bool projectionShaMatches = FixedHashEquals(baseline.ProjectionSha256, measurement.Read.ProjectionSha256);
        bool visibleSemanticShaMatches = FixedHashEquals(
            baseline.VisibleProjectionSemanticSha256,
            measurement.VisibleProjectionSemanticSha256);
        bool manifestItemIdMatches = string.Equals(
            baseline.ManifestItemId,
            measurement.Read.ManifestItemId,
            StringComparison.Ordinal);
        bool chunkCountMatches = baseline.ChunkCount == measurement.Read.ChunkCount;
        bool chunkItemIdsMatch = baseline.ChunkItemIds.SequenceEqual(
            measurement.Read.ChunkItemIds,
            StringComparer.Ordinal);
        bool decisionStateMatches = baseline.HasDecisionState == measurement.Read.HasDecisionState &&
            string.Equals(baseline.DecisionItemId, measurement.Read.DecisionItemId, StringComparison.Ordinal) &&
            FixedOptionalHashEquals(baseline.DecisionSha256, measurement.Read.DecisionSha256);
        bool generatorVersionMatches = string.Equals(
            baseline.GeneratorVersion,
            measurement.GeneratorVersion ?? "",
            StringComparison.Ordinal);
        bool packageHashMatches = FixedHashEquals(baseline.WorkbookSha256, measurement.WorkbookSha256);

        bool semanticRetentionMatched =
            projectIdMatches &&
            rowCountMatches &&
            declaredRowCountMatches &&
            snapshotCountMatches &&
            stateSizeMatches &&
            stateShaMatches &&
            projectionShaMatches &&
            visibleSemanticShaMatches &&
            manifestItemIdMatches &&
            chunkCountMatches &&
            chunkItemIdsMatch &&
            decisionStateMatches &&
            generatorVersionMatches;
        return new BaselineComparison
        {
            SemanticRetentionMatched = semanticRetentionMatched,
            ProjectIdMatches = projectIdMatches,
            RowCountMatches = rowCountMatches,
            DeclaredRowCountMatches = declaredRowCountMatches,
            SnapshotCountMatches = snapshotCountMatches,
            StateUtf8BytesMatches = stateSizeMatches,
            StateSha256Matches = stateShaMatches,
            ProjectionSha256Matches = projectionShaMatches,
            VisibleProjectionSemanticSha256Matches = visibleSemanticShaMatches,
            ManifestItemIdMatches = manifestItemIdMatches,
            ChunkCountMatches = chunkCountMatches,
            ChunkItemIdsMatch = chunkItemIdsMatch,
            DecisionStateMatches = decisionStateMatches,
            GeneratorVersionMatches = generatorVersionMatches,
            PackageSha256Matches = packageHashMatches,
            WorkbookBytesDelta = measurement.WorkbookBytes - baseline.WorkbookBytes,
            BaselineWorkbookBytes = baseline.WorkbookBytes,
            BaselineWorkbookSha256 = baseline.WorkbookSha256,
        };
    }

    private static bool FixedHashEquals(string left, string right)
    {
        if (left.Length != right.Length) return false;
        byte[] leftBytes = Encoding.ASCII.GetBytes(left.ToLowerInvariant());
        byte[] rightBytes = Encoding.ASCII.GetBytes(right.ToLowerInvariant());
        return CryptographicOperations.FixedTimeEquals(leftBytes, rightBytes);
    }

    private static bool FixedOptionalHashEquals(string left, string right) =>
        left.Length == 0 && right.Length == 0 || FixedHashEquals(left, right);

    private static CommandMetrics CreateGenerateMetrics(GeneratedArtifact artifact)
    {
        Measurement measurement = artifact.Measurement;
        bool stateSizeSatisfied = measurement.StateUtf8Bytes >= artifact.RequestedStateBytes;
        return new CommandMetrics
        {
            Schema = MetricsSchema,
            Command = "generate",
            Success = true,
            OutputPath = measurement.Path,
            BaselineOutputPath = artifact.BaselinePath,
            ProjectId = measurement.Read.ProjectId,
            RowCount = measurement.RowCount,
            DeclaredRowCount = measurement.DeclaredRowCount,
            SnapshotCount = measurement.SnapshotCount,
            RequestedStateUtf8Bytes = artifact.RequestedStateBytes,
            StateUtf8Bytes = measurement.StateUtf8Bytes,
            WorkbookBytes = measurement.WorkbookBytes,
            WorkbookSha256 = measurement.WorkbookSha256,
            StateSha256 = measurement.Read.StateSha256,
            ProjectionSha256 = measurement.Read.ProjectionSha256,
            VisibleProjectionSemanticSha256 = measurement.VisibleProjectionSemanticSha256,
            ManifestItemId = measurement.Read.ManifestItemId,
            ChunkCount = measurement.Read.ChunkCount,
            ExpectedChunkCount = artifact.ExpectedChunkCount,
            ChunkSizeBytes = artifact.ChunkSize,
            PackageEntryCount = measurement.PackageEntryCount,
            PackageDeclaredUncompressedBytes = measurement.PackageDeclaredUncompressedBytes,
            LargestPackageEntryName = measurement.LargestPackageEntryName,
            LargestPackageEntryUncompressedBytes = measurement.LargestPackageEntryUncompressedBytes,
            ChangesWorksheetUncompressedBytes = measurement.ChangesWorksheetUncompressedBytes,
            PrepareElapsedMs = Milliseconds(artifact.Prepare),
            WriteElapsedMs = Milliseconds(artifact.Write),
            ValidateElapsedMs = Milliseconds(measurement.Validate),
            InspectElapsedMs = Milliseconds(measurement.Inspect),
            HashElapsedMs = Milliseconds(measurement.Hash),
            BaselineWriteElapsedMs = Milliseconds(artifact.BaselineWrite),
            TotalElapsedMs = Milliseconds(artifact.Total),
            Timings = new TimingMetrics
            {
                PrepareElapsedMs = Milliseconds(artifact.Prepare),
                WriteElapsedMs = Milliseconds(artifact.Write),
                ValidateElapsedMs = Milliseconds(measurement.Validate),
                InspectElapsedMs = Milliseconds(measurement.Inspect),
                HashElapsedMs = Milliseconds(measurement.Hash),
                BaselineWriteElapsedMs = Milliseconds(artifact.BaselineWrite),
                TotalElapsedMs = Milliseconds(artifact.Total),
            },
            Sizes = CreateSizeMetrics(measurement, artifact.RequestedStateBytes),
            Hashes = CreateHashMetrics(measurement),
            Resources = CreateResourceMetrics(artifact.AllocatedBytes, artifact.ProcessPeakWorkingSetBytes),
            Runtime = CreateRuntimeMetrics(),
            Disk = CreateDiskMetrics(measurement.Path),
            Validation = new ValidationMetrics
            {
                StrictReaderPassed = true,
                StateRoundTripMatched = true,
                ProjectedRowCountMatched = true,
                DeclaredRowCountMatched = measurement.DeclaredRowCount == measurement.RowCount,
                SnapshotCountMatched = true,
                RequestedStateSizeSatisfied = stateSizeSatisfied,
                ChunkCountMatched = measurement.Read.ChunkCount == artifact.ExpectedChunkCount,
                BaselineCompared = false,
            },
        };
    }

    private static CommandMetrics CreateVerifyMetrics(
        Measurement measurement,
        string? baselineInputPath,
        string? baselineOutputPath,
        BaselineComparison? comparison,
        Stopwatch baselineRead,
        Stopwatch baselineWrite,
        Stopwatch total,
        bool success,
        long allocatedBytes,
        long processPeakWorkingSetBytes)
    {
        return new CommandMetrics
        {
            Schema = MetricsSchema,
            Command = "verify",
            Success = success,
            InputPath = measurement.Path,
            BaselineInputPath = baselineInputPath,
            BaselineOutputPath = baselineOutputPath,
            ProjectId = measurement.Read.ProjectId,
            RowCount = measurement.RowCount,
            DeclaredRowCount = measurement.DeclaredRowCount,
            SnapshotCount = measurement.SnapshotCount,
            StateUtf8Bytes = measurement.StateUtf8Bytes,
            WorkbookBytes = measurement.WorkbookBytes,
            WorkbookSha256 = measurement.WorkbookSha256,
            StateSha256 = measurement.Read.StateSha256,
            ProjectionSha256 = measurement.Read.ProjectionSha256,
            VisibleProjectionSemanticSha256 = measurement.VisibleProjectionSemanticSha256,
            ManifestItemId = measurement.Read.ManifestItemId,
            ChunkCount = measurement.Read.ChunkCount,
            PackageEntryCount = measurement.PackageEntryCount,
            PackageDeclaredUncompressedBytes = measurement.PackageDeclaredUncompressedBytes,
            LargestPackageEntryName = measurement.LargestPackageEntryName,
            LargestPackageEntryUncompressedBytes = measurement.LargestPackageEntryUncompressedBytes,
            ChangesWorksheetUncompressedBytes = measurement.ChangesWorksheetUncompressedBytes,
            ValidateElapsedMs = Milliseconds(measurement.Validate),
            InspectElapsedMs = Milliseconds(measurement.Inspect),
            HashElapsedMs = Milliseconds(measurement.Hash),
            BaselineReadElapsedMs = Milliseconds(baselineRead),
            BaselineWriteElapsedMs = Milliseconds(baselineWrite),
            TotalElapsedMs = Milliseconds(total),
            Timings = new TimingMetrics
            {
                ValidateElapsedMs = Milliseconds(measurement.Validate),
                InspectElapsedMs = Milliseconds(measurement.Inspect),
                HashElapsedMs = Milliseconds(measurement.Hash),
                BaselineReadElapsedMs = Milliseconds(baselineRead),
                BaselineWriteElapsedMs = Milliseconds(baselineWrite),
                TotalElapsedMs = Milliseconds(total),
            },
            Sizes = CreateSizeMetrics(measurement, null),
            Hashes = CreateHashMetrics(measurement),
            Resources = CreateResourceMetrics(allocatedBytes, processPeakWorkingSetBytes),
            Runtime = CreateRuntimeMetrics(),
            Disk = CreateDiskMetrics(measurement.Path),
            Validation = new ValidationMetrics
            {
                StrictReaderPassed = true,
                ProjectedRowCountMatched = true,
                DeclaredRowCountMatched = !measurement.DeclaredRowCount.HasValue ||
                    measurement.DeclaredRowCount.Value == measurement.RowCount,
                SnapshotCountMatched = true,
                BaselineCompared = comparison != null,
                SemanticRetentionMatched = comparison?.SemanticRetentionMatched,
            },
            Retention = comparison,
        };
    }

    private static SizeMetrics CreateSizeMetrics(Measurement measurement, int? requestedStateBytes) => new()
    {
        RequestedStateUtf8Bytes = requestedStateBytes,
        StateUtf8Bytes = measurement.StateUtf8Bytes,
        WorkbookBytes = measurement.WorkbookBytes,
        PackageDeclaredUncompressedBytes = measurement.PackageDeclaredUncompressedBytes,
        LargestPackageEntryName = measurement.LargestPackageEntryName,
        LargestPackageEntryUncompressedBytes = measurement.LargestPackageEntryUncompressedBytes,
        ChangesWorksheetUncompressedBytes = measurement.ChangesWorksheetUncompressedBytes,
        PackageEntryCount = measurement.PackageEntryCount,
        ProjectedRowCount = measurement.RowCount,
        SnapshotCount = measurement.SnapshotCount,
        ChunkCount = measurement.Read.ChunkCount,
    };

    private static HashMetrics CreateHashMetrics(Measurement measurement) => new()
    {
        WorkbookSha256 = measurement.WorkbookSha256,
        StateSha256 = measurement.Read.StateSha256,
        ProjectionSha256 = measurement.Read.ProjectionSha256,
        VisibleProjectionSemanticSha256 = measurement.VisibleProjectionSemanticSha256,
        DecisionSha256 = measurement.Read.HasDecisionState ? measurement.Read.DecisionSha256 : null,
        RawWorkbookSha256Deterministic = false,
        SemanticHashesComparableAcrossExcelSave = true,
    };

    private static ResourceMetrics CreateResourceMetrics(long allocatedBytes, long processPeakWorkingSetBytes)
    {
        using Process process = Process.GetCurrentProcess();
        return new ResourceMetrics
        {
            AllocatedBytesForCommand = allocatedBytes,
            ProcessPeakWorkingSetBytes = processPeakWorkingSetBytes,
            ProcessWorkingSetBytesAtEnd = process.WorkingSet64,
            ManagedHeapBytesAtEnd = GC.GetTotalMemory(forceFullCollection: false),
        };
    }

    private static RuntimeMetrics CreateRuntimeMetrics() => new()
    {
        GeneratorName = GeneratorName,
        GeneratorVersion = GeneratorVersion,
        FrameworkDescription = RuntimeInformation.FrameworkDescription,
        RuntimeIdentifier = RuntimeInformation.RuntimeIdentifier,
        OsDescription = RuntimeInformation.OSDescription,
        ProcessArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
        Is64BitProcess = Environment.Is64BitProcess,
        ProcessorCount = Environment.ProcessorCount,
        ServerGc = GCSettings.IsServerGC,
    };

    private static DiskMetrics CreateDiskMetrics(string path)
    {
        string root = Path.GetPathRoot(path) ?? throw new InvalidDataException("The artifact path has no volume root.");
        var drive = new DriveInfo(root);
        return new DiskMetrics
        {
            DriveType = drive.DriveType.ToString(),
            DriveFormat = drive.DriveFormat,
            TotalBytes = drive.TotalSize,
            AvailableFreeBytesAtEnd = drive.AvailableFreeSpace,
        };
    }

    private static int GetChunkSize(IReadOnlyDictionary<string, string> options) => GetInteger(
        options,
        "chunk-size",
        ReviewWorkbookWriteOptions.DefaultStateChunkSizeBytes,
        1,
        ReviewWorkbookFile.MaxStateChunkSizeBytes);

    private static string GetWorkbookOutputPath(IReadOnlyDictionary<string, string> options, string name) =>
        GetOutputPath(options, name, ".xlsx");

    private static string GetOutputPath(
        IReadOnlyDictionary<string, string> options,
        string name,
        string requiredExtension)
    {
        string path = Path.GetFullPath(GetRequired(options, name));
        EnsureNewOutputPath(path, requiredExtension);
        return path;
    }

    private static string? GetOptionalOutputPath(
        IReadOnlyDictionary<string, string> options,
        string name,
        string requiredExtension)
    {
        string? value = GetOptional(options, name);
        if (value == null) return null;
        string path = Path.GetFullPath(value);
        EnsureNewOutputPath(path, requiredExtension);
        return path;
    }

    private static string GetInputPath(
        IReadOnlyDictionary<string, string> options,
        string name,
        string requiredExtension)
    {
        string path = Path.GetFullPath(GetRequired(options, name));
        EnsureInputPath(path, requiredExtension);
        return path;
    }

    private static string? GetOptionalInputPath(
        IReadOnlyDictionary<string, string> options,
        string name,
        string requiredExtension)
    {
        string? value = GetOptional(options, name);
        if (value == null) return null;
        string path = Path.GetFullPath(value);
        EnsureInputPath(path, requiredExtension);
        return path;
    }

    private static void EnsureNewOutputPath(string path, string requiredExtension)
    {
        if (!string.Equals(Path.GetExtension(path), requiredExtension, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("The output path must end in " + requiredExtension + ": " + path);
        string? directory = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
            throw new DirectoryNotFoundException("The output directory does not exist: " + directory);
        EnsureLocalFixedPath(path, directory);
        if (File.Exists(path) || Directory.Exists(path))
            throw new IOException("The output path already exists: " + path);
    }

    private static void EnsureInputPath(string path, string requiredExtension)
    {
        if (!string.Equals(Path.GetExtension(path), requiredExtension, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("The input path must end in " + requiredExtension + ": " + path);
        if (!File.Exists(path)) throw new FileNotFoundException("The input file does not exist.", path);
    }

    private static void EnsureLocalFixedPath(string path, string existingDirectory)
    {
        if (Uri.TryCreate(path, UriKind.Absolute, out Uri? uri) && uri.IsUnc)
            throw new NotSupportedException("Create-new publication is limited to a local fixed drive.");
        string? root = Path.GetPathRoot(path);
        if (string.IsNullOrEmpty(root) || new DriveInfo(root).DriveType != DriveType.Fixed)
            throw new NotSupportedException("Create-new publication is limited to a local fixed drive.");

        DirectoryInfo? current = new(Path.GetFullPath(existingDirectory));
        while (current != null)
        {
            if ((current.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new NotSupportedException(
                    "Create-new publication does not support a reparse-point output directory or ancestor: " +
                    current.FullName);
            current = current.Parent;
        }
    }

    private static void EnsureDistinctPaths(string first, string second)
    {
        if (string.Equals(first, second, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Input and output paths must be distinct.");
    }

    private static string ValidateProjectId(string projectId)
    {
        if (string.IsNullOrWhiteSpace(projectId) || projectId.Length > MaximumProjectIdCharacters ||
            projectId.Any(char.IsControl))
            throw new ArgumentException(
                "--project-id must contain 1 through " +
                MaximumProjectIdCharacters.ToString(CultureInfo.InvariantCulture) +
                " non-control characters.");
        return projectId;
    }

    private static void ValidateFilePrefix(string prefix)
    {
        if (prefix.Length == 0 || prefix.Length > MaximumPrefixCharacters ||
            prefix is "." or ".." ||
            prefix.Any(character =>
                !(char.IsAsciiLetterOrDigit(character) || character is '-' or '_' or '.')))
            throw new ArgumentException(
                "--prefix must contain 1 through " +
                MaximumPrefixCharacters.ToString(CultureInfo.InvariantCulture) +
                " ASCII letters, digits, periods, underscores, or hyphens.");
    }

    private static double Milliseconds(Stopwatch stopwatch) =>
        Math.Round(stopwatch.Elapsed.TotalMilliseconds, 3, MidpointRounding.AwayFromZero);

    private static Dictionary<string, string> ParseOptions(IReadOnlyList<string> args)
    {
        if (args.Count > MaximumArgumentCount)
            throw new ArgumentException("Too many command-line arguments were supplied.");
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < args.Count; index += 2)
        {
            string name = args[index];
            if (!name.StartsWith("--", StringComparison.Ordinal) || name.Length == 2)
                throw new ArgumentException("Expected an option name beginning with --; received: " + name);
            if (index + 1 >= args.Count)
                throw new ArgumentException("Option requires a value: " + name);
            string value = args[index + 1];
            if (value.Length > MaximumOptionValueCharacters)
                throw new ArgumentException("Option value is too long: " + name);
            string key = name[2..];
            if (!result.TryAdd(key, value))
                throw new ArgumentException("Option was specified more than once: " + name);
        }
        return result;
    }

    private static void RequireOnly(IReadOnlyDictionary<string, string> options, params string[] allowed)
    {
        var names = new HashSet<string>(allowed, StringComparer.OrdinalIgnoreCase);
        string? unknown = options.Keys.FirstOrDefault(key => !names.Contains(key));
        if (unknown != null) throw new ArgumentException("Unknown option: --" + unknown);
    }

    private static string GetRequired(IReadOnlyDictionary<string, string> options, string name)
    {
        if (!options.TryGetValue(name, out string? value) || string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Required option is missing: --" + name);
        return value;
    }

    private static string? GetOptional(IReadOnlyDictionary<string, string> options, string name) =>
        options.TryGetValue(name, out string? value) && !string.IsNullOrWhiteSpace(value) ? value : null;

    private static int GetInteger(
        IReadOnlyDictionary<string, string> options,
        string name,
        int defaultValue,
        int minimum,
        int maximum)
    {
        if (!options.TryGetValue(name, out string? text)) return defaultValue;
        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) ||
            value < minimum || value > maximum)
            throw new ArgumentOutOfRangeException(
                name,
                "--" + name + " must be an integer from " + minimum.ToString(CultureInfo.InvariantCulture) +
                " through " + maximum.ToString(CultureInfo.InvariantCulture) + ".");
        return value;
    }

    private static bool IsHelp(string value) =>
        string.Equals(value, "help", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "--help", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "-h", StringComparison.OrdinalIgnoreCase);

    private static void WriteUsage()
    {
        Console.Out.WriteLine("ExcelPowerTools Review.xlsx retention and scale spike");
        Console.Out.WriteLine();
        Console.Out.WriteLine("Generate and immediately validate a workbook:");
        Console.Out.WriteLine("  generate --output <file.xlsx> [--rows <0..1048575>] [--project-id <id>]");
        Console.Out.WriteLine("           [--chunk-size <bytes>] [--state-bytes <0..33554432>]");
        Console.Out.WriteLine("           [--baseline-output <new.json>]");
        Console.Out.WriteLine();
        Console.Out.WriteLine("Generate the canonical 10k/100k/250k scale matrix and baselines:");
        Console.Out.WriteLine("  matrix --output-directory <existing-dir> [--prefix <safe-name>]");
        Console.Out.WriteLine("         [--chunk-size <bytes>] [--state-bytes <1..33554432>]");
        Console.Out.WriteLine();
        Console.Out.WriteLine("Strictly validate an existing workbook:");
        Console.Out.WriteLine("  verify --input <file.xlsx> [--baseline-output <new.json>]");
        Console.Out.WriteLine("  verify --input <file.xlsx> --baseline <before-save.json>");
        Console.Out.WriteLine();
        Console.Out.WriteLine("Successful commands write one compact JSON metrics object to stdout.");
        Console.Out.WriteLine("A semantic retention mismatch writes metrics and exits 3; command errors exit 1.");
        Console.Out.WriteLine("A matrix with one or more failed tiers writes all tier results and exits 4.");
        Console.Out.WriteLine("All workbook and baseline outputs are create-new and are never overwritten.");
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch
        {
        }
    }

    private sealed record CommandExecution(object Metrics, int ExitCode);

    private sealed record MatrixPlan(
        string OutputPath,
        string BaselinePath,
        string ProjectId,
        int RowCount,
        int RequestedStateBytes);

    private sealed record GeneratedArtifact(
        Measurement Measurement,
        string? BaselinePath,
        int RequestedStateBytes,
        int ExpectedChunkCount,
        int ChunkSize,
        Stopwatch Prepare,
        Stopwatch Write,
        Stopwatch BaselineWrite,
        Stopwatch Total,
        long AllocatedBytes,
        long ProcessPeakWorkingSetBytes);

    private sealed record Measurement(
        string Path,
        ReviewWorkbookReadResult Read,
        int RowCount,
        int? DeclaredRowCount,
        int? SnapshotCount,
        string? GeneratorVersion,
        int StateUtf8Bytes,
        long WorkbookBytes,
        string WorkbookSha256,
        int PackageEntryCount,
        long PackageDeclaredUncompressedBytes,
        string LargestPackageEntryName,
        long LargestPackageEntryUncompressedBytes,
        long ChangesWorksheetUncompressedBytes,
        string VisibleProjectionSemanticSha256,
        Stopwatch Validate,
        Stopwatch Inspect,
        Stopwatch Hash,
        Stopwatch Total);

    private sealed record PackageSizeInfo(
        int EntryCount,
        long DeclaredUncompressedBytes,
        string LargestEntryName,
        long LargestEntryUncompressedBytes);

    private sealed record ChangeInspection(
        int RowCount,
        long WorksheetUncompressedBytes,
        string SemanticSha256);

    private sealed record SpikeStateInfo(int? RowCount, int? SnapshotCount, string? GeneratorVersion);

    private sealed record SpikeState(
        string Schema,
        string Generator,
        string GeneratorVersion,
        string ProjectId,
        int RowCount,
        int SnapshotGroupSize,
        int SnapshotRows,
        int SnapshotColumns,
        int SnapshotCount,
        IReadOnlyList<SpikeSnapshot> Snapshots,
        string Padding);

    private sealed record SpikeSnapshot(
        string GroupId,
        int FirstChangeOrdinal,
        IReadOnlyList<string> Before,
        IReadOnlyList<string> After);

    private sealed class RetentionBaseline
    {
        public string Schema { get; set; } = "";
        public string Generator { get; set; } = "";
        public string GeneratorVersion { get; set; } = "";
        public string FileName { get; set; } = "";
        public string ProjectId { get; set; } = "";
        public int RowCount { get; set; }
        public int? DeclaredRowCount { get; set; }
        public int? SnapshotCount { get; set; }
        public int StateUtf8Bytes { get; set; }
        public long WorkbookBytes { get; set; }
        public string WorkbookSha256 { get; set; } = "";
        public string StateSha256 { get; set; } = "";
        public string ProjectionSha256 { get; set; } = "";
        public string VisibleProjectionSemanticSha256 { get; set; } = "";
        public string ManifestItemId { get; set; } = "";
        public int ChunkCount { get; set; }
        public string[] ChunkItemIds { get; set; } = Array.Empty<string>();
        public bool HasDecisionState { get; set; }
        public string DecisionItemId { get; set; } = "";
        public string DecisionSha256 { get; set; } = "";
    }

    private sealed class MatrixMetrics
    {
        public string Schema { get; init; } = "";
        public string Command { get; init; } = "";
        public bool Success { get; init; }
        public string OutputDirectory { get; init; } = "";
        public string Prefix { get; init; } = "";
        public int[] RowCounts { get; init; } = Array.Empty<int>();
        public int[] RequestedStateUtf8Bytes { get; init; } = Array.Empty<int>();
        public int ChunkSizeBytes { get; init; }
        public double TotalElapsedMs { get; init; }
        public TimingMetrics Timings { get; init; } = new();
        public object[] Artifacts { get; init; } = Array.Empty<object>();
    }

    private sealed class MatrixArtifactFailure
    {
        public string Schema { get; init; } = "";
        public string Command { get; init; } = "";
        public bool Success { get; init; }
        public string OutputPath { get; init; } = "";
        public string BaselineOutputPath { get; init; } = "";
        public string ProjectId { get; init; } = "";
        public int RowCount { get; init; }
        public int RequestedStateUtf8Bytes { get; init; }
        public int ChunkSizeBytes { get; init; }
        public string ErrorType { get; init; } = "";
        public string Error { get; init; } = "";
        public bool OutputPathExistsAfterFailure { get; init; }
        public bool BaselinePathExistsAfterFailure { get; init; }
    }

    private sealed class CommandMetrics
    {
        public string Schema { get; init; } = "";
        public string Command { get; init; } = "";
        public bool Success { get; init; }
        public string? OutputPath { get; init; }
        public string? InputPath { get; init; }
        public string? BaselineInputPath { get; init; }
        public string? BaselineOutputPath { get; init; }
        public string ProjectId { get; init; } = "";
        public int RowCount { get; init; }
        public int? DeclaredRowCount { get; init; }
        public int? SnapshotCount { get; init; }
        public int? RequestedStateUtf8Bytes { get; init; }
        public int StateUtf8Bytes { get; init; }
        public long WorkbookBytes { get; init; }
        public string WorkbookSha256 { get; init; } = "";
        public string StateSha256 { get; init; } = "";
        public string ProjectionSha256 { get; init; } = "";
        public string VisibleProjectionSemanticSha256 { get; init; } = "";
        public string ManifestItemId { get; init; } = "";
        public int ChunkCount { get; init; }
        public int? ExpectedChunkCount { get; init; }
        public int? ChunkSizeBytes { get; init; }
        public int PackageEntryCount { get; init; }
        public long PackageDeclaredUncompressedBytes { get; init; }
        public string LargestPackageEntryName { get; init; } = "";
        public long LargestPackageEntryUncompressedBytes { get; init; }
        public long ChangesWorksheetUncompressedBytes { get; init; }
        public double? PrepareElapsedMs { get; init; }
        public double? WriteElapsedMs { get; init; }
        public double ValidateElapsedMs { get; init; }
        public double InspectElapsedMs { get; init; }
        public double HashElapsedMs { get; init; }
        public double BaselineReadElapsedMs { get; init; }
        public double BaselineWriteElapsedMs { get; init; }
        public double TotalElapsedMs { get; init; }
        public TimingMetrics Timings { get; init; } = new();
        public SizeMetrics Sizes { get; init; } = new();
        public HashMetrics Hashes { get; init; } = new();
        public ResourceMetrics Resources { get; init; } = new();
        public RuntimeMetrics Runtime { get; init; } = new();
        public DiskMetrics Disk { get; init; } = new();
        public ValidationMetrics Validation { get; init; } = new();
        public BaselineComparison? Retention { get; init; }
    }

    private sealed class TimingMetrics
    {
        public double? PrepareElapsedMs { get; init; }
        public double? WriteElapsedMs { get; init; }
        public double? ValidateElapsedMs { get; init; }
        public double? InspectElapsedMs { get; init; }
        public double? HashElapsedMs { get; init; }
        public double? BaselineReadElapsedMs { get; init; }
        public double? BaselineWriteElapsedMs { get; init; }
        public double TotalElapsedMs { get; init; }
    }

    private sealed class SizeMetrics
    {
        public int? RequestedStateUtf8Bytes { get; init; }
        public int StateUtf8Bytes { get; init; }
        public long WorkbookBytes { get; init; }
        public long PackageDeclaredUncompressedBytes { get; init; }
        public string LargestPackageEntryName { get; init; } = "";
        public long LargestPackageEntryUncompressedBytes { get; init; }
        public long ChangesWorksheetUncompressedBytes { get; init; }
        public int PackageEntryCount { get; init; }
        public int ProjectedRowCount { get; init; }
        public int? SnapshotCount { get; init; }
        public int ChunkCount { get; init; }
    }

    private sealed class HashMetrics
    {
        public string WorkbookSha256 { get; init; } = "";
        public string StateSha256 { get; init; } = "";
        public string ProjectionSha256 { get; init; } = "";
        public string VisibleProjectionSemanticSha256 { get; init; } = "";
        public string? DecisionSha256 { get; init; }
        public bool RawWorkbookSha256Deterministic { get; init; }
        public bool SemanticHashesComparableAcrossExcelSave { get; init; }
    }

    private sealed class ResourceMetrics
    {
        public long AllocatedBytesForCommand { get; init; }
        public long ProcessPeakWorkingSetBytes { get; init; }
        public long ProcessWorkingSetBytesAtEnd { get; init; }
        public long ManagedHeapBytesAtEnd { get; init; }
    }

    private sealed class RuntimeMetrics
    {
        public string GeneratorName { get; init; } = "";
        public string GeneratorVersion { get; init; } = "";
        public string FrameworkDescription { get; init; } = "";
        public string RuntimeIdentifier { get; init; } = "";
        public string OsDescription { get; init; } = "";
        public string ProcessArchitecture { get; init; } = "";
        public bool Is64BitProcess { get; init; }
        public int ProcessorCount { get; init; }
        public bool ServerGc { get; init; }
    }

    private sealed class DiskMetrics
    {
        public string DriveType { get; init; } = "";
        public string DriveFormat { get; init; } = "";
        public long TotalBytes { get; init; }
        public long AvailableFreeBytesAtEnd { get; init; }
    }

    private sealed class ValidationMetrics
    {
        public bool StrictReaderPassed { get; init; }
        public bool? StateRoundTripMatched { get; init; }
        public bool ProjectedRowCountMatched { get; init; }
        public bool? DeclaredRowCountMatched { get; init; }
        public bool? SnapshotCountMatched { get; init; }
        public bool? RequestedStateSizeSatisfied { get; init; }
        public bool? ChunkCountMatched { get; init; }
        public bool BaselineCompared { get; init; }
        public bool? SemanticRetentionMatched { get; init; }
    }

    private sealed class BaselineComparison
    {
        public bool SemanticRetentionMatched { get; init; }
        public bool ProjectIdMatches { get; init; }
        public bool RowCountMatches { get; init; }
        public bool DeclaredRowCountMatches { get; init; }
        public bool SnapshotCountMatches { get; init; }
        public bool StateUtf8BytesMatches { get; init; }
        public bool StateSha256Matches { get; init; }
        public bool ProjectionSha256Matches { get; init; }
        public bool VisibleProjectionSemanticSha256Matches { get; init; }
        public bool ManifestItemIdMatches { get; init; }
        public bool ChunkCountMatches { get; init; }
        public bool ChunkItemIdsMatch { get; init; }
        public bool DecisionStateMatches { get; init; }
        public bool GeneratorVersionMatches { get; init; }
        public bool PackageSha256Matches { get; init; }
        public long WorkbookBytesDelta { get; init; }
        public long BaselineWorkbookBytes { get; init; }
        public string BaselineWorkbookSha256 { get; init; } = "";
    }
}
