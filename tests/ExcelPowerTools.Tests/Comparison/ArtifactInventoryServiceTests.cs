using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Inventory;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ArtifactInventoryServiceTests
    {
        [Fact]
        public async Task Streams_exact_bytes_reports_progress_and_preserves_the_source()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("large.csv");
            byte[] row = Encoding.UTF8.GetBytes("id,name,value\r\n1,alpha,42\r\n");
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                for (int i = 0; i < 20000; i++)
                    output.Write(row, 0, row.Length);
            }

            byte[] bytesBefore = File.ReadAllBytes(path);
            DateTime lastWriteBefore = File.GetLastWriteTimeUtc(path);
            var updates = new List<ArtifactInventoryProgress>();
            var options = new ArtifactInventoryOptions(
                bufferSize: 4096,
                progressIntervalBytes: 8192,
                contentSampleBytes: 4096);

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(
                path,
                new InlineProgress<ArtifactInventoryProgress>(updates.Add),
                options: options);

            Assert.Equal(Path.GetFullPath(path), result.FullNormalizedPath);
            Assert.Equal(bytesBefore.LongLength, result.ByteLength);
            Assert.Equal(Sha256(bytesBefore), result.ContentSha256);
            Assert.True(StableId.IsSha256(result.InventoryFingerprint));
            Assert.NotEqual(result.ContentSha256, result.InventoryFingerprint);
            Assert.Equal(result.InventoryFingerprint, result.OccurrenceId);
            Assert.Equal(ArtifactFormatKind.Csv, result.Format.EffectiveFormat);
            Assert.True(result.Format.IsSupportedForComparison);
            Assert.True(updates.Count(p => p.Stage == ArtifactInventoryStage.Hashing && p.BytesRead > 0) > 2);
            Assert.Equal(ArtifactInventoryStage.Completed, updates.Last().Stage);
            Assert.Equal(1d, updates.Last().Fraction);
            Assert.Equal(lastWriteBefore, File.GetLastWriteTimeUtc(path));
            Assert.Equal(bytesBefore, File.ReadAllBytes(path));
        }

        [Fact]
        public async Task Same_bytes_at_different_paths_share_content_id_but_not_inventory_fingerprint()
        {
            using var temporary = new TemporaryDirectory();
            byte[] bytes = Encoding.UTF8.GetBytes("key,value\r\nA,1\r\n");
            string firstPath = temporary.PathFor("first.csv");
            string secondPath = temporary.PathFor("second.csv");
            File.WriteAllBytes(firstPath, bytes);
            File.WriteAllBytes(secondPath, bytes);

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord first = await service.InventoryAsync(firstPath);
            ArtifactInventoryRecord second = await service.InventoryAsync(secondPath);

            Assert.Equal(first.ContentSha256, second.ContentSha256);
            Assert.NotEqual(first.InventoryFingerprint, second.InventoryFingerprint);
            Assert.NotEqual(first.FullNormalizedPath, second.FullNormalizedPath);
        }

        [Fact]
        public async Task Reinventorying_an_unchanged_occurrence_is_deterministic()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("stable.tsv");
            File.WriteAllText(path, "id\tvalue\n1\talpha\n", new UTF8Encoding(false));

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord first = await service.InventoryAsync(path);
            ArtifactInventoryRecord second = await service.InventoryAsync(path);

            Assert.Equal(first.InventoryFingerprint, second.InventoryFingerprint);
            Assert.Equal(first.ContentSha256, second.ContentSha256);
            Assert.Equal(first.WeakCreationTimeUtc, second.WeakCreationTimeUtc);
            Assert.Equal(first.WeakLastWriteTimeUtc, second.WeakLastWriteTimeUtc);
        }

        [Fact]
        public async Task Normalizes_dot_segments_to_a_full_path()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("normalized.csv");
            File.WriteAllText(path, "a,b\n1,2\n");
            string withDotSegment = System.IO.Path.Combine(temporary.Root, "unused", "..", "normalized.csv");

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(withDotSegment);

            Assert.Equal(System.IO.Path.GetFullPath(path), result.FullNormalizedPath);
        }

        [Fact]
        public async Task Detects_supported_open_xml_package_formats_from_structure()
        {
            using var temporary = new TemporaryDirectory();
            string xlsx = temporary.PathFor("book.xlsx");
            string xlsm = temporary.PathFor("book.xlsm");
            string xlsb = temporary.PathFor("book.xlsb");
            CreateWorkbookPackage(xlsx, ArtifactFormatKind.Xlsx);
            CreateWorkbookPackage(xlsm, ArtifactFormatKind.Xlsm);
            CreateWorkbookPackage(xlsb, ArtifactFormatKind.Xlsb);

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord xlsxResult = await service.InventoryAsync(xlsx);
            ArtifactInventoryRecord xlsmResult = await service.InventoryAsync(xlsm);
            ArtifactInventoryRecord xlsbResult = await service.InventoryAsync(xlsb);

            AssertFormat(xlsxResult, ArtifactFormatKind.Xlsx, ArtifactFormatEvidenceKind.OpenXmlWorkbookXml);
            AssertFormat(xlsmResult, ArtifactFormatKind.Xlsm, ArtifactFormatEvidenceKind.OpenXmlMacroWorkbook);
            AssertFormat(xlsbResult, ArtifactFormatKind.Xlsb, ArtifactFormatEvidenceKind.OpenXmlBinaryWorkbook);
        }

        [Fact]
        public async Task Accepts_the_workbook_content_type_from_an_applicable_OPC_default()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("default-content-type.xlsx");
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false))
            {
                WriteEntry(
                    archive,
                    "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"xml\" ContentType=\"" +
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"" +
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "</Types>");
                WriteEntry(
                    archive,
                    "_rels/.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" " +
                    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" " +
                    "Target=\"/xl/workbook.xml\"/>" +
                    "</Relationships>");
                WriteEntry(archive, "xl/workbook.xml", "<workbook/>");
                WriteEntry(archive, "xl/worksheets/sheet1.xml", "<worksheet/>");
            }

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(path);

            AssertFormat(result, ArtifactFormatKind.Xlsx, ArtifactFormatEvidenceKind.OpenXmlWorkbookXml);
        }

        [Fact]
        public async Task Rejects_a_ZIP_whose_authoritative_workbook_content_type_is_not_Excel()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("fake.xlsx");
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false))
            {
                WriteEntry(
                    archive,
                    "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"image/jpeg\"/>" +
                    "</Types>");
                WriteEntry(
                    archive,
                    "_rels/.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" " +
                    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" " +
                    "Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                WriteEntry(archive, "xl/workbook.xml", "arbitrary bytes");
            }

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(path);

            Assert.Equal(ArtifactFormatEvidenceKind.InvalidZipPackage, result.Format.EvidenceKind);
            Assert.Equal(ArtifactFormatMatchKind.UnrecognizedContent, result.Format.MatchKind);
            Assert.False(result.Format.IsSupportedForComparison);
        }

        [Fact]
        public async Task Package_inspection_limits_are_not_misreported_as_corruption()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("too-many-entries.xlsx");
            CreateWorkbookPackage(path, ArtifactFormatKind.Xlsx);
            byte[] bytes = File.ReadAllBytes(path);
            int eocd = FindEndOfCentralDirectory(bytes);
            WriteUInt16(bytes, eocd + 8, 10001);
            WriteUInt16(bytes, eocd + 10, 10001);
            File.WriteAllBytes(path, bytes);

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(path);

            Assert.Equal(ArtifactFormatEvidenceKind.InspectionLimitExceeded, result.Format.EvidenceKind);
            Assert.Equal(ArtifactFormatMatchKind.UnrecognizedContent, result.Format.MatchKind);
            Assert.Contains("not being labelled corrupt", result.Format.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Detects_csv_tsv_and_generic_separated_text_content()
        {
            using var temporary = new TemporaryDirectory();
            string csv = temporary.PathFor("data.csv");
            string tsv = temporary.PathFor("data.tsv");
            string text = temporary.PathFor("data.txt");
            string utf16Tsv = temporary.PathFor("unicode.tsv");
            string bomlessUtf16Tsv = temporary.PathFor("bomless-unicode.tsv");
            File.WriteAllText(csv, "id,name\n1,alpha\n", new UTF8Encoding(false));
            File.WriteAllText(tsv, "id\tname\n1\talpha\n", new UTF8Encoding(false));
            File.WriteAllText(text, "id;name\n1;alpha\n", new UTF8Encoding(false));
            File.WriteAllText(utf16Tsv, "id\tname\n1\tβeta\n", Encoding.Unicode);
            File.WriteAllBytes(
                bomlessUtf16Tsv,
                Encoding.Unicode.GetBytes("id\tname\n1\tβeta\n"));

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord csvResult = await service.InventoryAsync(csv);
            ArtifactInventoryRecord tsvResult = await service.InventoryAsync(tsv);
            ArtifactInventoryRecord textResult = await service.InventoryAsync(text);
            ArtifactInventoryRecord utf16Result = await service.InventoryAsync(utf16Tsv);
            ArtifactInventoryRecord bomlessUtf16Result = await service.InventoryAsync(bomlessUtf16Tsv);

            AssertFormat(csvResult, ArtifactFormatKind.Csv, ArtifactFormatEvidenceKind.CommaDelimiter);
            AssertFormat(tsvResult, ArtifactFormatKind.Tsv, ArtifactFormatEvidenceKind.TabDelimiter);
            Assert.Equal(ArtifactFormatKind.DelimitedText, textResult.Format.DeclaredFormat);
            Assert.Equal(ArtifactFormatKind.DelimitedText, textResult.Format.DetectedFormat);
            Assert.Equal(ArtifactFormatMatchKind.Match, textResult.Format.MatchKind);
            AssertFormat(utf16Result, ArtifactFormatKind.Tsv, ArtifactFormatEvidenceKind.TabDelimiter);
            AssertFormat(bomlessUtf16Result, ArtifactFormatKind.Tsv, ArtifactFormatEvidenceKind.TabDelimiter);
        }

        [Fact]
        public async Task One_line_prose_with_a_comma_remains_a_confirmable_text_hint()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("note.txt");
            File.WriteAllText(path, "This is prose, not a separated-value table.");

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(path);

            Assert.Equal(ArtifactFormatKind.Unknown, result.Format.DetectedFormat);
            Assert.Equal(ArtifactFormatMatchKind.DeclaredHintOnly, result.Format.MatchKind);
            Assert.Equal(ArtifactFormatKind.DelimitedText, result.Format.EffectiveFormat);
        }

        [Fact]
        public async Task Txt_accepts_detected_csv_and_semicolon_csv_is_compatible()
        {
            using var temporary = new TemporaryDirectory();
            string txt = temporary.PathFor("export.txt");
            string csv = temporary.PathFor("locale.csv");
            File.WriteAllText(txt, "id,name\n1,alpha\n");
            File.WriteAllText(csv, "id;name\n1;alpha\n");

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord txtResult = await service.InventoryAsync(txt);
            ArtifactInventoryRecord csvResult = await service.InventoryAsync(csv);

            Assert.Equal(ArtifactFormatMatchKind.Compatible, txtResult.Format.MatchKind);
            Assert.Equal(ArtifactFormatKind.Csv, txtResult.Format.DetectedFormat);
            Assert.True(txtResult.Format.IsSupportedForComparison);
            Assert.Equal(ArtifactFormatMatchKind.Compatible, csvResult.Format.MatchKind);
            Assert.Equal(ArtifactFormatKind.DelimitedText, csvResult.Format.DetectedFormat);
            Assert.True(csvResult.Format.IsSupportedForComparison);
        }

        [Fact]
        public async Task Single_column_text_remains_an_explicit_declared_only_hint()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("one-column.csv");
            File.WriteAllText(path, "header\nalpha\nbeta\n");

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(path);

            Assert.Equal(ArtifactFormatKind.Unknown, result.Format.DetectedFormat);
            Assert.Equal(ArtifactFormatKind.Csv, result.Format.EffectiveFormat);
            Assert.Equal(ArtifactFormatMatchKind.DeclaredHintOnly, result.Format.MatchKind);
            Assert.True(result.Format.IsSupportedForComparison);
            Assert.Contains("weak format hint", result.Format.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Reliable_extension_content_mismatch_is_explicit_and_not_silently_supported()
        {
            using var temporary = new TemporaryDirectory();
            string textNamedWorkbook = temporary.PathFor("not-a-workbook.xlsx");
            string workbookNamedTsv = temporary.PathFor("not-text.tsv");
            File.WriteAllText(textNamedWorkbook, "id,name\n1,alpha\n");
            CreateWorkbookPackage(workbookNamedTsv, ArtifactFormatKind.Xlsx);

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord textResult = await service.InventoryAsync(textNamedWorkbook);
            ArtifactInventoryRecord packageResult = await service.InventoryAsync(workbookNamedTsv);

            Assert.Equal(ArtifactFormatMatchKind.Mismatch, textResult.Format.MatchKind);
            Assert.Equal(ArtifactFormatKind.Xlsx, textResult.Format.DeclaredFormat);
            Assert.Equal(ArtifactFormatKind.Csv, textResult.Format.DetectedFormat);
            Assert.True(textResult.Format.HasMismatch);
            Assert.False(textResult.Format.IsSupportedForComparison);
            Assert.Contains("re-export", textResult.Format.Message, StringComparison.OrdinalIgnoreCase);

            Assert.Equal(ArtifactFormatMatchKind.Mismatch, packageResult.Format.MatchKind);
            Assert.Equal(ArtifactFormatKind.Xlsx, packageResult.Format.DetectedFormat);
            Assert.False(packageResult.Format.IsSupportedForComparison);
        }

        [Fact]
        public async Task Unsupported_extension_and_invalid_package_have_actionable_assessments()
        {
            using var temporary = new TemporaryDirectory();
            string unsupported = temporary.PathFor("data.dat");
            string malformedPackage = temporary.PathFor("broken.xlsx");
            File.WriteAllText(unsupported, "id,name\n1,alpha\n");
            File.WriteAllBytes(malformedPackage, new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x01, 0x02, 0x03 });

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord unsupportedResult = await service.InventoryAsync(unsupported);
            ArtifactInventoryRecord malformedResult = await service.InventoryAsync(malformedPackage);

            Assert.Equal(ArtifactFormatMatchKind.UnsupportedExtension, unsupportedResult.Format.MatchKind);
            Assert.Equal(ArtifactFormatKind.Csv, unsupportedResult.Format.DetectedFormat);
            Assert.False(unsupportedResult.Format.IsSupportedForComparison);
            Assert.Contains("Rename or export", unsupportedResult.Format.Message, StringComparison.OrdinalIgnoreCase);

            Assert.Equal(ArtifactFormatEvidenceKind.InvalidZipPackage, malformedResult.Format.EvidenceKind);
            Assert.Equal(ArtifactFormatMatchKind.UnrecognizedContent, malformedResult.Format.MatchKind);
            Assert.False(malformedResult.Format.IsSupportedForComparison);
            Assert.Contains("Recover or re-export", malformedResult.Format.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Compound_office_signature_is_reported_without_trying_Excel()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("encrypted.xlsx");
            File.WriteAllBytes(path, new byte[]
            {
                0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,
                0, 0, 0, 0, 0, 0, 0, 0
            });

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(path);

            Assert.Equal(ArtifactFormatEvidenceKind.CompoundOfficeFile, result.Format.EvidenceKind);
            Assert.Equal(ArtifactFormatMatchKind.UnrecognizedContent, result.Format.MatchKind);
            Assert.Contains("encrypted", result.Format.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Missing_and_folder_paths_return_typed_actionable_errors()
        {
            using var temporary = new TemporaryDirectory();
            var service = new ArtifactInventoryService();

            ArtifactInventoryException missing = await Assert.ThrowsAsync<ArtifactInventoryException>(
                () => service.InventoryAsync(temporary.PathFor("missing.xlsx")));
            ArtifactInventoryException folder = await Assert.ThrowsAsync<ArtifactInventoryException>(
                () => service.InventoryAsync(temporary.Root));

            Assert.Equal(ArtifactInventoryErrorCode.MissingArtifact, missing.ErrorCode);
            Assert.Contains("Verify the path", missing.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(ArtifactInventoryErrorCode.PathIsNotAFile, folder.ErrorCode);
            Assert.Contains("not a file", folder.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Exclusive_lock_returns_unreadable_error_with_corrective_action()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("locked.csv");
            File.WriteAllText(path, "a,b\n1,2\n");
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);

            ArtifactInventoryException error = await Assert.ThrowsAsync<ArtifactInventoryException>(
                () => new ArtifactInventoryService().InventoryAsync(path));

            Assert.Equal(ArtifactInventoryErrorCode.UnreadableArtifact, error.ErrorCode);
            Assert.Contains("exclusive lock", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.NotNull(error.InnerException);
        }

        [Fact]
        public async Task Immutable_read_lease_denies_a_concurrent_writer_and_hash_matches_the_file()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("changing.csv");
            byte[] block = Encoding.UTF8.GetBytes(new string('A', 4090) + ",1\r\n");
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                for (int i = 0; i < 128; i++)
                    output.Write(block, 0, block.Length);
            }

            bool attempted = false;
            bool denied = false;
            var progress = new InlineProgress<ArtifactInventoryProgress>(update =>
            {
                if (attempted || update.Stage != ArtifactInventoryStage.Hashing || update.BytesRead == 0)
                    return;
                attempted = true;
                try
                {
                    using var writer = new FileStream(
                        path,
                        FileMode.Append,
                        FileAccess.Write,
                        FileShare.ReadWrite | FileShare.Delete);
                    writer.WriteByte(0x0A);
                    writer.Flush(true);
                }
                catch (IOException)
                {
                    denied = true;
                }
            });
            var options = new ArtifactInventoryOptions(4096, 4096, 4096);

            ArtifactInventoryRecord result = await new ArtifactInventoryService()
                .InventoryAsync(path, progress, options: options);

            Assert.True(attempted);
            Assert.True(denied);
            Assert.Equal(Sha256(File.ReadAllBytes(path)), result.ContentSha256);
        }

        [Fact]
        public async Task Cancellation_interrupts_a_large_stream_without_wrapping_the_exception()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("cancel.csv");
            File.WriteAllBytes(path, Enumerable.Repeat((byte)'x', 2 * 1024 * 1024).ToArray());
            using var cancellation = new CancellationTokenSource();
            var progress = new InlineProgress<ArtifactInventoryProgress>(update =>
            {
                if (update.Stage == ArtifactInventoryStage.Hashing && update.BytesRead > 0)
                    cancellation.Cancel();
            });
            var options = new ArtifactInventoryOptions(4096, 4096, 4096);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => new ArtifactInventoryService().InventoryAsync(
                    path,
                    progress,
                    cancellation.Token,
                    options));
        }

        [Fact]
        public async Task Cancellation_from_the_completed_notification_does_not_turn_success_into_failure()
        {
            using var temporary = new TemporaryDirectory();
            string path = temporary.PathFor("completed.csv");
            File.WriteAllText(path, "a,b\n1,2\n");
            using var cancellation = new CancellationTokenSource();
            var progress = new InlineProgress<ArtifactInventoryProgress>(update =>
            {
                if (update.Stage == ArtifactInventoryStage.Completed)
                    cancellation.Cancel();
            });

            ArtifactInventoryRecord result = await new ArtifactInventoryService().InventoryAsync(
                path,
                progress,
                cancellation.Token);

            Assert.Equal(Sha256(File.ReadAllBytes(path)), result.ContentSha256);
        }

        [Fact]
        public async Task Empty_workbook_is_invalid_but_empty_csv_remains_a_one_column_hint()
        {
            using var temporary = new TemporaryDirectory();
            string workbook = temporary.PathFor("empty.xlsx");
            string csv = temporary.PathFor("empty.csv");
            File.WriteAllBytes(workbook, Array.Empty<byte>());
            File.WriteAllBytes(csv, Array.Empty<byte>());

            var service = new ArtifactInventoryService();
            ArtifactInventoryRecord workbookResult = await service.InventoryAsync(workbook);
            ArtifactInventoryRecord csvResult = await service.InventoryAsync(csv);

            Assert.Equal(ArtifactFormatMatchKind.UnrecognizedContent, workbookResult.Format.MatchKind);
            Assert.False(workbookResult.Format.IsSupportedForComparison);
            Assert.Equal(ArtifactFormatMatchKind.DeclaredHintOnly, csvResult.Format.MatchKind);
            Assert.True(csvResult.Format.IsSupportedForComparison);
        }

        private static void AssertFormat(
            ArtifactInventoryRecord record,
            ArtifactFormatKind expected,
            ArtifactFormatEvidenceKind evidence)
        {
            Assert.Equal(expected, record.Format.DeclaredFormat);
            Assert.Equal(expected, record.Format.DetectedFormat);
            Assert.Equal(expected, record.Format.EffectiveFormat);
            Assert.Equal(ArtifactFormatMatchKind.Match, record.Format.MatchKind);
            Assert.Equal(evidence, record.Format.EvidenceKind);
            Assert.True(record.Format.IsSupportedForComparison);
        }

        private static void CreateWorkbookPackage(string path, ArtifactFormatKind format)
        {
            using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);
            string workbookPart = format == ArtifactFormatKind.Xlsb ? "xl/workbook.bin" : "xl/workbook.xml";
            string workbookContentType;
            switch (format)
            {
                case ArtifactFormatKind.Xlsm:
                    workbookContentType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
                    break;
                case ArtifactFormatKind.Xlsb:
                    workbookContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
                    break;
                default:
                    workbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
                    break;
            }

            WriteEntry(
                archive,
                "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Override PartName=\"/" + workbookPart + "\" ContentType=\"" + workbookContentType + "\"/>" +
                "</Types>");
            WriteEntry(
                archive,
                "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" " +
                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" " +
                "Target=\"" + workbookPart + "\"/>" +
                "</Relationships>");
            WriteEntry(archive, workbookPart, format == ArtifactFormatKind.Xlsb ? "binary-placeholder" : "<workbook/>");
        }

        private static void WriteEntry(ZipArchive archive, string name, string content)
        {
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
            using Stream stream = entry.Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string Sha256(byte[] bytes)
        {
            using SHA256 sha = SHA256.Create();
            byte[] hash = sha.ComputeHash(bytes);
            return string.Concat(hash.Select(value => value.ToString("X2")));
        }

        private static int FindEndOfCentralDirectory(byte[] bytes)
        {
            for (int offset = bytes.Length - 22; offset >= 0; offset--)
            {
                if (bytes[offset] == 0x50 && bytes[offset + 1] == 0x4B &&
                    bytes[offset + 2] == 0x05 && bytes[offset + 3] == 0x06)
                    return offset;
            }
            throw new InvalidOperationException("Fixture ZIP has no EOCD record.");
        }

        private static void WriteUInt16(byte[] bytes, int offset, ushort value)
        {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private sealed class InlineProgress<T> : IProgress<T>
        {
            private readonly Action<T> _report;

            internal InlineProgress(Action<T> report)
            {
                _report = report;
            }

            public void Report(T value)
            {
                _report(value);
            }
        }

        private sealed class TemporaryDirectory : IDisposable
        {
            internal TemporaryDirectory()
            {
                Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelPowerTools.Inventory.Tests", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(Root);
            }

            internal string Root { get; }

            internal string PathFor(string fileName)
            {
                return System.IO.Path.Combine(Root, fileName);
            }

            public void Dispose()
            {
                try
                {
                    Directory.Delete(Root, recursive: true);
                }
                catch (IOException)
                {
                    // A failed test should not be hidden by best-effort temp cleanup.
                }
                catch (UnauthorizedAccessException)
                {
                    // A failed test should not be hidden by best-effort temp cleanup.
                }
            }
        }
    }
}
