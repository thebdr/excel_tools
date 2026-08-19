using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Xml;
using ExcelPowerTools.Core.Csv;

namespace ExcelPowerTools.Core.Comparison.Inventory
{
    internal static class ArtifactFormatDetector
    {
        private const long MaxContentTypesBytes = 512 * 1024;
        private const long MaxRootRelationshipsBytes = 512 * 1024;
        private const int MaxPackageEntries = 10000;
        private const long MaxCentralDirectoryBytes = 16 * 1024 * 1024;
        private const long MaxDeclaredUncompressedBytes = 8L * 1024 * 1024 * 1024;
        private const string ContentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string PackageRelationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string OfficeDocumentRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string StrictOfficeDocumentRelationshipType =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
        private const string XlsxWorkbookContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string XlsmWorkbookContentType =
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
        private const string XlsbWorkbookContentType =
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main";

        internal static ArtifactFormatAssessment Assess(
            Stream stream,
            byte[] sample,
            string extension,
            CancellationToken cancellationToken)
        {
            string normalizedExtension = (extension ?? "").ToLowerInvariant();
            ArtifactFormatKind declared = DeclaredFormat(normalizedExtension);
            bool supportedExtension = declared != ArtifactFormatKind.Unknown;

            Detection detection = DetectContent(stream, sample ?? Array.Empty<byte>(), cancellationToken);
            return Reconcile(normalizedExtension, declared, supportedExtension, detection);
        }

        private static Detection DetectContent(Stream stream, byte[] sample, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (sample.Length == 0)
                return new Detection(ArtifactFormatKind.Unknown, ArtifactFormatEvidenceKind.EmptyFile);

            if (HasZipSignature(sample))
                return InspectZipPackage(stream, cancellationToken);

            if (HasCompoundFileSignature(sample))
                return new Detection(ArtifactFormatKind.Unknown, ArtifactFormatEvidenceKind.CompoundOfficeFile);

            string text;
            if (!TryDecodeText(sample, out text))
                return new Detection(ArtifactFormatKind.Unknown, ArtifactFormatEvidenceKind.BinaryContent);

            char delimiter;
            if (!TryFindStableDelimiter(text, out delimiter))
                return new Detection(ArtifactFormatKind.Unknown, ArtifactFormatEvidenceKind.TextWithoutStableDelimiter);

            if (delimiter == ',')
                return new Detection(ArtifactFormatKind.Csv, ArtifactFormatEvidenceKind.CommaDelimiter);
            if (delimiter == '\t')
                return new Detection(ArtifactFormatKind.Tsv, ArtifactFormatEvidenceKind.TabDelimiter);

            return new Detection(ArtifactFormatKind.DelimitedText, ArtifactFormatEvidenceKind.OtherDelimiter);
        }

        private static Detection InspectZipPackage(Stream stream, CancellationToken cancellationToken)
        {
            try
            {
                PreflightZipEnvelope(stream);
                stream.Position = 0;
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
                {
                    ZipArchiveEntry contentTypes = null;
                    ZipArchiveEntry rootRelationships = null;
                    bool hasWorkbookXml = false;
                    bool hasWorkbookBin = false;
                    bool hasVbaProject = false;
                    int entryCount = 0;
                    long totalDeclaredLength = 0;

                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        entryCount++;
                        if (entryCount > MaxPackageEntries)
                            throw new InspectionLimitException(
                                "The package contains too many entries to inspect safely.");
                        if (entry.Length < 0 || entry.Length > MaxDeclaredUncompressedBytes - totalDeclaredLength)
                            throw new InspectionLimitException(
                                "The package declares an excessive uncompressed size.");
                        totalDeclaredLength += entry.Length;
                        string name = entry.FullName;
                        if (name.IndexOf('\\') >= 0)
                            throw new InvalidDataException("OPC package part names must use forward slashes.");
                        if (IsCriticalPartNameWithWrongCase(name))
                            throw new InvalidDataException("A critical OPC package part name has invalid casing.");
                        if (string.Equals(name, "[Content_Types].xml", StringComparison.Ordinal))
                        {
                            if (contentTypes != null)
                                throw new InvalidDataException("The package contains duplicate content-type manifests.");
                            contentTypes = entry;
                        }
                        else if (string.Equals(name, "_rels/.rels", StringComparison.Ordinal))
                        {
                            if (rootRelationships != null)
                                throw new InvalidDataException("The package contains duplicate root relationships.");
                            rootRelationships = entry;
                        }
                        else if (string.Equals(name, "xl/workbook.xml", StringComparison.Ordinal))
                        {
                            if (hasWorkbookXml)
                                throw new InvalidDataException("The package contains duplicate workbook parts.");
                            hasWorkbookXml = true;
                        }
                        else if (string.Equals(name, "xl/workbook.bin", StringComparison.Ordinal))
                        {
                            if (hasWorkbookBin)
                                throw new InvalidDataException("The package contains duplicate binary workbook parts.");
                            hasWorkbookBin = true;
                        }
                        else if (string.Equals(name, "xl/vbaProject.bin", StringComparison.Ordinal))
                            hasVbaProject = true;
                    }

                    if (contentTypes == null || rootRelationships == null || hasWorkbookXml == hasWorkbookBin)
                        return new Detection(ArtifactFormatKind.Unknown, ArtifactFormatEvidenceKind.ZipWithoutWorkbook);

                    string workbookPartName = hasWorkbookBin ? "/xl/workbook.bin" : "/xl/workbook.xml";
                    string relationshipTarget = ReadRootWorkbookTarget(
                        rootRelationships,
                        cancellationToken);
                    if (!string.Equals(
                            relationshipTarget,
                            workbookPartName.Substring(1),
                            StringComparison.Ordinal))
                        throw new InvalidDataException(
                            "The root officeDocument relationship does not target the workbook part.");
                    ArtifactFormatKind contentTypeFormat = ReadWorkbookContentType(
                        contentTypes,
                        workbookPartName,
                        cancellationToken);
                    if (contentTypeFormat == ArtifactFormatKind.Unknown ||
                        (hasWorkbookBin && contentTypeFormat != ArtifactFormatKind.Xlsb) ||
                        (hasWorkbookXml && contentTypeFormat == ArtifactFormatKind.Xlsb) ||
                        (hasVbaProject && contentTypeFormat == ArtifactFormatKind.Xlsx))
                        throw new InvalidDataException("The workbook part and authoritative content type disagree.");

                    switch (contentTypeFormat)
                    {
                        case ArtifactFormatKind.Xlsb:
                            return new Detection(
                                ArtifactFormatKind.Xlsb,
                                ArtifactFormatEvidenceKind.OpenXmlBinaryWorkbook);
                        case ArtifactFormatKind.Xlsm:
                            return new Detection(
                                ArtifactFormatKind.Xlsm,
                                ArtifactFormatEvidenceKind.OpenXmlMacroWorkbook);
                        default:
                            return new Detection(
                                ArtifactFormatKind.Xlsx,
                                ArtifactFormatEvidenceKind.OpenXmlWorkbookXml);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (InspectionLimitException)
            {
                return new Detection(
                    ArtifactFormatKind.Unknown,
                    ArtifactFormatEvidenceKind.InspectionLimitExceeded);
            }
            catch (InvalidDataException)
            {
                return new Detection(ArtifactFormatKind.Unknown, ArtifactFormatEvidenceKind.InvalidZipPackage);
            }
            catch (XmlException)
            {
                return new Detection(ArtifactFormatKind.Unknown, ArtifactFormatEvidenceKind.InvalidZipPackage);
            }
        }

        private static ArtifactFormatKind ReadWorkbookContentType(
            ZipArchiveEntry contentTypes,
            string workbookPartName,
            CancellationToken cancellationToken)
        {
            if (contentTypes.Length < 0 || contentTypes.Length > MaxContentTypesBytes)
                throw new InspectionLimitException(
                    "The package content-type manifest exceeds the inventory inspection limit.");

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                MaxCharactersInDocument = MaxContentTypesBytes
            };

            using (Stream entryStream = contentTypes.Open())
            using (XmlReader reader = XmlReader.Create(entryStream, settings))
            {
                ArtifactFormatKind overrideFormat = ArtifactFormatKind.Unknown;
                ArtifactFormatKind defaultFormat = ArtifactFormatKind.Unknown;
                bool sawOverride = false;
                bool sawApplicableDefault = false;
                string workbookExtension = Path.GetExtension(workbookPartName).TrimStart('.');
                while (reader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType != XmlNodeType.Element ||
                        !string.Equals(reader.NamespaceURI, ContentTypesNamespace, StringComparison.Ordinal))
                        continue;

                    if (string.Equals(reader.LocalName, "Override", StringComparison.Ordinal))
                    {
                        string partName = reader.GetAttribute("PartName") ?? "";
                        if (!string.Equals(partName, workbookPartName, StringComparison.Ordinal)) continue;
                        if (sawOverride)
                            throw new InvalidDataException(
                                "The workbook part has duplicate content-type declarations.");
                        sawOverride = true;
                        overrideFormat = ClassifyWorkbookContentType(reader.GetAttribute("ContentType") ?? "");
                        continue;
                    }

                    if (!string.Equals(reader.LocalName, "Default", StringComparison.Ordinal)) continue;
                    string extension = reader.GetAttribute("Extension") ?? "";
                    if (!string.Equals(extension, workbookExtension, StringComparison.OrdinalIgnoreCase)) continue;
                    if (sawApplicableDefault)
                        throw new InvalidDataException(
                            "The package has duplicate default content types for the workbook extension.");
                    sawApplicableDefault = true;
                    defaultFormat = ClassifyWorkbookContentType(reader.GetAttribute("ContentType") ?? "");
                }

                // OPC Overrides take precedence. The Open XML SDK can legitimately compact a package by using
                // a Default for .xml and explicit Overrides for every non-workbook XML part, so an applicable
                // Default is authoritative when the workbook has no specific Override.
                return sawOverride ? overrideFormat : defaultFormat;
            }
        }

        private static ArtifactFormatKind ClassifyWorkbookContentType(string contentType)
        {
            if (string.Equals(contentType, XlsxWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                return ArtifactFormatKind.Xlsx;
            if (string.Equals(contentType, XlsmWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                return ArtifactFormatKind.Xlsm;
            if (string.Equals(contentType, XlsbWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                return ArtifactFormatKind.Xlsb;
            return ArtifactFormatKind.Unknown;
        }

        private static string ReadRootWorkbookTarget(
            ZipArchiveEntry relationships,
            CancellationToken cancellationToken)
        {
            if (relationships.Length < 0 || relationships.Length > MaxRootRelationshipsBytes)
                throw new InspectionLimitException(
                    "The root relationship part exceeds the inventory inspection limit.");

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                MaxCharactersInDocument = MaxRootRelationshipsBytes,
            };
            string target = null;
            using (Stream stream = relationships.Open())
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                while (reader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType != XmlNodeType.Element ||
                        reader.LocalName != "Relationship" ||
                        reader.NamespaceURI != PackageRelationshipsNamespace)
                        continue;
                    string type = reader.GetAttribute("Type") ?? "";
                    if (type != OfficeDocumentRelationshipType &&
                        type != StrictOfficeDocumentRelationshipType)
                        continue;
                    if (!string.IsNullOrEmpty(reader.GetAttribute("TargetMode")) &&
                        !string.Equals(reader.GetAttribute("TargetMode"), "Internal", StringComparison.Ordinal))
                        throw new InvalidDataException("The root workbook relationship must be internal.");
                    string current = reader.GetAttribute("Target") ?? "";
                    if (current.StartsWith("/", StringComparison.Ordinal)) current = current.Substring(1);
                    if (current.Length == 0 || current.IndexOf('\\') >= 0 || current.Contains(".."))
                        throw new InvalidDataException("The root workbook relationship target is invalid.");
                    if (target != null)
                        throw new InvalidDataException("The package has more than one officeDocument relationship.");
                    target = current;
                }
            }

            if (target == null)
                throw new InvalidDataException("The package has no root officeDocument relationship.");
            return target;
        }

        private static bool IsCriticalPartNameWithWrongCase(string name)
        {
            string[] critical =
            {
                "[Content_Types].xml",
                "_rels/.rels",
                "xl/workbook.xml",
                "xl/workbook.bin",
                "xl/vbaProject.bin",
            };
            foreach (string expected in critical)
            {
                if (string.Equals(name, expected, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(name, expected, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        private static ArtifactFormatAssessment Reconcile(
            string extension,
            ArtifactFormatKind declared,
            bool supportedExtension,
            Detection detection)
        {
            if (!supportedExtension)
            {
                string detected = detection.Format == ArtifactFormatKind.Unknown
                    ? "No supported workbook or separated-value content was detected."
                    : "The content resembles " + Display(detection.Format) + ".";
                return new ArtifactFormatAssessment(
                    extension,
                    declared,
                    detection.Format,
                    detection.Format,
                    ArtifactFormatMatchKind.UnsupportedExtension,
                    detection.Evidence,
                    false,
                    "The extension '" + (extension.Length == 0 ? "(none)" : extension) +
                    "' is not supported. " + detected + " Rename or export the artifact as .xlsx, .xlsm, .xlsb, .csv, .tsv, or .txt, then inventory it again.");
            }

            if (detection.Format == declared)
            {
                return new ArtifactFormatAssessment(
                    extension,
                    declared,
                    detection.Format,
                    effectiveFormat: detection.Format,
                    matchKind: ArtifactFormatMatchKind.Match,
                    evidenceKind: detection.Evidence,
                    isSupportedForComparison: true,
                    message: "The file extension and content evidence both identify " + Display(declared) + ".");
            }

            if (AreCompatibleSeparatedFormats(declared, detection.Format))
            {
                return new ArtifactFormatAssessment(
                    extension,
                    declared,
                    detection.Format,
                    detection.Format,
                    ArtifactFormatMatchKind.Compatible,
                    detection.Evidence,
                    true,
                    "The content uses " + Display(detection.Format) +
                    " delimiters compatible with the declared " + Display(declared) + " artifact.");
            }

            if (detection.Format == ArtifactFormatKind.Unknown)
            {
                bool textHintCanStandAlone = IsSeparatedFormat(declared) &&
                    (detection.Evidence == ArtifactFormatEvidenceKind.TextWithoutStableDelimiter ||
                     detection.Evidence == ArtifactFormatEvidenceKind.EmptyFile);
                if (textHintCanStandAlone)
                {
                    return new ArtifactFormatAssessment(
                        extension,
                        declared,
                        ArtifactFormatKind.Unknown,
                        declared,
                        ArtifactFormatMatchKind.DeclaredHintOnly,
                        detection.Evidence,
                        true,
                        "No stable delimiter was present in the bounded sample. The supported " + Display(declared) +
                        " extension remains a weak format hint; confirm the import profile before comparing.");
                }

                return new ArtifactFormatAssessment(
                    extension,
                    declared,
                    ArtifactFormatKind.Unknown,
                    ArtifactFormatKind.Unknown,
                    ArtifactFormatMatchKind.UnrecognizedContent,
                    detection.Evidence,
                    false,
                    UnrecognizedMessage(declared, detection.Evidence));
            }

            return new ArtifactFormatAssessment(
                extension,
                declared,
                detection.Format,
                detection.Format,
                ArtifactFormatMatchKind.Mismatch,
                detection.Evidence,
                false,
                "The '" + extension + "' extension declares " + Display(declared) +
                ", but content evidence identifies " + Display(detection.Format) +
                ". Verify the source and rename or re-export the file before comparing it.");
        }

        private static string UnrecognizedMessage(ArtifactFormatKind declared, ArtifactFormatEvidenceKind evidence)
        {
            switch (evidence)
            {
                case ArtifactFormatEvidenceKind.EmptyFile:
                    return "The artifact is empty and cannot be a valid " + Display(declared) +
                           " workbook. Restore or re-export the source file.";
                case ArtifactFormatEvidenceKind.CompoundOfficeFile:
                    return "The artifact has an OLE compound-file signature. It may be encrypted or an older Excel format; decrypt or export it to a supported unencrypted format.";
                case ArtifactFormatEvidenceKind.InvalidZipPackage:
                    return "The artifact starts like an Open XML package but its ZIP/package structure is invalid. Recover or re-export the workbook before comparing it.";
                case ArtifactFormatEvidenceKind.ZipWithoutWorkbook:
                    return "The ZIP package does not contain a recognizable Excel workbook part. Verify that the selected file is an Excel workbook.";
                case ArtifactFormatEvidenceKind.BinaryContent:
                    return "The artifact contains binary data that does not match its declared format. Verify the file and export it to a supported format.";
                case ArtifactFormatEvidenceKind.InspectionLimitExceeded:
                    return "The artifact exceeds the safe in-process package inspection limits. " +
                           "Use the isolated/large-file comparison path or reduce the package complexity; " +
                           "the source is not being labelled corrupt.";
                default:
                    return "The artifact content does not match a supported " + Display(declared) +
                           " representation. Verify the source and export it again.";
            }
        }

        private static ArtifactFormatKind DeclaredFormat(string extension)
        {
            switch (extension)
            {
                case ".xlsx": return ArtifactFormatKind.Xlsx;
                case ".xlsm": return ArtifactFormatKind.Xlsm;
                case ".xlsb": return ArtifactFormatKind.Xlsb;
                case ".csv": return ArtifactFormatKind.Csv;
                case ".tsv": return ArtifactFormatKind.Tsv;
                case ".txt": return ArtifactFormatKind.DelimitedText;
                default: return ArtifactFormatKind.Unknown;
            }
        }

        private static bool AreCompatibleSeparatedFormats(ArtifactFormatKind declared, ArtifactFormatKind detected)
        {
            if (declared == ArtifactFormatKind.DelimitedText && IsSeparatedFormat(detected))
                return true;

            // Semicolon/pipe separated files commonly carry the .csv extension in locale-specific exports.
            return declared == ArtifactFormatKind.Csv && detected == ArtifactFormatKind.DelimitedText;
        }

        private static bool IsSeparatedFormat(ArtifactFormatKind format)
        {
            return format == ArtifactFormatKind.Csv ||
                   format == ArtifactFormatKind.Tsv ||
                   format == ArtifactFormatKind.DelimitedText;
        }

        private static string Display(ArtifactFormatKind format)
        {
            switch (format)
            {
                case ArtifactFormatKind.Xlsx: return "XLSX";
                case ArtifactFormatKind.Xlsm: return "XLSM";
                case ArtifactFormatKind.Xlsb: return "XLSB";
                case ArtifactFormatKind.Csv: return "comma-separated text";
                case ArtifactFormatKind.Tsv: return "tab-separated text";
                case ArtifactFormatKind.DelimitedText: return "separated text";
                default: return "unknown content";
            }
        }

        private static bool HasZipSignature(byte[] sample)
        {
            return sample.Length >= 4 && sample[0] == 0x50 && sample[1] == 0x4B &&
                   ((sample[2] == 0x03 && sample[3] == 0x04) ||
                    (sample[2] == 0x05 && sample[3] == 0x06) ||
                    (sample[2] == 0x07 && sample[3] == 0x08));
        }

        private static void PreflightZipEnvelope(Stream stream)
        {
            if (!stream.CanSeek || stream.Length < 22)
                throw new InvalidDataException("The ZIP end-of-central-directory record is missing.");
            int tailLength = (int)Math.Min(stream.Length, 65557);
            var tail = new byte[tailLength];
            stream.Position = stream.Length - tailLength;
            int read = 0;
            while (read < tail.Length)
            {
                int current = stream.Read(tail, read, tail.Length - read);
                if (current == 0)
                    throw new InvalidDataException("The ZIP central directory is truncated.");
                read += current;
            }

            for (int offset = tail.Length - 22; offset >= 0; offset--)
            {
                if (ReadUInt32(tail, offset) != 0x06054B50U) continue;
                ushort commentLength = ReadUInt16(tail, offset + 20);
                if (offset + 22 + commentLength != tail.Length) continue;
                ushort entryCount = ReadUInt16(tail, offset + 10);
                uint centralDirectorySize = ReadUInt32(tail, offset + 12);
                if (entryCount == ushort.MaxValue || centralDirectorySize == uint.MaxValue)
                    throw new InspectionLimitException(
                        "ZIP64 packages exceed the bounded inventory signature inspection envelope.");
                if (entryCount > MaxPackageEntries || centralDirectorySize > MaxCentralDirectoryBytes)
                    throw new InspectionLimitException(
                        "The ZIP central directory exceeds the bounded inventory inspection envelope.");
                return;
            }

            throw new InvalidDataException("The ZIP end-of-central-directory record is invalid.");
        }

        private static ushort ReadUInt16(byte[] bytes, int offset)
        {
            return (ushort)(bytes[offset] | bytes[offset + 1] << 8);
        }

        private static uint ReadUInt32(byte[] bytes, int offset)
        {
            return (uint)(bytes[offset] |
                          bytes[offset + 1] << 8 |
                          bytes[offset + 2] << 16 |
                          bytes[offset + 3] << 24);
        }

        private static bool HasCompoundFileSignature(byte[] sample)
        {
            byte[] signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            if (sample.Length < signature.Length)
                return false;
            for (int i = 0; i < signature.Length; i++)
            {
                if (sample[i] != signature[i])
                    return false;
            }
            return true;
        }

        private static bool TryDecodeText(byte[] sample, out string text)
        {
            try
            {
                var detected = CsvEncoding.Detect(sample);
                text = CsvEncoding.DecodeText(sample, detected.encoding, detected.hasBom);
            }
            catch (DecoderFallbackException)
            {
                text = "";
                return false;
            }

            int disallowedControls = 0;
            int nullCharacters = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                if (value == '\0') nullCharacters++;
                if (value < 0x20 && value != '\r' && value != '\n' && value != '\t')
                    disallowedControls++;
            }

            if (nullCharacters > 0 || disallowedControls > Math.Max(4, text.Length / 50))
            {
                text = "";
                return false;
            }
            return true;
        }

        private static bool TryFindStableDelimiter(string text, out char delimiter)
        {
            char[] candidates = { '\t', ',', ';', '|' };
            var records = new List<int[]>();
            var counts = new int[candidates.Length];
            bool quoted = false;
            bool hasContent = false;

            for (int i = 0; i < text.Length && records.Count < 32; i++)
            {
                char c = text[i];
                if (c == '"')
                {
                    if (quoted && i + 1 < text.Length && text[i + 1] == '"')
                    {
                        i++;
                        hasContent = true;
                        continue;
                    }
                    quoted = !quoted;
                    hasContent = true;
                    continue;
                }

                if (!quoted)
                {
                    for (int candidateIndex = 0; candidateIndex < candidates.Length; candidateIndex++)
                    {
                        if (c == candidates[candidateIndex])
                            counts[candidateIndex]++;
                    }

                    if (c == '\r' || c == '\n')
                    {
                        if (hasContent || HasAnyCount(counts))
                            records.Add((int[])counts.Clone());
                        Array.Clear(counts, 0, counts.Length);
                        hasContent = false;
                        if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                            i++;
                        continue;
                    }
                }

                if (!char.IsWhiteSpace(c))
                    hasContent = true;
            }

            if ((hasContent || HasAnyCount(counts)) && records.Count < 32)
                records.Add((int[])counts.Clone());

            // One line with punctuation is weak evidence (for example prose containing a comma). Keep the
            // declared extension as a confirmable hint instead of silently selecting a parser.
            if (records.Count < 2)
            {
                delimiter = '\0';
                return false;
            }

            long bestScore = -1;
            int bestCandidate = -1;
            for (int candidateIndex = 0; candidateIndex < candidates.Length; candidateIndex++)
            {
                var frequencies = new Dictionary<int, int>();
                foreach (int[] record in records)
                {
                    int count = record[candidateIndex];
                    if (count <= 0)
                        continue;
                    int frequency;
                    frequencies.TryGetValue(count, out frequency);
                    frequencies[count] = frequency + 1;
                }

                int mode = 0;
                int modeFrequency = 0;
                foreach (KeyValuePair<int, int> pair in frequencies)
                {
                    if (pair.Value > modeFrequency || (pair.Value == modeFrequency && pair.Key > mode))
                    {
                        mode = pair.Key;
                        modeFrequency = pair.Value;
                    }
                }

                int minimumFrequency = Math.Max(2, (records.Count * 4 + 4) / 5);
                if (modeFrequency < minimumFrequency || mode == 0)
                    continue;

                // Earlier candidates win an otherwise exact tie: tab, comma, semicolon, then pipe.
                long score = modeFrequency * 100000L + mode * 100L + (candidates.Length - candidateIndex);
                if (score > bestScore)
                {
                    bestScore = score;
                    bestCandidate = candidateIndex;
                }
            }

            if (bestCandidate < 0)
            {
                delimiter = '\0';
                return false;
            }

            delimiter = candidates[bestCandidate];
            return true;
        }

        private static bool HasAnyCount(int[] counts)
        {
            for (int i = 0; i < counts.Length; i++)
            {
                if (counts[i] != 0)
                    return true;
            }
            return false;
        }

        private struct Detection
        {
            internal Detection(ArtifactFormatKind format, ArtifactFormatEvidenceKind evidence)
            {
                Format = format;
                Evidence = evidence;
            }

            internal ArtifactFormatKind Format { get; }
            internal ArtifactFormatEvidenceKind Evidence { get; }
        }

        private sealed class InspectionLimitException : IOException
        {
            internal InspectionLimitException(string message) : base(message)
            {
            }
        }
    }
}
