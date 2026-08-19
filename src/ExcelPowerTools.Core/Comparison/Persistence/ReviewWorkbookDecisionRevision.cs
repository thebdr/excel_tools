using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using ExcelPowerTools.Core.Comparison.Decisions;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    public static partial class ReviewWorkbookFile
    {
        /// <summary>
        /// Creates a new immutable review-workbook revision containing one additional decision event. The
        /// validated source is never modified and an existing destination is never replaced.
        /// </summary>
        public static void CreateDecisionRevision(
            string sourcePath,
            string destinationPath,
            ReviewDecisionRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            CreateDecisionRevisionBatch(
                sourcePath,
                destinationPath,
                new[] { request });
        }

        /// <summary>
        /// Creates one new immutable Review workbook containing an ordered batch of already chained decision
        /// requests. The source is validated once, every request must target the generation produced by its
        /// predecessor, and only the final fully replayable generation is published.
        /// </summary>
        public static void CreateDecisionRevisionBatch(
            string sourcePath,
            string destinationPath,
            IEnumerable<ReviewDecisionRequest> requests)
        {
            CreateDecisionRevisionBatchCore(
                sourcePath,
                destinationPath,
                requests,
                null);
        }

        /// <summary>
        /// Creates one new immutable Review workbook from the exact hydrated source occurrence. Package bytes,
        /// stable file identity and all validated state identities are checked again on the same protected handle
        /// used for derivation, closing the interaction-boundary replacement race.
        /// </summary>
        public static void CreateDecisionRevisionBatch(
            ReviewWorkbookHydrationSession expectedSource,
            string destinationPath,
            IEnumerable<ReviewDecisionRequest> requests)
        {
            if (expectedSource == null) throw new ArgumentNullException(nameof(expectedSource));
            CreateDecisionRevisionBatchCore(
                expectedSource.FullPath,
                destinationPath,
                requests,
                expectedSource);
        }

        private static void CreateDecisionRevisionBatchCore(
            string sourcePath,
            string destinationPath,
            IEnumerable<ReviewDecisionRequest> requests,
            ReviewWorkbookHydrationSession expectedSource)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
                throw new ArgumentException("A source workbook path is required.", nameof(sourcePath));
            if (string.IsNullOrWhiteSpace(destinationPath))
                throw new ArgumentException("A destination path is required.", nameof(destinationPath));
            if (requests == null) throw new ArgumentNullException(nameof(requests));
            var frozenRequests = new List<ReviewDecisionRequest>(requests);
            if (frozenRequests.Count == 0)
                throw new ArgumentException(
                    "A decision revision batch requires at least one request.",
                    nameof(requests));
            if (frozenRequests.Count > MaxDecisionEventCount)
                throw new ArgumentException(
                    "The decision revision batch exceeds the event quota.",
                    nameof(requests));
            if (frozenRequests.Any(item => item == null))
                throw new ArgumentException(
                    "A decision revision batch cannot contain a null request.",
                    nameof(requests));

            string source = Path.GetFullPath(sourcePath);
            string destination = Path.GetFullPath(destinationPath);
            if (!string.Equals(Path.GetExtension(source), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    "A review workbook source must end in .xlsx.",
                    nameof(sourcePath));
            if (!string.Equals(Path.GetExtension(destination), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    "A review workbook destination must end in .xlsx.",
                    nameof(destinationPath));
            if (string.Equals(source, destination, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    "The decision revision destination must differ from its source.",
                    nameof(destinationPath));
            EnsureSupportedPublicationPath(destination);

            string directory = Path.GetDirectoryName(destination);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                throw new DirectoryNotFoundException(
                    "The review workbook destination directory does not exist.");
            if (File.Exists(destination) || Directory.Exists(destination))
                throw new IOException("The review workbook destination already exists: " + destination);

            string temporary = Path.Combine(
                directory,
                "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                // FileShare.Read deliberately withholds write and delete sharing for the complete derivation.
                // This binds validation, optimistic concurrency, copying and publication to one source object.
                using (var sourceStream = new FileStream(
                           source,
                           FileMode.Open,
                           FileAccess.Read,
                           FileShare.Read,
                           64 * 1024,
                           FileOptions.SequentialScan))
                {
                    ReviewWorkbookReadResult current = ReadAndValidate(sourceStream);
                    if (!current.HasDecisionState)
                        throw new InvalidDataException(
                            "The source review workbook has no validated decision sidecar to revise.");
                    long validatedSourceLength = sourceStream.Length;
                    if (validatedSourceLength < 0 || validatedSourceLength > MaxReviewPackageBytes)
                        throw new InvalidDataException(
                            "The validated source length exceeds the safe review-package limit.");
                    if (expectedSource != null)
                    {
                        string packageSha256 =
                            ReviewWorkbookHydrationService.ComputeBoundedPackageSha256(
                                sourceStream,
                                validatedSourceLength);
                        ReviewWorkbookHandleMetadata metadata =
                            ReviewWorkbookHandleMetadataReader.Capture(sourceStream, source);
                        EnsureExpectedSourceOccurrence(
                            expectedSource,
                            source,
                            validatedSourceLength,
                            packageSha256,
                            metadata,
                            current);
                    }

                    if (current.DecisionGeneration.Events.Count + frozenRequests.Count >
                        MaxDecisionEventCount)
                        throw new InvalidDataException(
                            "The decision revision batch would exceed the durable event quota.");
                    ReviewDecisionGeneration next = current.DecisionGeneration;
                    foreach (ReviewDecisionRequest request in frozenRequests)
                        next = ReviewDecisionEngine.Apply(next, request);

                    sourceStream.Position = 0;
                    using (var stage = new FileStream(
                               temporary,
                               FileMode.CreateNew,
                               FileAccess.Write,
                               FileShare.None,
                               64 * 1024,
                               FileOptions.SequentialScan))
                    {
                        CopyValidatedSourceBytes(
                            sourceStream,
                            stage,
                            validatedSourceLength);
                    }

                    string decisionSha256 = RewriteDecisionRevisionParts(
                        temporary,
                        current,
                        next);
                    using (var verifiedStage = new FileStream(
                               temporary,
                               FileMode.Open,
                               FileAccess.ReadWrite,
                               FileShare.None,
                               4096,
                               FileOptions.WriteThrough))
                    {
                        ReviewWorkbookReadResult revised = ReadAndValidate(verifiedStage);
                        VerifyDecisionRevision(current, next, revised, decisionSha256);
                        if (sourceStream.Length != validatedSourceLength)
                            throw new InvalidDataException(
                                "The source review workbook length changed before publication.");
                        verifiedStage.Flush(true);
                        if (sourceStream.Length != validatedSourceLength)
                            throw new InvalidDataException(
                                "The source review workbook length changed during durable staging.");
                    }
                }

                // Same-directory File.Move is the create-new publication boundary and the final fallible
                // operation after every source/package handle is closed. A racing publisher wins; the loser
                // receives IOException and its private temporary is removed by the catch below. A portable path
                // API still leaves a tiny close-to-rename window in a directory controlled by a malicious local
                // writer; a later Windows-specific handle-rename spike can close that residual gap.
                File.Move(temporary, destination);
            }
            catch
            {
                TryDelete(temporary);
                throw;
            }
        }

        private static void EnsureExpectedSourceOccurrence(
            ReviewWorkbookHydrationSession expected,
            string sourcePath,
            long sourceLength,
            string packageSha256,
            ReviewWorkbookHandleMetadata metadata,
            ReviewWorkbookReadResult current)
        {
            bool currentHasStableToken = metadata.FileOccurrenceToken.Length != 0;
            bool matches =
                string.Equals(
                    expected.FullPath,
                    sourcePath,
                    StringComparison.OrdinalIgnoreCase) &&
                expected.FileByteLength == sourceLength &&
                metadata.FileByteLength == sourceLength &&
                FixedTimeEquals(expected.PackageSha256, packageSha256) &&
                expected.HasStableOccurrenceToken == currentHasStableToken &&
                (!expected.HasStableOccurrenceToken || string.Equals(
                    expected.FileOccurrenceToken,
                    metadata.FileOccurrenceToken,
                    StringComparison.Ordinal)) &&
                ReviewWorkbookHydrationService.HasSameValidatedIdentity(
                    expected.ReadResult,
                    current);
            if (!matches)
                throw new InvalidDataException(
                    "The selected Review workbook no longer matches the hydrated source occurrence. " +
                    "Reload it before saving decisions; no revision was published.");
        }

        private static void CopyValidatedSourceBytes(
            Stream source,
            Stream stage,
            long validatedSourceLength)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (stage == null) throw new ArgumentNullException(nameof(stage));
            if (!source.CanRead || !source.CanSeek || !stage.CanWrite || !stage.CanSeek)
                throw new InvalidDataException(
                    "Decision revision copying requires readable and writable seekable streams.");
            if (validatedSourceLength < 0 || validatedSourceLength > MaxReviewPackageBytes)
                throw new InvalidDataException(
                    "The validated source length exceeds the safe review-package limit.");

            long stageStart = stage.Position;
            using (var bounded = new BoundedReadStream(source, MaxReviewPackageBytes))
                bounded.CopyTo(stage, 64 * 1024);
            long copiedLength = stage.Position - stageStart;
            if (copiedLength != validatedSourceLength || source.Length != validatedSourceLength)
                throw new InvalidDataException(
                    "The source review workbook length changed after validation.");
        }

        private static string RewriteDecisionRevisionParts(
            string temporaryPath,
            ReviewWorkbookReadResult current,
            ReviewDecisionGeneration next)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(
                       temporaryPath,
                       true,
                       CreateBoundedOpenSettings()))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                    throw new InvalidDataException("The copied package has no workbook part.");

                CustomXmlPart decisionPart = FindOwnedCustomXmlPart(
                    workbookPart,
                    current.DecisionItemId,
                    DecisionCustomXmlNamespace,
                    "decisionState");
                CustomXmlPart manifestPart = FindOwnedCustomXmlPart(
                    workbookPart,
                    current.ManifestItemId,
                    CustomXmlNamespace,
                    "reviewManifest");

                string digest = WriteDecisionStatePart(
                    decisionPart,
                    next,
                    current.DecisionItemId,
                    current.ManifestItemId,
                    current.StateSha256,
                    current.ProjectionSha256);
                RewriteManifestDecisionDigest(
                    manifestPart,
                    current.DecisionSha256,
                    digest.ToUpperInvariant());
                return digest;
            }
        }

        private static CustomXmlPart FindOwnedCustomXmlPart(
            WorkbookPart workbookPart,
            string expectedItemId,
            string expectedNamespace,
            string expectedRootName)
        {
            if (!Guid.TryParse(expectedItemId, out Guid expectedPersistentItemId))
                throw new InvalidDataException("A validated Custom XML item ID became invalid.");

            List<CustomXmlPart> customParts = workbookPart.CustomXmlParts
                .Take(MaxCustomXmlPartCount + 1)
                .ToList();
            if (customParts.Count > MaxCustomXmlPartCount)
                throw new InvalidDataException(
                    "The copied workbook contains too many Custom XML parts to revise safely.");

            CustomXmlPart found = null;
            foreach (CustomXmlPart part in customParts)
            {
                CustomXmlRoot root = ReadCustomXmlRoot(part);
                if (!string.Equals(root.NamespaceName, expectedNamespace, StringComparison.Ordinal) ||
                    !string.Equals(root.LocalName, expectedRootName, StringComparison.Ordinal))
                    continue;
                Guid persistentItemId = ReadAndValidatePersistentItemId(part, expectedNamespace);
                if (persistentItemId != expectedPersistentItemId) continue;
                if (found != null)
                    throw new InvalidDataException(
                        "The copied workbook repeats a validated Custom XML item ID.");
                found = part;
            }

            if (found == null)
                throw new InvalidDataException(
                    "The copied workbook is missing a validated Custom XML part: " + expectedItemId);
            return found;
        }

        private static void RewriteManifestDecisionDigest(
            CustomXmlPart manifestPart,
            string expectedOldDigest,
            string newDigest)
        {
            RequireRevisionSha256(expectedOldDigest, nameof(expectedOldDigest));
            RequireRevisionSha256(newDigest, nameof(newDigest));
            string uppercaseNewDigest = newDigest.ToUpperInvariant();
            if (!string.Equals(newDigest, uppercaseNewDigest, StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The replacement manifest decision digest must be uppercase hexadecimal.");

            XDocument manifest = LoadCustomXml(manifestPart, MaxManifestCharacters);
            XElement root = manifest.Root;
            if (root == null || root.Name != StateNamespace + "reviewManifest")
                throw new InvalidDataException("The copied comparison manifest has an invalid root element.");
            List<XAttribute> digestAttributes = root.Attributes("decisionSha256").ToList();
            if (digestAttributes.Count != 1 ||
                !FixedTimeEquals(digestAttributes[0].Value, expectedOldDigest))
                throw new InvalidDataException(
                    "The copied comparison manifest does not contain exactly its validated decision digest.");

            byte[] bytes;
            using (Stream input = manifestPart.GetStream(FileMode.Open, FileAccess.Read))
            using (var bounded = new BoundedReadStream(input, checked(MaxManifestCharacters * 4L)))
            using (var memory = new MemoryStream())
            {
                bounded.CopyTo(memory);
                bytes = memory.ToArray();
            }
            try
            {
                // Decode only to reject malformed UTF-8. Mutation below remains byte-for-byte and never
                // normalizes the XML declaration, whitespace, timestamp, attribute order or namespace syntax.
                new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (DecoderFallbackException exception)
            {
                throw new InvalidDataException(
                    "The copied comparison manifest is not strict UTF-8.",
                    exception);
            }

            byte[] original = (byte[])bytes.Clone();
            byte[] attributeName = Encoding.ASCII.GetBytes("decisionSha256");
            int valueOffset = FindUniqueRootAttributeValue(bytes, attributeName);

            string rawOldDigest = Encoding.ASCII.GetString(bytes, valueOffset, 64);
            RequireRevisionSha256(rawOldDigest, "manifest decisionSha256");
            if (!FixedTimeEquals(rawOldDigest, expectedOldDigest) ||
                !string.Equals(rawOldDigest, digestAttributes[0].Value, StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The raw manifest decision digest differs from its validated XML value.");

            byte[] replacement = Encoding.ASCII.GetBytes(uppercaseNewDigest);
            Buffer.BlockCopy(replacement, 0, bytes, valueOffset, replacement.Length);
            if (bytes.Length != original.Length)
                throw new InvalidDataException("The manifest digest rewrite changed the part length.");
            for (int index = 0; index < bytes.Length; index++)
            {
                if (index >= valueOffset && index < valueOffset + 64) continue;
                if (bytes[index] != original[index])
                    throw new InvalidDataException(
                        "The manifest digest rewrite changed bytes outside the decision digest.");
            }

            using (Stream output = manifestPart.GetStream(FileMode.Create, FileAccess.Write))
                output.Write(bytes, 0, bytes.Length);
        }

        private static int FindUniqueRootAttributeValue(byte[] bytes, byte[] attributeName)
        {
            int cursor = 0;
            if (StartsWith(bytes, cursor, 0xEF, 0xBB, 0xBF)) cursor += 3;
            while (true)
            {
                SkipXmlWhitespace(bytes, ref cursor);
                if (StartsWith(bytes, cursor, (byte)'<', (byte)'?'))
                {
                    cursor = FindMarkupEnd(bytes, cursor + 2, (byte)'?', (byte)'>');
                    continue;
                }
                if (StartsWith(bytes, cursor, (byte)'<', (byte)'!', (byte)'-', (byte)'-'))
                {
                    cursor = FindMarkupEnd(bytes, cursor + 4, (byte)'-', (byte)'-', (byte)'>');
                    continue;
                }
                break;
            }

            if (cursor >= bytes.Length || bytes[cursor++] != (byte)'<' ||
                cursor >= bytes.Length || bytes[cursor] == (byte)'/' ||
                bytes[cursor] == (byte)'!' || bytes[cursor] == (byte)'?')
                throw new InvalidDataException(
                    "The comparison manifest raw root start tag could not be located safely.");

            int rootNameStart = cursor;
            while (cursor < bytes.Length && !IsAttributeNameTerminator(bytes[cursor])) cursor++;
            if (cursor == rootNameStart)
                throw new InvalidDataException("The comparison manifest raw root name is empty.");

            int foundValueOffset = -1;
            int syntacticMatchCount = 0;
            while (cursor < bytes.Length)
            {
                SkipXmlWhitespace(bytes, ref cursor);
                if (cursor >= bytes.Length)
                    break;
                if (bytes[cursor] == (byte)'>')
                {
                    cursor++;
                    break;
                }
                if (bytes[cursor] == (byte)'/' &&
                    cursor + 1 < bytes.Length && bytes[cursor + 1] == (byte)'>')
                {
                    cursor += 2;
                    break;
                }

                int nameOffset = cursor;
                while (cursor < bytes.Length && !IsAttributeNameTerminator(bytes[cursor])) cursor++;
                int nameLength = cursor - nameOffset;
                if (nameLength == 0)
                    throw new InvalidDataException(
                        "The comparison manifest root has unsupported raw attribute syntax.");
                bool isDecisionDigest = ByteRangeEquals(
                    bytes,
                    nameOffset,
                    nameLength,
                    attributeName);

                SkipXmlWhitespace(bytes, ref cursor);
                if (cursor >= bytes.Length || bytes[cursor++] != (byte)'=')
                    throw new InvalidDataException(
                        "The comparison manifest root attribute has no raw equals delimiter.");
                SkipXmlWhitespace(bytes, ref cursor);
                if (cursor >= bytes.Length ||
                    (bytes[cursor] != (byte)'\"' && bytes[cursor] != (byte)'\''))
                    throw new InvalidDataException(
                        "The comparison manifest root attribute has no supported raw quote.");
                byte quote = bytes[cursor++];
                int valueOffset = cursor;
                while (cursor < bytes.Length && bytes[cursor] != quote) cursor++;
                if (cursor >= bytes.Length)
                    throw new InvalidDataException(
                        "The comparison manifest root attribute value is unterminated.");

                if (isDecisionDigest)
                {
                    syntacticMatchCount++;
                    if (syntacticMatchCount > 1)
                        throw new InvalidDataException(
                            "The comparison manifest repeats the root decision digest attribute.");
                    if (cursor - valueOffset != 64)
                        throw new InvalidDataException(
                            "The raw manifest decision digest is not a fixed-length SHA-256 value.");
                    foundValueOffset = valueOffset;
                }
                cursor++;
            }

            if (syntacticMatchCount != 1 || foundValueOffset < 0)
                throw new InvalidDataException(
                    "The comparison manifest root has no unique raw decision digest attribute.");
            return foundValueOffset;
        }

        private static bool IsAttributeNameTerminator(byte value) =>
            value == (byte)' ' || value == (byte)'\t' || value == (byte)'\r' ||
            value == (byte)'\n' || value == (byte)'=' || value == (byte)'>' ||
            value == (byte)'/';

        private static bool ByteRangeEquals(
            byte[] bytes,
            int offset,
            int length,
            byte[] expected)
        {
            if (length != expected.Length || offset < 0 || offset > bytes.Length - length)
                return false;
            for (int index = 0; index < length; index++)
            {
                if (bytes[offset + index] != expected[index]) return false;
            }
            return true;
        }

        private static bool StartsWith(byte[] bytes, int offset, params byte[] expected)
        {
            if (offset < 0 || offset > bytes.Length - expected.Length) return false;
            for (int index = 0; index < expected.Length; index++)
            {
                if (bytes[offset + index] != expected[index]) return false;
            }
            return true;
        }

        private static int FindMarkupEnd(byte[] bytes, int offset, params byte[] terminator)
        {
            for (int cursor = offset; cursor <= bytes.Length - terminator.Length; cursor++)
            {
                if (!StartsWith(bytes, cursor, terminator)) continue;
                return cursor + terminator.Length;
            }
            throw new InvalidDataException(
                "The comparison manifest contains unterminated prolog markup before its root.");
        }

        private static void SkipXmlWhitespace(byte[] bytes, ref int cursor)
        {
            while (cursor < bytes.Length &&
                   (bytes[cursor] == (byte)' ' || bytes[cursor] == (byte)'\t' ||
                    bytes[cursor] == (byte)'\r' || bytes[cursor] == (byte)'\n'))
                cursor++;
        }

        private static void RequireRevisionSha256(string value, string name)
        {
            if (value == null || value.Length != 64 || value.Any(character => !IsHex(character)))
                throw new InvalidDataException("An invalid SHA-256 value was supplied for " + name + ".");
        }

        private static void VerifyDecisionRevision(
            ReviewWorkbookReadResult source,
            ReviewDecisionGeneration expectedGeneration,
            ReviewWorkbookReadResult revised,
            string expectedDecisionSha256)
        {
            bool sourceIdentityPreserved =
                string.Equals(revised.ProjectId, source.ProjectId, StringComparison.Ordinal) &&
                string.Equals(revised.StateJson, source.StateJson, StringComparison.Ordinal) &&
                string.Equals(revised.ManifestItemId, source.ManifestItemId, StringComparison.Ordinal) &&
                FixedTimeEquals(revised.StateSha256, source.StateSha256) &&
                FixedTimeEquals(revised.ProjectionSha256, source.ProjectionSha256) &&
                revised.ChunkItemIds.SequenceEqual(source.ChunkItemIds, StringComparer.Ordinal) &&
                string.Equals(revised.DecisionItemId, source.DecisionItemId, StringComparison.Ordinal);
            if (!sourceIdentityPreserved)
                throw new InvalidDataException(
                    "The decision revision altered immutable review state, projection, chunk or item identities.");
            if (!FixedTimeEquals(revised.DecisionSha256, expectedDecisionSha256))
                throw new InvalidDataException(
                    "The revised decision part does not reproduce its manifest-bound digest.");
            if (!DecisionGenerationMatches(revised.DecisionGeneration, expectedGeneration))
                throw new InvalidDataException(
                    "The closed review workbook did not reproduce the requested decision generation.");
        }

        private static bool DecisionGenerationMatches(
            ReviewDecisionGeneration actual,
            ReviewDecisionGeneration expected)
        {
            if (actual == null || expected == null ||
                !string.Equals(actual.ProjectId, expected.ProjectId, StringComparison.Ordinal) ||
                !string.Equals(actual.CatalogId, expected.CatalogId, StringComparison.Ordinal) ||
                !string.Equals(actual.BaseGenerationId, expected.BaseGenerationId, StringComparison.Ordinal) ||
                !string.Equals(actual.MappingGenerationId, expected.MappingGenerationId, StringComparison.Ordinal) ||
                !string.Equals(actual.GenerationId, expected.GenerationId, StringComparison.Ordinal) ||
                !string.Equals(actual.PriorGenerationId, expected.PriorGenerationId, StringComparison.Ordinal) ||
                !actual.Events.Select(item => item.EventId)
                    .SequenceEqual(expected.Events.Select(item => item.EventId), StringComparer.Ordinal) ||
                !actual.Events.Select(item => item.RequestId)
                    .SequenceEqual(expected.Events.Select(item => item.RequestId), StringComparer.Ordinal) ||
                actual.SubjectStates.Count != expected.SubjectStates.Count)
                return false;

            foreach (KeyValuePair<string, ReviewDecisionState> pair in expected.SubjectStates)
            {
                if (!actual.SubjectStates.TryGetValue(pair.Key, out ReviewDecisionState state) ||
                    state.Status != pair.Value.Status ||
                    !string.Equals(
                        state.SelectedObservationId,
                        pair.Value.SelectedObservationId,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        state.EffectiveEventId,
                        pair.Value.EffectiveEventId,
                        StringComparison.Ordinal))
                    return false;
            }
            return true;
        }
    }
}
