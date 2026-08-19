using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    /// <summary>
    /// Strict canonical UTF-8 XML codec for portable recovery records. Decode accepts only the one byte-level
    /// canonical representation emitted by <see cref="Serialize"/>; alternative encodings, whitespace,
    /// declarations, attribute orders, aliases, extensions, and noncanonical scalars fail closed.
    /// </summary>
    public static class MergeRecoveryJournalCodec
    {
        private const string RootName = "mergeRecoveryJournal";
        private const string RecordName = "record";
        private const string BindingName = "binding";
        private const string ValidationName = "validation";
        private const string PublicationName = "publication";
        private const string AuditName = "audit";
        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);
        private static readonly XNamespace JournalNamespace =
            MergeRecoveryJournalContract.XmlNamespace;

        public static byte[] Serialize(IEnumerable<MergeRecoveryJournalRecord> records)
        {
            MergeRecoveryJournalReplay replay = MergeRecoveryJournal.Replay(records);
            foreach (MergeRecoveryJournalRecord record in replay.Records)
            {
                byte[] fragment = SerializeRecordFragment(record);
                if (fragment.Length > MergeRecoveryJournalContract.MaxRecordBytes)
                    throw CodecFailure("A canonical recovery checkpoint exceeds 64 KiB.");
            }

            using (var output = new MemoryStream())
            {
                using (XmlWriter writer = XmlWriter.Create(output, WriterSettings(false)))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement(RootName, MergeRecoveryJournalContract.XmlNamespace);
                    writer.WriteAttributeString(
                        "xmlns",
                        null,
                        null,
                        MergeRecoveryJournalContract.XmlNamespace);
                    writer.WriteAttributeString(
                        "contractVersion",
                        MergeRecoveryJournalContract.Version.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString(
                        "recordCount",
                        replay.Records.Count.ToString(CultureInfo.InvariantCulture));
                    foreach (MergeRecoveryJournalRecord record in replay.Records)
                        WriteRecord(writer, record);
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
                byte[] bytes = output.ToArray();
                if (bytes.Length > MergeRecoveryJournalContract.MaxJournalBytes)
                    throw CodecFailure("A canonical recovery journal exceeds its bounded document limit.");
                return bytes;
            }
        }

        public static IReadOnlyList<MergeRecoveryJournalRecord> Deserialize(byte[] bytes) =>
            DeserializeAndReplay(bytes).Records;

        public static MergeRecoveryJournalReplay DeserializeAndReplay(byte[] bytes)
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (bytes.Length == 0)
                throw CodecFailure("A recovery journal payload is empty.");
            if (bytes.Length > MergeRecoveryJournalContract.MaxJournalBytes)
                throw CodecFailure("A recovery journal payload exceeds its bounded document limit.");
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                throw CodecFailure("A canonical recovery journal cannot contain a UTF-8 byte-order mark.");

            try
            {
                // Decode once with a throwing UTF-8 implementation before XML examines its own declaration. This
                // prevents a valid alternate encoding declaration from escaping the canonical UTF-8 boundary.
                StrictUtf8.GetString(bytes);
            }
            catch (DecoderFallbackException exception)
            {
                throw CodecFailure("A recovery journal is not well-formed UTF-8.", exception);
            }

            try
            {
                XDocument document;
                using (var input = new MemoryStream(bytes, writable: false))
                using (XmlReader reader = XmlReader.Create(input, ReaderSettings()))
                    document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);

                XElement root = document.Root;
                if (root == null || root.Name != JournalNamespace + RootName)
                    throw CodecFailure("The recovery journal root or namespace is invalid.");
                EnsureDocumentContainsOnlyRoot(document, root);
                RequireRootAttributes(root);

                int contractVersion = ParseCanonicalInt(
                    RequiredAttribute(root, "contractVersion"),
                    "contractVersion");
                if (contractVersion != MergeRecoveryJournalContract.Version)
                    throw CodecFailure("The recovery journal contract version is not supported.");
                int declaredCount = ParseCanonicalInt(
                    RequiredAttribute(root, "recordCount"),
                    "recordCount");
                if (declaredCount <= 0 || declaredCount > MergeRecoveryJournalContract.MaxRecords)
                    throw CodecFailure("The recovery journal record count is outside its safe bound.");

                var parsed = new List<MergeRecoveryJournalRecord>(declaredCount);
                foreach (XNode node in root.Nodes())
                {
                    XElement element = node as XElement;
                    if (element == null || element.Name != JournalNamespace + RecordName)
                        throw CodecFailure(
                            "The recovery journal may contain only canonical record elements without whitespace or comments.");
                    if (parsed.Count == MergeRecoveryJournalContract.MaxRecords)
                        throw CodecFailure("The recovery journal contains more than four records.");
                    parsed.Add(ParseRecord(element));
                }
                if (parsed.Count != declaredCount)
                    throw CodecFailure("The recovery journal record count does not match its envelope.");

                MergeRecoveryJournalReplay replay = MergeRecoveryJournal.Replay(parsed);
                byte[] canonical = Serialize(replay.Records);
                if (!ByteArraysEqual(bytes, canonical))
                    throw CodecFailure(
                        "The recovery journal is valid XML but is not the exact canonical UTF-8 representation.");
                return replay;
            }
            catch (MergeRecoveryJournalException)
            {
                throw;
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is XmlException ||
                exception is ArgumentException ||
                exception is FormatException ||
                exception is OverflowException)
            {
                throw CodecFailure("The recovery journal payload is malformed.", exception);
            }
        }

        private static MergeRecoveryJournalRecord ParseRecord(XElement element)
        {
            RequireAttributes(
                element,
                "contractVersion",
                "sequence",
                "state",
                "recordId",
                "priorRecordId");
            int contractVersion = ParseCanonicalInt(
                RequiredAttribute(element, "contractVersion"),
                "contractVersion");
            int sequence = ParseCanonicalInt(RequiredAttribute(element, "sequence"), "sequence");
            MergeRecoveryJournalState state = ParseState(RequiredAttribute(element, "state"));
            string recordId = RequiredAttribute(element, "recordId");
            string priorRecordId = RequiredAttribute(element, "priorRecordId");

            List<XElement> children = DirectElementChildren(element);
            int position = 0;
            XElement bindingElement = RequiredChild(children, ref position, BindingName);
            MergeRecoveryPreparedBinding binding = ParseBinding(bindingElement);

            MergeRecoveryValidationEvidence validation = null;
            string outputRevisionId = "";
            string auditCommitId = "";
            if (state >= MergeRecoveryJournalState.Validated)
            {
                XElement validationElement = RequiredChild(children, ref position, ValidationName);
                validation = ParseValidation(
                    validationElement,
                    out outputRevisionId,
                    out auditCommitId);
            }

            MergePublicationEvidenceKind? publicationEvidenceKind = null;
            DateTimeOffset? observedUtc = null;
            if (state >= MergeRecoveryJournalState.OutputPublished)
            {
                XElement publicationElement = RequiredChild(children, ref position, PublicationName);
                ParsePublication(
                    publicationElement,
                    out MergePublicationEvidenceKind parsedKind,
                    out DateTimeOffset parsedObservedUtc);
                publicationEvidenceKind = parsedKind;
                observedUtc = parsedObservedUtc;
            }

            MergeRecoveryAuditReceipt auditReceipt = null;
            if (state == MergeRecoveryJournalState.AuditCommitted)
            {
                XElement auditElement = RequiredChild(children, ref position, AuditName);
                auditReceipt = ParseAudit(auditElement);
            }
            if (position != children.Count)
                throw CodecFailure("A recovery checkpoint contains an unknown or out-of-order state block.");

            return MergeRecoveryJournal.CreateParsedRecord(
                contractVersion,
                sequence,
                state,
                recordId,
                priorRecordId,
                binding,
                validation,
                outputRevisionId,
                auditCommitId,
                publicationEvidenceKind,
                observedUtc,
                auditReceipt);
        }

        private static MergeRecoveryPreparedBinding ParseBinding(XElement element)
        {
            RequireEmpty(element);
            RequireAttributes(
                element,
                "projectId",
                "mergeId",
                "sourceMasterRevisionId",
                "sourceProjectStateRoot",
                "baseByteLength",
                "baseSha256",
                "evidenceGenerationId",
                "mappingGenerationId",
                "decisionGenerationId",
                "decisionSetHash",
                "mergePlanHash",
                "destinationLocatorId",
                "destinationLeaf");
            return new MergeRecoveryPreparedBinding(
                RequiredAttribute(element, "projectId"),
                RequiredAttribute(element, "mergeId"),
                RequiredAttribute(element, "sourceMasterRevisionId"),
                RequiredAttribute(element, "sourceProjectStateRoot"),
                ParseCanonicalLong(RequiredAttribute(element, "baseByteLength"), "baseByteLength"),
                RequiredAttribute(element, "baseSha256"),
                RequiredAttribute(element, "evidenceGenerationId"),
                RequiredAttribute(element, "mappingGenerationId"),
                RequiredAttribute(element, "decisionGenerationId"),
                RequiredAttribute(element, "decisionSetHash"),
                RequiredAttribute(element, "mergePlanHash"),
                RequiredAttribute(element, "destinationLocatorId"),
                RequiredAttribute(element, "destinationLeaf"));
        }

        private static MergeRecoveryValidationEvidence ParseValidation(
            XElement element,
            out string outputRevisionId,
            out string auditCommitId)
        {
            RequireEmpty(element);
            RequireAttributes(
                element,
                "stageLeaf",
                "validationContractVersion",
                "validationReceiptHash",
                "expectedOutputLength",
                "expectedOutputSha256",
                "outputRevisionId",
                "auditCommitId");
            outputRevisionId = RequiredAttribute(element, "outputRevisionId");
            auditCommitId = RequiredAttribute(element, "auditCommitId");
            return new MergeRecoveryValidationEvidence(
                RequiredAttribute(element, "stageLeaf"),
                RequiredAttribute(element, "validationContractVersion"),
                RequiredAttribute(element, "validationReceiptHash"),
                ParseCanonicalLong(
                    RequiredAttribute(element, "expectedOutputLength"),
                    "expectedOutputLength"),
                RequiredAttribute(element, "expectedOutputSha256"));
        }

        private static void ParsePublication(
            XElement element,
            out MergePublicationEvidenceKind evidenceKind,
            out DateTimeOffset observedUtc)
        {
            RequireEmpty(element);
            RequireAttributes(element, "evidenceKind", "observedUtc");
            string kind = RequiredAttribute(element, "evidenceKind");
            if (kind == "DirectPublisher")
                evidenceKind = MergePublicationEvidenceKind.DirectPublisher;
            else if (kind == "RestartReconciled")
                evidenceKind = MergePublicationEvidenceKind.RestartReconciled;
            else
                throw CodecFailure("A publication evidence kind is not supported.");
            observedUtc = MergeRecoveryJournalValidation.ParseCanonicalTimestamp(
                RequiredAttribute(element, "observedUtc"),
                "observedUtc");
        }

        private static MergeRecoveryAuditReceipt ParseAudit(XElement element)
        {
            RequireEmpty(element);
            RequireAttributes(
                element,
                "auditCommitId",
                "outputRevisionId",
                "committedProjectStateRoot",
                "committedReviewPackageSha256",
                "committedUtc");
            return new MergeRecoveryAuditReceipt(
                RequiredAttribute(element, "auditCommitId"),
                RequiredAttribute(element, "outputRevisionId"),
                RequiredAttribute(element, "committedProjectStateRoot"),
                RequiredAttribute(element, "committedReviewPackageSha256"),
                MergeRecoveryJournalValidation.ParseCanonicalTimestamp(
                    RequiredAttribute(element, "committedUtc"),
                    "committedUtc"));
        }

        private static void WriteRecord(XmlWriter writer, MergeRecoveryJournalRecord record)
        {
            writer.WriteStartElement(RecordName, MergeRecoveryJournalContract.XmlNamespace);
            writer.WriteAttributeString(
                "contractVersion",
                record.ContractVersion.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("sequence", record.Sequence.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("state", StateText(record.State));
            writer.WriteAttributeString("recordId", record.RecordId);
            writer.WriteAttributeString("priorRecordId", record.PriorRecordId);
            WriteBinding(writer, record.Binding);
            if (record.Validation != null)
                WriteValidation(
                    writer,
                    record.Validation,
                    record.OutputRevisionId,
                    record.AuditCommitId);
            if (record.PublicationEvidenceKind.HasValue)
                WritePublication(
                    writer,
                    record.PublicationEvidenceKind.Value,
                    record.ObservedUtc.Value);
            if (record.AuditReceipt != null)
                WriteAudit(writer, record.AuditReceipt);
            writer.WriteEndElement();
        }

        private static void WriteBinding(XmlWriter writer, MergeRecoveryPreparedBinding binding)
        {
            writer.WriteStartElement(BindingName, MergeRecoveryJournalContract.XmlNamespace);
            writer.WriteAttributeString("projectId", binding.ProjectId);
            writer.WriteAttributeString("mergeId", binding.MergeId);
            writer.WriteAttributeString("sourceMasterRevisionId", binding.SourceMasterRevisionId);
            writer.WriteAttributeString("sourceProjectStateRoot", binding.SourceProjectStateRoot);
            writer.WriteAttributeString(
                "baseByteLength",
                binding.BaseByteLength.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("baseSha256", binding.BaseSha256);
            writer.WriteAttributeString("evidenceGenerationId", binding.EvidenceGenerationId);
            writer.WriteAttributeString("mappingGenerationId", binding.MappingGenerationId);
            writer.WriteAttributeString("decisionGenerationId", binding.DecisionGenerationId);
            writer.WriteAttributeString("decisionSetHash", binding.DecisionSetHash);
            writer.WriteAttributeString("mergePlanHash", binding.MergePlanHash);
            writer.WriteAttributeString("destinationLocatorId", binding.DestinationLocatorId);
            writer.WriteAttributeString("destinationLeaf", binding.DestinationLeaf);
            writer.WriteEndElement();
        }

        private static void WriteValidation(
            XmlWriter writer,
            MergeRecoveryValidationEvidence validation,
            string outputRevisionId,
            string auditCommitId)
        {
            writer.WriteStartElement(ValidationName, MergeRecoveryJournalContract.XmlNamespace);
            writer.WriteAttributeString("stageLeaf", validation.StageLeaf);
            writer.WriteAttributeString("validationContractVersion", validation.ValidationContractVersion);
            writer.WriteAttributeString("validationReceiptHash", validation.ValidationReceiptHash);
            writer.WriteAttributeString(
                "expectedOutputLength",
                validation.ExpectedOutputLength.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("expectedOutputSha256", validation.ExpectedOutputSha256);
            writer.WriteAttributeString("outputRevisionId", outputRevisionId);
            writer.WriteAttributeString("auditCommitId", auditCommitId);
            writer.WriteEndElement();
        }

        private static void WritePublication(
            XmlWriter writer,
            MergePublicationEvidenceKind kind,
            DateTimeOffset observedUtc)
        {
            writer.WriteStartElement(PublicationName, MergeRecoveryJournalContract.XmlNamespace);
            writer.WriteAttributeString("evidenceKind", PublicationKindText(kind));
            writer.WriteAttributeString(
                "observedUtc",
                MergeRecoveryJournalValidation.TimestampText(observedUtc));
            writer.WriteEndElement();
        }

        private static void WriteAudit(XmlWriter writer, MergeRecoveryAuditReceipt audit)
        {
            writer.WriteStartElement(AuditName, MergeRecoveryJournalContract.XmlNamespace);
            writer.WriteAttributeString("auditCommitId", audit.AuditCommitId);
            writer.WriteAttributeString("outputRevisionId", audit.OutputRevisionId);
            writer.WriteAttributeString("committedProjectStateRoot", audit.CommittedProjectStateRoot);
            writer.WriteAttributeString(
                "committedReviewPackageSha256",
                audit.CommittedReviewPackageSha256);
            writer.WriteAttributeString(
                "committedUtc",
                MergeRecoveryJournalValidation.TimestampText(audit.CommittedUtc));
            writer.WriteEndElement();
        }

        private static byte[] SerializeRecordFragment(MergeRecoveryJournalRecord record)
        {
            using (var output = new MemoryStream())
            {
                using (XmlWriter writer = XmlWriter.Create(output, WriterSettings(true)))
                    WriteRecord(writer, record);
                return output.ToArray();
            }
        }

        private static XmlWriterSettings WriterSettings(bool fragment) => new XmlWriterSettings
        {
            Encoding = StrictUtf8,
            OmitXmlDeclaration = fragment,
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            NamespaceHandling = NamespaceHandling.OmitDuplicates,
            CheckCharacters = true,
            ConformanceLevel = fragment ? ConformanceLevel.Fragment : ConformanceLevel.Document,
            CloseOutput = false
        };

        private static XmlReaderSettings ReaderSettings() => new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreWhitespace = false,
            IgnoreComments = false,
            IgnoreProcessingInstructions = false,
            CheckCharacters = true,
            MaxCharactersInDocument = MergeRecoveryJournalContract.MaxJournalBytes,
            MaxCharactersFromEntities = 0,
            ConformanceLevel = ConformanceLevel.Document,
            CloseInput = false
        };

        private static void RequireRootAttributes(XElement root)
        {
            List<XAttribute> attributes = root.Attributes().ToList();
            if (attributes.Count != 3 ||
                !attributes[0].IsNamespaceDeclaration ||
                attributes[0].Name.LocalName != "xmlns" ||
                !string.Equals(
                    attributes[0].Value,
                    MergeRecoveryJournalContract.XmlNamespace,
                    StringComparison.Ordinal) ||
                attributes[1].Name != XName.Get("contractVersion") ||
                attributes[2].Name != XName.Get("recordCount"))
                throw CodecFailure("The recovery journal root attributes are unknown or out of canonical order.");
        }

        private static void RequireAttributes(XElement element, params string[] expectedNames)
        {
            List<XAttribute> attributes = element.Attributes().ToList();
            if (attributes.Count != expectedNames.Length)
                throw CodecFailure("A recovery journal element has missing or unknown attributes.");
            for (int index = 0; index < expectedNames.Length; index++)
            {
                if (attributes[index].IsNamespaceDeclaration ||
                    attributes[index].Name.Namespace != XNamespace.None ||
                    !string.Equals(
                        attributes[index].Name.LocalName,
                        expectedNames[index],
                        StringComparison.Ordinal))
                    throw CodecFailure(
                        "A recovery journal element has unknown or out-of-order attributes.");
            }
        }

        private static string RequiredAttribute(XElement element, string name)
        {
            XAttribute attribute = element.Attribute(name);
            if (attribute == null)
                throw CodecFailure("A required recovery journal attribute is absent: " + name + ".");
            return attribute.Value;
        }

        private static List<XElement> DirectElementChildren(XElement element)
        {
            var result = new List<XElement>();
            foreach (XNode node in element.Nodes())
            {
                XElement child = node as XElement;
                if (child == null)
                    throw CodecFailure(
                        "A recovery checkpoint may contain only canonical state elements without whitespace or comments.");
                result.Add(child);
            }
            return result;
        }

        private static XElement RequiredChild(
            IList<XElement> children,
            ref int position,
            string expectedName)
        {
            if (position >= children.Count ||
                children[position].Name != JournalNamespace + expectedName)
                throw CodecFailure(
                    "A recovery checkpoint is missing or reorders the " + expectedName + " state block.");
            return children[position++];
        }

        private static void RequireEmpty(XElement element)
        {
            if (element.Nodes().Any())
                throw CodecFailure("A scalar recovery journal element cannot contain nested or text content.");
        }

        private static void EnsureDocumentContainsOnlyRoot(XDocument document, XElement root)
        {
            int roots = 0;
            foreach (XNode node in document.Nodes())
            {
                if (ReferenceEquals(node, root))
                    roots++;
                else
                    throw CodecFailure(
                        "A canonical recovery journal cannot contain comments, processing instructions, or extra roots.");
            }
            if (roots != 1)
                throw CodecFailure("A recovery journal must contain exactly one root element.");
        }

        private static int ParseCanonicalInt(string text, string field)
        {
            int value;
            if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value) ||
                !string.Equals(value.ToString(CultureInfo.InvariantCulture), text, StringComparison.Ordinal))
                throw CodecFailure("A recovery journal " + field + " value is not a canonical integer.");
            return value;
        }

        private static long ParseCanonicalLong(string text, string field)
        {
            long value;
            if (!long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value) ||
                !string.Equals(value.ToString(CultureInfo.InvariantCulture), text, StringComparison.Ordinal))
                throw CodecFailure("A recovery journal " + field + " value is not a canonical integer.");
            return value;
        }

        private static MergeRecoveryJournalState ParseState(string text)
        {
            switch (text)
            {
                case "Prepared": return MergeRecoveryJournalState.Prepared;
                case "Validated": return MergeRecoveryJournalState.Validated;
                case "OutputPublished": return MergeRecoveryJournalState.OutputPublished;
                case "AuditCommitted": return MergeRecoveryJournalState.AuditCommitted;
                default: throw CodecFailure("A recovery journal state is not supported.");
            }
        }

        private static string StateText(MergeRecoveryJournalState state)
        {
            switch (state)
            {
                case MergeRecoveryJournalState.Prepared: return "Prepared";
                case MergeRecoveryJournalState.Validated: return "Validated";
                case MergeRecoveryJournalState.OutputPublished: return "OutputPublished";
                case MergeRecoveryJournalState.AuditCommitted: return "AuditCommitted";
                default: throw new ArgumentOutOfRangeException(nameof(state));
            }
        }

        private static string PublicationKindText(MergePublicationEvidenceKind kind)
        {
            switch (kind)
            {
                case MergePublicationEvidenceKind.DirectPublisher: return "DirectPublisher";
                case MergePublicationEvidenceKind.RestartReconciled: return "RestartReconciled";
                default: throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private static bool ByteArraysEqual(byte[] left, byte[] right)
        {
            if (left.Length != right.Length) return false;
            int difference = 0;
            for (int index = 0; index < left.Length; index++)
                difference |= left[index] ^ right[index];
            return difference == 0;
        }

        private static MergeRecoveryJournalException CodecFailure(
            string message,
            Exception inner = null)
        {
            // The public replay exception deliberately carries stable code/sequence only. Preserve parser detail
            // in the message without making platform-specific exception text part of the portable contract.
            string detail = inner == null ? message : message + " " + inner.Message;
            return new MergeRecoveryJournalException(MergeRecoveryJournalCodes.Codec, -1, detail);
        }
    }
}
