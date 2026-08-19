using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    public static partial class ReviewWorkbookFile
    {
        private static DecisionPartData ParseDecisionStatePart(
            CustomXmlPart part,
            Guid persistentItemId)
        {
            string rawSha256 = ComputeDecisionPartSha256(part);
            PreflightDecisionXml(part);
            XDocument xml = LoadCustomXml(part, MaxDecisionXmlCharacters);
            XElement root = xml.Root;
            if (root == null || root.Name != DecisionNamespace + "decisionState")
                throw new InvalidDataException("The comparison decision state has an invalid root element.");

            try
            {
                ValidateAllowedAttributes(
                    root,
                    "marker", "schemaVersion", "itemId", "projectId", "manifestItemId",
                    "stateSha256", "projectionSha256", "catalogId", "evidenceGenerationId",
                    "mappingGenerationId", "masterRevisionId", "generationId", "priorGenerationId",
                    "revisionCount", "subjectCount", "observationCount", "groupCount",
                    "membershipCount", "eventCount");
                string decisionSchemaVersion = ValidateDecisionEnvelope(root, persistentItemId);
                RequireOrderedDecisionChildren(
                    root,
                    "evidence",
                    "subjects",
                    "observations",
                    "groups",
                    "generation");

                string projectId = DecisionIdentity(root, "projectId");
                Guid manifestItemId = ParseItemId(
                    DecisionIdentity(root, "manifestItemId"),
                    "manifestItemId");
                string stateSha256 = ParseSha256(root, "stateSha256");
                string projectionSha256 = ParseSha256(root, "projectionSha256");
                string catalogId = DecisionIdentity(root, "catalogId");
                string evidenceGenerationId = DecisionIdentity(root, "evidenceGenerationId");
                string mappingGenerationId = DecisionIdentity(root, "mappingGenerationId");
                string masterRevisionId = DecisionIdentity(root, "masterRevisionId");
                string generationId = DecisionIdentity(root, "generationId");
                string priorGenerationId = DecisionText(root, "priorGenerationId", MaxDecisionIdentityCharacters);

                int revisionCount = DecisionCount(root, "revisionCount", MaxDecisionRevisionCount);
                int subjectCount = DecisionCount(root, "subjectCount", MaxDecisionSubjectCount);
                int observationCount = DecisionCount(root, "observationCount", MaxDecisionObservationCount);
                int groupCount = DecisionCount(root, "groupCount", MaxDecisionGroupCount);
                int membershipCount = DecisionCount(root, "membershipCount", MaxDecisionMembershipCount);
                bool eventSchema = string.Equals(
                    decisionSchemaVersion,
                    DecisionEventStateSchemaVersion,
                    StringComparison.Ordinal);
                int eventCount = DecisionCount(
                    root,
                    "eventCount",
                    eventSchema ? MaxDecisionEventCount : 0);
                if (!eventSchema && (eventCount != 0 || priorGenerationId.Length != 0))
                    throw new InvalidDataException(
                        "Decision sidecar v1 accepts only an empty root generation.");
                if (eventSchema && (eventCount == 0 || priorGenerationId.Length == 0))
                    throw new InvalidDataException(
                        "Decision sidecar v2 requires a non-root generation with at least one event.");
                ValidateDecisionCounts(
                    revisionCount,
                    subjectCount,
                    observationCount,
                    groupCount,
                    membershipCount);

                XElement evidenceElement = root.Element(DecisionNamespace + "evidence");
                ValidateAllowedAttributes(evidenceElement, "effectiveMappingGenerationId");
                string effectiveMappingGenerationId = DecisionIdentity(
                    evidenceElement,
                    "effectiveMappingGenerationId");
                if (!string.Equals(
                        effectiveMappingGenerationId,
                        mappingGenerationId,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "The decision evidence and catalog mapping generations differ.");
                List<ReviewEvidenceRevisionIdentity> revisions = ParseDecisionRevisions(evidenceElement);
                if (revisions.Count != revisionCount)
                    throw new InvalidDataException("The decision revision count is inconsistent.");
                RequireCanonicalOrder(
                    revisions,
                    (left, right) => left.Order.CompareTo(right.Order),
                    "revision");
                var evidence = new ReviewEvidenceGeneration(
                    projectId,
                    effectiveMappingGenerationId,
                    revisions);
                if (!string.Equals(evidence.GenerationId, evidenceGenerationId, StringComparison.Ordinal))
                    throw new InvalidDataException("The decision evidence generation ID is invalid.");

                XElement subjectsElement = root.Element(DecisionNamespace + "subjects");
                ValidateContainer(subjectsElement, "subject");
                List<ReviewDecisionSubject> subjects = subjectsElement
                    .Elements(DecisionNamespace + "subject")
                    .Select(ParseDecisionSubject)
                    .ToList();
                if (subjects.Count != subjectCount)
                    throw new InvalidDataException("The decision subject count is inconsistent.");
                RequireCanonicalOrder(
                    subjects,
                    (left, right) => StringComparer.Ordinal.Compare(left.SubjectId, right.SubjectId),
                    "subject");

                XElement observationsElement = root.Element(DecisionNamespace + "observations");
                ValidateContainer(observationsElement, "observation");
                List<Observation> observations = observationsElement
                    .Elements(DecisionNamespace + "observation")
                    .Select(ParseDecisionObservation)
                    .ToList();
                if (observations.Count != observationCount)
                    throw new InvalidDataException("The decision observation count is inconsistent.");
                RequireCanonicalOrder(
                    observations,
                    (left, right) => StringComparer.Ordinal.Compare(
                        left.ObservationId,
                        right.ObservationId),
                    "observation");

                XElement groupsElement = root.Element(DecisionNamespace + "groups");
                ValidateContainer(groupsElement, "group");
                List<ReviewDecisionGroup> groups = groupsElement
                    .Elements(DecisionNamespace + "group")
                    .Select(ParseDecisionGroup)
                    .ToList();
                if (groups.Count != groupCount ||
                    groups.Sum(item => item.SubjectIds.Count) != membershipCount)
                    throw new InvalidDataException("The decision group counts are inconsistent.");
                RequireCanonicalOrder(
                    groups,
                    (left, right) => StringComparer.Ordinal.Compare(left.GroupId, right.GroupId),
                    "group");

                ValidateCompleteObservationMatrix(subjects, revisions, observations);
                var catalog = new ReviewDecisionCatalog(
                    evidence,
                    masterRevisionId,
                    subjects,
                    observations,
                    groups);
                if (!string.Equals(catalog.CatalogId, catalogId, StringComparison.Ordinal))
                    throw new InvalidDataException("The decision catalog ID is invalid.");

                XElement generationElement = root.Element(DecisionNamespace + "generation");
                ReviewDecisionGeneration generation = ParseDecisionGeneration(
                    generationElement,
                    catalog,
                    decisionSchemaVersion,
                    generationId,
                    priorGenerationId,
                    eventCount);
                return new DecisionPartData(
                    persistentItemId,
                    manifestItemId,
                    projectId,
                    stateSha256,
                    projectionSha256,
                    rawSha256,
                    generation);
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is InvalidOperationException ||
                exception is OverflowException)
            {
                throw new InvalidDataException(
                    "The comparison decision state violates its immutable model invariants.",
                    exception);
            }
        }

        private static string ComputeDecisionPartSha256(CustomXmlPart part)
        {
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                if (stream.CanSeek && stream.Length > MaxDecisionXmlBytes)
                    throw new InvalidDataException(
                        "The comparison decision state exceeds the safe source-size limit.");
                using (SHA256 sha256 = SHA256.Create())
                    return BytesToHex(sha256.ComputeHash(new BoundedReadStream(
                        stream,
                        MaxDecisionXmlBytes)));
            }
        }

        private static void PreflightDecisionXml(CustomXmlPart part)
        {
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = CreateBoundedXmlReader(stream, MaxDecisionXmlCharacters))
            {
                int elementCount = 0;
                bool sawRoot = false;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (++elementCount > MaxDecisionXmlElementCount)
                        throw new InvalidDataException(
                            "The comparison decision state exceeds the safe XML element-count limit.");
                    if (reader.Depth > MaxDecisionXmlDepth)
                        throw new InvalidDataException(
                            "The comparison decision state exceeds the safe XML depth limit.");
                    if (reader.AttributeCount > MaxDecisionAttributesPerElement)
                        throw new InvalidDataException(
                            "A comparison decision element has too many attributes.");
                    if (sawRoot) continue;
                    sawRoot = true;
                    if (reader.LocalName != "decisionState" ||
                        reader.NamespaceURI != DecisionCustomXmlNamespace)
                        throw new InvalidDataException(
                            "The comparison decision state has an invalid root element.");
                    string schemaVersion = reader.GetAttribute("schemaVersion");
                    bool rootSchema = string.Equals(
                        schemaVersion,
                        SchemaVersion,
                        StringComparison.Ordinal);
                    bool eventSchema = string.Equals(
                        schemaVersion,
                        DecisionEventStateSchemaVersion,
                        StringComparison.Ordinal);
                    if (!rootSchema && !eventSchema)
                        throw new InvalidDataException(
                            "The comparison decision schema version is unsupported.");
                    PreflightDecisionCount(reader, "revisionCount", MaxDecisionRevisionCount);
                    PreflightDecisionCount(reader, "subjectCount", MaxDecisionSubjectCount);
                    PreflightDecisionCount(reader, "observationCount", MaxDecisionObservationCount);
                    PreflightDecisionCount(reader, "groupCount", MaxDecisionGroupCount);
                    PreflightDecisionCount(reader, "membershipCount", MaxDecisionMembershipCount);
                    PreflightDecisionCount(
                        reader,
                        "eventCount",
                        eventSchema ? MaxDecisionEventCount : 0);
                }
                if (!sawRoot)
                    throw new InvalidDataException("The comparison decision state is empty.");
            }
        }

        private static void PreflightDecisionCount(
            XmlReader reader,
            string attributeName,
            int maximum)
        {
            string text = reader.GetAttribute(attributeName);
            if (text == null ||
                !int.TryParse(
                    text,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out int value) ||
                value < 0 || value > maximum)
                throw new InvalidDataException(
                    "A decision count is missing, invalid, or exceeds its safe limit: " +
                    attributeName);
        }

        private static string ValidateDecisionEnvelope(XElement root, Guid persistentItemId)
        {
            if (!string.Equals(RequireAttribute(root, "marker"), DecisionMarker, StringComparison.Ordinal))
                throw new InvalidDataException("The comparison decision marker is invalid.");
            string schemaVersion = RequireAttribute(root, "schemaVersion");
            if (!string.Equals(schemaVersion, SchemaVersion, StringComparison.Ordinal) &&
                !string.Equals(
                    schemaVersion,
                    DecisionEventStateSchemaVersion,
                    StringComparison.Ordinal))
                throw new InvalidDataException("The comparison decision schema version is unsupported.");
            if (ParseItemId(RequireAttribute(root, "itemId"), "itemId") != persistentItemId)
                throw new InvalidDataException(
                    "The decision Custom XML content and properties item IDs differ.");
            return schemaVersion;
        }

        private static ReviewDecisionGeneration ParseDecisionGeneration(
            XElement element,
            ReviewDecisionCatalog catalog,
            string schemaVersion,
            string expectedGenerationId,
            string expectedPriorGenerationId,
            int expectedEventCount)
        {
            ValidateAllowedAttributes(
                element,
                "generationId",
                "priorGenerationId",
                "eventCount");
            if (!string.Equals(
                    DecisionIdentity(element, "generationId"),
                    expectedGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    DecisionText(
                        element,
                        "priorGenerationId",
                        MaxDecisionIdentityCharacters),
                    expectedPriorGenerationId,
                    StringComparison.Ordinal) ||
                DecisionCount(element, "eventCount", MaxDecisionEventCount) != expectedEventCount)
                throw new InvalidDataException("The decision generation envelope is inconsistent.");

            bool eventSchema = string.Equals(
                schemaVersion,
                DecisionEventStateSchemaVersion,
                StringComparison.Ordinal);
            if (!eventSchema)
            {
                EnsureEmptyElement(element);
                ReviewDecisionGeneration root = ReviewDecisionEngine.Create(catalog);
                if (!string.Equals(
                        root.GenerationId,
                        expectedGenerationId,
                        StringComparison.Ordinal))
                    throw new InvalidDataException("The decision root generation ID is invalid.");
                return root;
            }

            EnsureOnlyDecisionElements(element, "replayRecord");
            List<XElement> recordElements = element
                .Elements(DecisionNamespace + "replayRecord")
                .ToList();
            if (recordElements.Count != expectedEventCount)
                throw new InvalidDataException("The decision replay-record count is inconsistent.");
            var records = new List<ReviewDecisionReplayRecord>(recordElements.Count);
            for (int sequence = 0; sequence < recordElements.Count; sequence++)
            {
                ReviewDecisionReplayRecord record = ParseDecisionReplayRecord(recordElements[sequence]);
                if (record.Sequence != sequence)
                    throw new InvalidDataException(
                        "Decision replay-record sequences must be contiguous in append order.");
                records.Add(record);
            }

            try
            {
                return ReviewDecisionReplayEngine.Replay(
                    catalog,
                    records,
                    expectedGenerationId,
                    expectedPriorGenerationId);
            }
            catch (ReviewDecisionReplayException exception)
            {
                throw new InvalidDataException(
                    "The comparison decision event ledger failed deterministic replay: " +
                    exception.Code,
                    exception);
            }
        }

        private static ReviewDecisionReplayRecord ParseDecisionReplayRecord(XElement element)
        {
            ValidateAllowedAttributes(
                element,
                "contractVersion",
                "sequence",
                "priorGenerationId",
                "newGenerationId");
            RequireOrderedDecisionChildren(element, "event");
            int contractVersion = DecisionCount(element, "contractVersion", int.MaxValue);
            if (contractVersion != ReviewDecisionEventContractVersions.V1)
                throw new InvalidDataException("The decision event contract version is unsupported.");
            int sequence = DecisionCount(element, "sequence", MaxDecisionEventCount - 1);
            return new ReviewDecisionReplayRecord(
                contractVersion,
                sequence,
                DecisionIdentity(element, "priorGenerationId"),
                DecisionIdentity(element, "newGenerationId"),
                ParseDecisionEvent(element.Element(DecisionNamespace + "event")));
        }

        private static DecisionEvent ParseDecisionEvent(XElement element)
        {
            ValidateAllowedAttributes(
                element,
                "eventId", "requestId", "projectId", "evidenceGenerationId",
                "mappingGenerationId", "priorDecision", "newDecision", "createdUtc",
                "selectionScope", "runId", "subjectId", "groupId",
                "selectedObservationId", "supersededEventId", "affectedSubjectCount",
                "supportingRevisionCount", "supportingContentHashCount",
                "reviewerSourceCount", "locationCount");
            RequireOrderedDecisionChildren(
                element,
                "affectedSubjects",
                "supportingRevisions",
                "supportingContentHashes",
                "reviewer",
                "addInBuild",
                "note",
                "locations");

            int affectedSubjectCount = DecisionCount(
                element,
                "affectedSubjectCount",
                MaxDecisionAffectedSubjectCountPerEvent);
            int supportingRevisionCount = DecisionCount(
                element,
                "supportingRevisionCount",
                MaxDecisionSupportingRevisionCountPerEvent);
            int supportingContentHashCount = DecisionCount(
                element,
                "supportingContentHashCount",
                MaxDecisionSupportingContentHashCountPerEvent);
            int reviewerSourceCount = DecisionCount(
                element,
                "reviewerSourceCount",
                MaxDecisionReviewerSourceCountPerEvent);
            int locationCount = DecisionCount(
                element,
                "locationCount",
                MaxDecisionLocationCountPerEvent);

            IReadOnlyList<string> affectedSubjects = ParseDecisionIdentityCollection(
                element.Element(DecisionNamespace + "affectedSubjects"),
                "subject",
                "subjectId",
                affectedSubjectCount,
                false);
            IReadOnlyList<string> supportingRevisions = ParseDecisionIdentityCollection(
                element.Element(DecisionNamespace + "supportingRevisions"),
                "revision",
                "revisionId",
                supportingRevisionCount,
                false);
            IReadOnlyList<string> supportingContentHashes = ParseDecisionIdentityCollection(
                element.Element(DecisionNamespace + "supportingContentHashes"),
                "content",
                "contentHash",
                supportingContentHashCount,
                true);
            ReviewerSnapshot reviewer = ParseDecisionReviewer(
                element.Element(DecisionNamespace + "reviewer"),
                reviewerSourceCount);
            XElement addInBuildElement = element.Element(DecisionNamespace + "addInBuild");
            ValidateAllowedAttributes(addInBuildElement, "encoding", "byteLength");
            XElement noteElement = element.Element(DecisionNamespace + "note");
            ValidateAllowedAttributes(noteElement, "encoding", "byteLength");
            IReadOnlyList<DecisionLocationContext> locations = ParseDecisionLocations(
                element.Element(DecisionNamespace + "locations"),
                locationCount);

            DateTimeOffset createdUtc = ParseDecisionTimestamp(element, "createdUtc");
            return new DecisionEvent(
                DecisionIdentity(element, "eventId"),
                DecisionIdentity(element, "projectId"),
                DecisionIdentity(element, "evidenceGenerationId"),
                DecisionIdentity(element, "mappingGenerationId"),
                ParseDecisionEnum<ReviewDecisionStatus>(element, "priorDecision"),
                ParseDecisionEnum<ReviewDecisionStatus>(element, "newDecision"),
                createdUtc,
                ParseDecisionEnum<DecisionSelectionScope>(element, "selectionScope"),
                DecisionText(element, "runId", MaxDecisionIdentityCharacters),
                DecisionText(element, "subjectId", MaxDecisionIdentityCharacters),
                DecisionText(element, "groupId", MaxDecisionIdentityCharacters),
                affectedSubjects,
                DecisionText(
                    element,
                    "selectedObservationId",
                    MaxDecisionIdentityCharacters),
                supportingRevisions,
                supportingContentHashes,
                reviewer,
                ParseDecisionEncodedText(addInBuildElement),
                ParseDecisionEncodedText(noteElement),
                DecisionText(
                    element,
                    "supersededEventId",
                    MaxDecisionIdentityCharacters),
                locations,
                DecisionIdentity(element, "requestId"));
        }

        private static IReadOnlyList<string> ParseDecisionIdentityCollection(
            XElement container,
            string childName,
            string attributeName,
            int expectedCount,
            bool sha256)
        {
            ValidateAllowedAttributes(container);
            EnsureOnlyDecisionElements(container, childName);
            List<XElement> elements = container.Elements(DecisionNamespace + childName).ToList();
            if (elements.Count != expectedCount)
                throw new InvalidDataException(
                    "A decision event collection count is inconsistent: " + childName);
            var result = new List<string>(elements.Count);
            foreach (XElement element in elements)
            {
                ValidateAllowedAttributes(element, attributeName);
                EnsureEmptyElement(element);
                result.Add(sha256
                    ? ParseSha256(element, attributeName)
                    : DecisionIdentity(element, attributeName));
            }
            return result;
        }

        private static ReviewerSnapshot ParseDecisionReviewer(
            XElement element,
            int expectedSourceCount)
        {
            ValidateAllowedAttributes(element);
            RequireOrderedDecisionChildren(
                element,
                "confirmedDisplayName",
                "officeDisplayName",
                "identitySources");
            XElement confirmed = element.Element(DecisionNamespace + "confirmedDisplayName");
            XElement office = element.Element(DecisionNamespace + "officeDisplayName");
            ValidateAllowedAttributes(confirmed, "encoding", "byteLength");
            ValidateAllowedAttributes(office, "encoding", "byteLength");
            XElement sourcesElement = element.Element(DecisionNamespace + "identitySources");
            ValidateAllowedAttributes(sourcesElement);
            EnsureOnlyDecisionElements(sourcesElement, "source");
            List<XElement> sourceElements = sourcesElement
                .Elements(DecisionNamespace + "source")
                .ToList();
            if (sourceElements.Count != expectedSourceCount)
                throw new InvalidDataException(
                    "The decision reviewer identity-source count is inconsistent.");
            var sources = new List<string>(sourceElements.Count);
            foreach (XElement source in sourceElements)
            {
                ValidateAllowedAttributes(source, "encoding", "byteLength");
                sources.Add(ParseDecisionEncodedText(source));
            }
            return new ReviewerSnapshot(
                ParseDecisionEncodedText(confirmed),
                ParseDecisionEncodedText(office),
                sources);
        }

        private static IReadOnlyList<DecisionLocationContext> ParseDecisionLocations(
            XElement container,
            int expectedCount)
        {
            ValidateAllowedAttributes(container);
            EnsureOnlyDecisionElements(container, "location");
            List<XElement> elements = container.Elements(DecisionNamespace + "location").ToList();
            if (elements.Count != expectedCount)
                throw new InvalidDataException("The decision event location count is inconsistent.");
            var locations = new List<DecisionLocationContext>(elements.Count);
            foreach (XElement element in elements)
            {
                ValidateAllowedAttributes(
                    element,
                    "role",
                    "revisionId",
                    "physicalRow",
                    "physicalColumn");
                RequireOrderedDecisionChildren(
                    element,
                    "logicalSheetId",
                    "sheetName",
                    "address",
                    "sourceLabel");
                List<XElement> text = DecisionElementChildren(element);
                foreach (XElement value in text)
                    ValidateAllowedAttributes(value, "encoding", "byteLength");
                string sourceLabel = ParseDecisionEncodedText(text[3]);
                ValidatePortableDecisionSourceLabel(sourceLabel);
                locations.Add(new DecisionLocationContext(
                    DecisionIdentity(element, "role"),
                    DecisionIdentity(element, "revisionId"),
                    ParseDecisionEncodedText(text[0]),
                    ParseDecisionEncodedText(text[1]),
                    ParseDecisionEncodedText(text[2]),
                    sourceLabel,
                    DecisionCount(element, "physicalRow", int.MaxValue),
                    DecisionCount(element, "physicalColumn", int.MaxValue)));
            }
            return locations;
        }

        private static DateTimeOffset ParseDecisionTimestamp(XElement element, string attributeName)
        {
            string text = DecisionText(element, attributeName, 64);
            if (!DateTimeOffset.TryParseExact(
                    text,
                    "O",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTimeOffset value) ||
                value.Offset != TimeSpan.Zero ||
                !string.Equals(
                    value.ToString("O", CultureInfo.InvariantCulture),
                    text,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "A decision event timestamp is not canonical UTC: " + attributeName);
            return value;
        }

        private static List<ReviewEvidenceRevisionIdentity> ParseDecisionRevisions(XElement container)
        {
            EnsureOnlyDecisionElements(container, "revision");
            var revisions = new List<ReviewEvidenceRevisionIdentity>();
            foreach (XElement element in container.Elements())
            {
                ValidateAllowedAttributes(element, "revisionId", "order", "byteContentId");
                EnsureEmptyElement(element);
                int order = DecisionCount(element, "order", int.MaxValue);
                revisions.Add(new ReviewEvidenceRevisionIdentity(
                    DecisionIdentity(element, "revisionId"),
                    order,
                    ParseSha256(element, "byteContentId")));
            }
            return revisions;
        }

        private static ReviewDecisionSubject ParseDecisionSubject(XElement element)
        {
            ValidateAllowedAttributes(element, "subjectId", "logicalSubjectKey", "aspect", "policy");
            EnsureEmptyElement(element);
            var subject = new ReviewDecisionSubject(
                DecisionIdentity(element, "subjectId"),
                DecisionIdentity(element, "logicalSubjectKey"),
                ParseDecisionEnum<ComparisonAspect>(element, "aspect"),
                ParseDecisionEnum<ReviewDecisionSubjectPolicy>(element, "policy"));
            if (!subject.HasExplicitMetadata)
                throw new InvalidDataException(
                    "A persisted decision subject must have explicit aspect and policy metadata.");
            return subject;
        }

        private static Observation ParseDecisionObservation(XElement element)
        {
            ValidateAllowedAttributes(
                element,
                "observationId", "subjectId", "revisionId", "exactType", "contentHash",
                "cachedResultTrust", "logicalSheetId", "sourceLabel", "sheetName", "address",
                "physicalRow", "physicalColumn");
            ObservationExactType exactType = ParseDecisionEnum<ObservationExactType>(element, "exactType");
            CachedResultTrust trust = ParseDecisionEnum<CachedResultTrust>(
                element,
                "cachedResultTrust");
            List<XElement> children = DecisionElementChildren(element);
            if (children.Count == 0 ||
                (children[0].Name != DecisionNamespace + "formula" &&
                 children[0].Name != DecisionNamespace + "value"))
                throw new InvalidDataException("A decision observation has no exact payload.");
            bool formulaPayload = children[0].Name == DecisionNamespace + "formula";
            if ((exactType == ObservationExactType.Formula) != formulaPayload)
                throw new InvalidDataException(
                    "A decision observation payload does not match its exact type.");
            if (formulaPayload)
                ValidateAllowedAttributes(children[0], "encoding", "byteLength");
            int index = 1;
            XElement cachedElement = null;
            XElement displayElement = null;
            if (index < children.Count && children[index].Name == DecisionNamespace + "cachedResult")
                cachedElement = children[index++];
            if (index < children.Count && children[index].Name == DecisionNamespace + "displayText")
                displayElement = children[index++];
            if (index != children.Count)
                throw new InvalidDataException("A decision observation contains unknown or reordered content.");
            if (displayElement != null)
                ValidateAllowedAttributes(displayElement, "encoding", "byteLength");

            string formula = formulaPayload ? ParseDecisionEncodedText(children[0]) : null;
            CanonicalScalar value = formulaPayload ? null : ParseDecisionScalar(children[0]);
            CanonicalScalar cachedResult = cachedElement == null
                ? null
                : ParseDecisionScalar(cachedElement);
            string displayText = displayElement == null
                ? ""
                : ParseDecisionEncodedText(displayElement);
            string sourceLabel = DecisionText(
                element,
                "sourceLabel",
                MaxDecisionLabelCharacters);
            ValidatePortableDecisionSourceLabel(sourceLabel);
            return new Observation(
                DecisionIdentity(element, "observationId"),
                DecisionIdentity(element, "subjectId"),
                DecisionIdentity(element, "revisionId"),
                exactType,
                DecisionIdentity(element, "contentHash"),
                formula,
                value,
                cachedResult,
                trust,
                displayText,
                new ObservationProvenance(
                    sourceLabel,
                    DecisionText(element, "sheetName", MaxDecisionLabelCharacters),
                    DecisionText(element, "address", MaxDecisionLabelCharacters),
                    DecisionCount(element, "physicalRow", int.MaxValue),
                    DecisionCount(element, "physicalColumn", int.MaxValue),
                    DecisionText(element, "logicalSheetId", MaxDecisionIdentityCharacters)));
        }

        private static ReviewDecisionGroup ParseDecisionGroup(XElement element)
        {
            ValidateAllowedAttributes(element, "groupId");
            EnsureOnlyDecisionElements(element, "member");
            var members = new List<string>();
            foreach (XElement member in element.Elements())
            {
                ValidateAllowedAttributes(member, "subjectId");
                EnsureEmptyElement(member);
                members.Add(DecisionIdentity(member, "subjectId"));
            }
            for (int index = 1; index < members.Count; index++)
            {
                if (StringComparer.Ordinal.Compare(members[index - 1], members[index]) >= 0)
                    throw new InvalidDataException(
                        "Decision group members are not in canonical ordinal order.");
            }
            return new ReviewDecisionGroup(DecisionIdentity(element, "groupId"), members);
        }

        private static CanonicalScalar ParseDecisionScalar(XElement element)
        {
            ValidateAllowedAttributes(element, "type", "dateSystem", "encoding", "byteLength");
            ObservationExactType type = ParseDecisionEnum<ObservationExactType>(element, "type");
            if (type == ObservationExactType.Formula)
                throw new InvalidDataException("A canonical scalar cannot have formula type.");
            ExcelDateSystem dateSystem = ParseDecisionEnum<ExcelDateSystem>(element, "dateSystem");
            return new CanonicalScalar(type, ParseDecisionEncodedText(element), dateSystem);
        }

        private static string ParseDecisionEncodedText(XElement element)
        {
            if (element.HasElements)
                throw new InvalidDataException("A decision text payload cannot contain child elements.");
            string encoding = RequireAttribute(element, "encoding");
            if (!string.Equals(encoding, "base64-utf8", StringComparison.Ordinal))
                throw new InvalidDataException("The decision text encoding is unsupported.");
            int byteLength = DecisionCount(element, "byteLength", MaxDecisionPayloadBytes);
            string encoded = element.Value;
            int expectedCharacters = checked(((byteLength + 2) / 3) * 4);
            if (encoded.Length != expectedCharacters)
                throw new InvalidDataException("A decision text payload length is inconsistent.");
            byte[] bytes;
            try
            {
                bytes = Convert.FromBase64String(encoded);
            }
            catch (FormatException exception)
            {
                throw new InvalidDataException("A decision text payload is not valid Base64.", exception);
            }
            if (bytes.Length != byteLength)
                throw new InvalidDataException("A decoded decision text payload length is inconsistent.");
            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (DecoderFallbackException exception)
            {
                throw new InvalidDataException("A decision text payload is not valid UTF-8.", exception);
            }
        }

        private static void ValidateCompleteObservationMatrix(
            IReadOnlyList<ReviewDecisionSubject> subjects,
            IReadOnlyList<ReviewEvidenceRevisionIdentity> revisions,
            IReadOnlyList<Observation> observations)
        {
            long expected = checked((long)subjects.Count * revisions.Count);
            if (expected != observations.Count)
                throw new InvalidDataException(
                    "The decision catalog must contain one observation for every subject and evidence revision.");
            var expectedRevisions = new HashSet<string>(
                revisions.Select(item => item.RevisionId),
                StringComparer.Ordinal);
            var pairs = new HashSet<string>(StringComparer.Ordinal);
            foreach (Observation observation in observations)
            {
                if (!expectedRevisions.Contains(observation.RevisionId) ||
                    !pairs.Add(observation.SubjectId + "\u001F" + observation.RevisionId))
                    throw new InvalidDataException(
                        "The decision observation matrix contains an unknown or duplicate revision occurrence.");
            }
        }

        private static void ValidateContainer(XElement container, string childName)
        {
            ValidateAllowedAttributes(container);
            EnsureOnlyDecisionElements(container, childName);
        }

        private static void RequireOrderedDecisionChildren(XElement root, params string[] names)
        {
            List<XElement> children = DecisionElementChildren(root);
            if (children.Count != names.Length)
                throw new InvalidDataException("The decision state has an invalid child-element count.");
            for (int index = 0; index < names.Length; index++)
            {
                if (children[index].Name != DecisionNamespace + names[index])
                    throw new InvalidDataException(
                        "The decision state contains unknown or reordered content.");
            }
        }

        private static void EnsureOnlyDecisionElements(XElement parent, string localName)
        {
            foreach (XElement element in DecisionElementChildren(parent))
            {
                if (element.Name != DecisionNamespace + localName)
                    throw new InvalidDataException("Unknown decision state element: " + element.Name);
            }
        }

        private static List<XElement> DecisionElementChildren(XElement parent)
        {
            var result = new List<XElement>();
            foreach (XNode node in parent.Nodes())
            {
                if (node is XElement element)
                    result.Add(element);
                else if (!(node is XText text) || !string.IsNullOrWhiteSpace(text.Value))
                    throw new InvalidDataException("Unexpected content in the comparison decision state.");
            }
            return result;
        }

        private static string DecisionIdentity(XElement element, string attributeName)
        {
            string value = DecisionText(element, attributeName, MaxDecisionIdentityCharacters);
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidDataException("A decision identity attribute is blank: " + attributeName);
            return value;
        }

        private static string DecisionText(XElement element, string attributeName, int maxCharacters)
        {
            string value = RequireAttribute(element, attributeName);
            if (value.Length > maxCharacters)
                throw new InvalidDataException("A decision attribute exceeds its safe limit: " + attributeName);
            return value;
        }

        private static int DecisionCount(XElement element, string attributeName, int maximum)
        {
            int value = ParseNonNegativeInt(element, attributeName);
            if (value > maximum)
                throw new InvalidDataException("A decision count exceeds its safe limit: " + attributeName);
            return value;
        }

        private static T ParseDecisionEnum<T>(XElement element, string attributeName)
            where T : struct
        {
            int numeric = DecisionCount(element, attributeName, int.MaxValue);
            if (!Enum.IsDefined(typeof(T), numeric))
                throw new InvalidDataException("A decision enum value is unsupported: " + attributeName);
            return (T)Enum.ToObject(typeof(T), numeric);
        }

        private static void RequireCanonicalOrder<T>(
            IReadOnlyList<T> items,
            Comparison<T> comparison,
            string label)
        {
            for (int index = 1; index < items.Count; index++)
            {
                if (comparison(items[index - 1], items[index]) >= 0)
                    throw new InvalidDataException(
                        "Decision " + label + " entries are not in canonical order.");
            }
        }

        private sealed class BoundedReadStream : Stream
        {
            private readonly Stream _inner;
            private readonly long _maximum;
            private long _read;

            public BoundedReadStream(Stream inner, long maximum)
            {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
                _maximum = maximum;
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                int result = _inner.Read(buffer, offset, count);
                _read = checked(_read + result);
                if (_read > _maximum)
                    throw new InvalidDataException(
                        "The comparison decision state exceeds the safe source-size limit.");
                return result;
            }

            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }
            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();
        }

        private sealed class BoundedWriteStream : Stream
        {
            private readonly Stream _inner;
            private readonly long _maximum;
            private long _written;

            public BoundedWriteStream(Stream inner, long maximum)
            {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
                _maximum = maximum;
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                if (count < 0 || _written > _maximum - count)
                    throw new InvalidDataException(
                        "The comparison decision state exceeds the safe source-size limit.");
                _inner.Write(buffer, offset, count);
                _written += count;
            }

            public override void Flush() => _inner.Flush();
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => _written;
            public override long Position
            {
                get => _written;
                set => throw new NotSupportedException();
            }
            public override int Read(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
        }
    }
}
