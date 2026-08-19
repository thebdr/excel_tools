using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.Cryptography;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    public static class MergeOpenXmlMasterPreflightCodes
    {
        public const string InvalidPath = "CMP-MERGE-MASTER-PATH-INVALID";
        public const string Unreadable = "CMP-MERGE-MASTER-UNREADABLE";
        public const string TooLarge = "CMP-MERGE-MASTER-TOO-LARGE";
        public const string ByteIdentityMismatch = "CMP-MERGE-MASTER-BYTES-MISMATCH";
        public const string CoverageBlocked = "CMP-MERGE-MASTER-COVERAGE-BLOCKED";
        public const string TargetMissing = "CMP-MERGE-MASTER-TARGET-MISSING";
        public const string PreconditionMismatch = "CMP-MERGE-MASTER-PRECONDITION-MISMATCH";
        public const string ProofRequirementFailed = "CMP-MERGE-MASTER-PROOF-REQUIREMENT-FAILED";
        public const string LeaseDisposed = "CMP-MERGE-MASTER-LEASE-DISPOSED";
    }

    public sealed class MergeOpenXmlMasterPreflightException : InvalidOperationException
    {
        internal MergeOpenXmlMasterPreflightException(
            string code,
            string detail,
            string correctiveAction,
            string sourcePath = "",
            string targetId = "",
            Exception innerException = null)
            : base(
                MergePlanValidation.RequireIdentity(code, nameof(code)) + ": " +
                (detail ?? "") +
                " Corrective action: " +
                (correctiveAction ?? "Reacquire and revalidate the master workbook."),
                innerException)
        {
            Code = code;
            Detail = detail ?? "";
            CorrectiveAction = correctiveAction ?? "";
            SourcePath = sourcePath ?? "";
            TargetId = targetId ?? "";
        }

        public string Code { get; }
        public string Detail { get; }
        public string CorrectiveAction { get; }
        /// <summary>Machine-local diagnostic context; never copied into a portable proof.</summary>
        public string SourcePath { get; }
        public string TargetId { get; }
    }

    /// <summary>Path-free exact full-content observation of one target in retained master bytes.</summary>
    public sealed class MergeMasterTargetObservation
    {
        internal MergeMasterTargetObservation(
            MergePlanTarget target,
            CanonicalCellState cell)
        {
            Target = target ?? throw new ArgumentNullException(nameof(target));
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            HasFormula = cell.HasFormula;
            Formula = cell.FormulaText ?? "";
            ConstantValue = ToScalar(cell.ConstantValue);
            ObservationHash = StableId.Compute(
                "merge-master-target-observation-v1",
                Target.TargetId,
                HasFormula ? "1" : "0",
                Formula,
                ((int)ConstantValue.Type).ToString(CultureInfo.InvariantCulture),
                ConstantValue.CanonicalText,
                ((int)ConstantValue.DateSystem).ToString(CultureInfo.InvariantCulture));
        }

        public string ObservationHash { get; }
        public MergePlanTarget Target { get; }
        public bool HasFormula { get; }
        public string Formula { get; }
        public CanonicalScalar ConstantValue { get; }

        internal static CanonicalScalar ToScalar(CanonicalValue value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            ObservationExactType type;
            switch (value.Kind)
            {
                case CanonicalValueKind.Blank:
                    return CanonicalScalar.Blank;
                case CanonicalValueKind.Text:
                    type = ObservationExactType.String;
                    break;
                case CanonicalValueKind.Number:
                    type = ObservationExactType.Number;
                    break;
                case CanonicalValueKind.Boolean:
                    type = ObservationExactType.Boolean;
                    break;
                case CanonicalValueKind.Error:
                    type = ObservationExactType.Error;
                    break;
                case CanonicalValueKind.DateSerial:
                    type = ObservationExactType.DateSerial;
                    break;
                case CanonicalValueKind.IsoDateTime:
                    type = ObservationExactType.IsoDateTime;
                    break;
                default:
                    throw new InvalidOperationException(
                        "The retained master contains an unsupported canonical value kind.");
            }
            return new CanonicalScalar(type, value.ExactText, value.DateSystem);
        }
    }

    /// <summary>
    /// Local retained source occurrence plus portable byte/precondition proofs. The lease must remain active until
    /// the future writer has copied the base bytes through CopyMasterBytesTo; reopening SourcePath is not valid.
    /// </summary>
    public sealed class MergeOpenXmlMasterPreflight : IDisposable
    {
        private readonly RetainedOpenXmlMasterLease _lease;

        internal MergeOpenXmlMasterPreflight(
            RetainedOpenXmlMasterLease lease,
            MergePlanValidationProof sourceIdentityProof,
            MergePlanValidationProof targetPreconditionsProof,
            IEnumerable<MergeMasterTargetObservation> targetObservations)
        {
            _lease = lease ?? throw new ArgumentNullException(nameof(lease));
            SourceIdentityProof = sourceIdentityProof ??
                                  throw new ArgumentNullException(nameof(sourceIdentityProof));
            TargetPreconditionsProof = targetPreconditionsProof ??
                                       throw new ArgumentNullException(nameof(targetPreconditionsProof));
            TargetObservations = new ReadOnlyCollection<MergeMasterTargetObservation>(
                new List<MergeMasterTargetObservation>(
                    targetObservations ??
                    throw new ArgumentNullException(nameof(targetObservations))));
        }

        public string SourcePath => _lease.FullPath;
        public string MasterRevisionId => _lease.MasterRevisionId;
        public long ByteLength => _lease.ByteLength;
        public string ByteContentId => _lease.ByteContentId;
        public bool IsLeaseActive => _lease.IsActive;
        public MergePlanValidationProof SourceIdentityProof { get; }
        public MergePlanValidationProof TargetPreconditionsProof { get; }
        public IReadOnlyList<MergeMasterTargetObservation> TargetObservations { get; }

        public void CopyMasterBytesTo(
            Stream destination,
            CancellationToken cancellationToken = default(CancellationToken)) =>
            _lease.CopyOwnedBytesTo(destination, cancellationToken);

        public void Dispose() => _lease.Dispose();
    }

    public static class MergeOpenXmlMasterPreflightVerifier
    {
        public const string SourceIdentityVerifierContract =
            "retained-openxml-master-byte-identity-v1";
        public const string TargetPreconditionsVerifierContract =
            "retained-openxml-master-target-preconditions-v1";

        public static MergeOpenXmlMasterPreflight OpenAndProve(
            string sourcePath,
            MergePlan plan,
            MergeCellWriteSynthesis synthesis,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            MergeValidationReceiptPolicy.ValidateBindings(plan, synthesis);
            cancellationToken.ThrowIfCancellationRequested();
            RetainedOpenXmlMasterLease lease = null;
            try
            {
                lease = RetainedOpenXmlMasterLease.Open(
                    sourcePath,
                    plan.MasterRevisionId,
                    plan.MasterByteContentId,
                    cancellationToken);
                OpenXmlWorkbookSnapshot snapshot = lease.Snapshot;
                if (!snapshot.CanAssertExactCellEquality)
                {
                    string codes = CoverageCodes(snapshot);
                    throw Failure(
                        MergeOpenXmlMasterPreflightCodes.CoverageBlocked,
                        "The retained master has equality-affecting formula/value coverage gaps" +
                        (codes.Length == 0 ? "." : ": " + codes + "."),
                        "Resolve the reported OOXML coverage findings or use the later verified live-Excel path.",
                        lease);
                }

                IReadOnlyDictionary<string, OpenXmlWorksheetSnapshot> sheets =
                    FreezeSheets(snapshot, lease);
                var observations = new List<MergeMasterTargetObservation>(
                    synthesis.Writes.Count);
                foreach (MergeCellWrite write in synthesis.Writes)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!sheets.TryGetValue(
                            write.Target.SheetName,
                            out OpenXmlWorksheetSnapshot sheet) ||
                        sheet.Kind != OpenXmlPhysicalSheetKind.Worksheet ||
                        !sheet.Snapshot.TryGetCell(
                            write.Target.Address,
                            out CanonicalCellState cell))
                    {
                        throw Failure(
                            MergeOpenXmlMasterPreflightCodes.TargetMissing,
                            "The retained master does not contain exact target " +
                            write.Target.SheetName + "!" + write.Target.Address + ".",
                            "Rebuild mapping/evidence or restore the exact analyzed master occurrence.",
                            lease,
                            write.Target.TargetId);
                    }

                    var observation = new MergeMasterTargetObservation(
                        write.Target,
                        cell);
                    ValidatePreconditions(write, observation, lease);
                    ValidateProofRequirements(write, observation, lease);
                    observations.Add(observation);
                }
                observations.Sort((left, right) =>
                    StringComparer.Ordinal.Compare(
                        left.Target.TargetId,
                        right.Target.TargetId));

                string sourceEvidenceHash = StableId.Compute(
                    "retained-openxml-master-byte-evidence-v1",
                    plan.MasterRevisionId,
                    lease.ByteLength.ToString(CultureInfo.InvariantCulture),
                    lease.ByteContentId);
                MergePlanValidationProof sourceProof =
                    MergePlanValidationProof.CreateTrusted(
                        plan,
                        synthesis,
                        MergePlanRequiredValidation.SourceMasterByteIdentity,
                        SourceIdentityVerifierContract,
                        sourceEvidenceHash);

                var targetComponents = new List<string>
                {
                    plan.PlanHash,
                    synthesis.SynthesisHash,
                    lease.ByteContentId,
                    observations.Count.ToString(CultureInfo.InvariantCulture),
                };
                targetComponents.AddRange(
                    observations.Select(item => item.ObservationHash));
                string targetEvidenceHash = StableId.Compute(
                    "retained-openxml-master-target-evidence-v1",
                    targetComponents.ToArray());
                MergePlanValidationProof targetProof =
                    MergePlanValidationProof.CreateTrusted(
                        plan,
                        synthesis,
                        MergePlanRequiredValidation.MasterTargetPreconditions,
                        TargetPreconditionsVerifierContract,
                        targetEvidenceHash);

                var result = new MergeOpenXmlMasterPreflight(
                    lease,
                    sourceProof,
                    targetProof,
                    observations);
                lease = null;
                return result;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (MergeOpenXmlMasterPreflightException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is InvalidDataException ||
                exception is IOException ||
                exception is UnauthorizedAccessException ||
                exception is SecurityException ||
                exception is NotSupportedException)
            {
                string normalized = lease?.FullPath ?? SafeFullPath(sourcePath);
                throw new MergeOpenXmlMasterPreflightException(
                    MergeOpenXmlMasterPreflightCodes.Unreadable,
                    "The master could not be retained and parsed as bounded OOXML: " +
                    exception.Message,
                    "Close active writers, verify the workbook format and permissions, and retry.",
                    normalized,
                    innerException: exception);
            }
            finally
            {
                lease?.Dispose();
            }
        }

        private static IReadOnlyDictionary<string, OpenXmlWorksheetSnapshot> FreezeSheets(
            OpenXmlWorkbookSnapshot snapshot,
            RetainedOpenXmlMasterLease lease)
        {
            var sheets = new Dictionary<string, OpenXmlWorksheetSnapshot>(
                StringComparer.Ordinal);
            var caseInsensitiveNames = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (OpenXmlWorksheetSnapshot sheet in snapshot.Sheets)
            {
                if (sheet == null ||
                    string.IsNullOrWhiteSpace(sheet.Name) ||
                    sheets.ContainsKey(sheet.Name) ||
                    !caseInsensitiveNames.Add(sheet.Name))
                {
                    throw Failure(
                        MergeOpenXmlMasterPreflightCodes.CoverageBlocked,
                        "The retained master has a missing or duplicate worksheet name.",
                        "Repair the workbook structure on an authorized copy and rebuild comparison evidence.",
                        lease);
                }
                sheets.Add(sheet.Name, sheet);
            }
            return new ReadOnlyDictionary<string, OpenXmlWorksheetSnapshot>(sheets);
        }

        private static void ValidatePreconditions(
            MergeCellWrite write,
            MergeMasterTargetObservation observed,
            RetainedOpenXmlMasterLease lease)
        {
            if (write.FormulaMasterPrecondition != null)
            {
                MergePlanCellState expected = write.FormulaMasterPrecondition;
                bool matches = expected.ExactType == ObservationExactType.Formula
                    ? observed.HasFormula &&
                      string.Equals(
                          observed.Formula,
                          expected.Formula,
                          StringComparison.Ordinal)
                    : expected.ExactType == ObservationExactType.Blank &&
                      !observed.HasFormula;
                if (!matches)
                {
                    throw PreconditionMismatch(
                        write,
                        "Formula",
                        lease);
                }
            }

            if (write.ValueMasterPrecondition != null &&
                !ScalarMatches(
                    write.ValueMasterPrecondition,
                    observed.ConstantValue))
            {
                throw PreconditionMismatch(write, "Value", lease);
            }
        }

        private static void ValidateProofRequirements(
            MergeCellWrite write,
            MergeMasterTargetObservation observed,
            RetainedOpenXmlMasterLease lease)
        {
            if (write.ProofRequirements.HasFlag(
                    MergeCellWriteProofRequirement.MasterHasNoFormula) &&
                observed.HasFormula)
            {
                throw Failure(
                    MergeOpenXmlMasterPreflightCodes.ProofRequirementFailed,
                    "Target " + write.Target.SheetName + "!" +
                    write.Target.Address +
                    " still has a formula, so the selected Value-only write would remove an unselected formula.",
                    "Select a compatible Formula change or keep the master formula.",
                    lease,
                    write.Target.TargetId);
            }
            if (write.ProofRequirements.HasFlag(
                    MergeCellWriteProofRequirement.MasterConstantValueIsBlank) &&
                observed.ConstantValue.Type != ObservationExactType.Blank)
            {
                throw Failure(
                    MergeOpenXmlMasterPreflightCodes.ProofRequirementFailed,
                    "Target " + write.Target.SheetName + "!" +
                    write.Target.Address +
                    " has a non-blank master constant, so adding the selected formula would discard an unselected value.",
                    "Select the compatible blank Value change or keep the master constant.",
                    lease,
                    write.Target.TargetId);
            }
        }

        private static bool ScalarMatches(
            MergePlanCellState expected,
            CanonicalScalar observed) =>
            expected != null &&
            expected.Value != null &&
            observed != null &&
            expected.ExactType == observed.Type &&
            expected.Value.Type == observed.Type &&
            expected.Value.DateSystem == observed.DateSystem &&
            string.Equals(
                expected.Value.CanonicalText,
                observed.CanonicalText,
                StringComparison.Ordinal);

        private static MergeOpenXmlMasterPreflightException PreconditionMismatch(
            MergeCellWrite write,
            string aspect,
            RetainedOpenXmlMasterLease lease) =>
            Failure(
                MergeOpenXmlMasterPreflightCodes.PreconditionMismatch,
                aspect + " content at " + write.Target.SheetName + "!" +
                write.Target.Address +
                " no longer equals the analyzed master precondition.",
                "Discard the stale plan, rerun comparison, and review the change again.",
                lease,
                write.Target.TargetId);

        private static string CoverageCodes(OpenXmlWorkbookSnapshot snapshot)
        {
            var codes = new HashSet<string>(StringComparer.Ordinal);
            foreach (OpenXmlCoverageMessage message in snapshot.CoverageMessages)
                if (message != null && message.AffectsExactCellComparison)
                    codes.Add(message.Code);
            foreach (OpenXmlWorksheetSnapshot sheet in snapshot.Sheets)
                foreach (OpenXmlCoverageMessage message in
                         sheet?.CoverageMessages ?? Array.Empty<OpenXmlCoverageMessage>())
                    if (message != null && message.AffectsExactCellComparison)
                        codes.Add(message.Code);
            return string.Join(
                ", ",
                codes.OrderBy(item => item, StringComparer.Ordinal).Take(8));
        }

        private static MergeOpenXmlMasterPreflightException Failure(
            string code,
            string detail,
            string correctiveAction,
            RetainedOpenXmlMasterLease lease,
            string targetId = "") =>
            new MergeOpenXmlMasterPreflightException(
                code,
                detail,
                correctiveAction,
                lease?.FullPath ?? "",
                targetId);

        private static string SafeFullPath(string sourcePath)
        {
            try
            {
                return string.IsNullOrWhiteSpace(sourcePath)
                    ? sourcePath ?? ""
                    : Path.GetFullPath(sourcePath);
            }
            catch
            {
                return sourcePath ?? "";
            }
        }
    }

    internal sealed class RetainedOpenXmlMasterLease : IDisposable
    {
        private readonly object _gate = new object();
        private FileStream _stream;
        private byte[] _ownedBytes;

        private RetainedOpenXmlMasterLease(
            string fullPath,
            string masterRevisionId,
            FileStream stream,
            byte[] ownedBytes,
            string byteContentId,
            OpenXmlWorkbookSnapshot snapshot)
        {
            FullPath = fullPath;
            MasterRevisionId = masterRevisionId;
            _stream = stream;
            _ownedBytes = ownedBytes;
            ByteLength = ownedBytes.LongLength;
            ByteContentId = byteContentId;
            Snapshot = snapshot;
        }

        internal string FullPath { get; }
        internal string MasterRevisionId { get; }
        internal long ByteLength { get; }
        internal string ByteContentId { get; }
        internal OpenXmlWorkbookSnapshot Snapshot { get; }

        internal bool IsActive
        {
            get
            {
                lock (_gate)
                    return _stream != null && _ownedBytes != null;
            }
        }

        internal static RetainedOpenXmlMasterLease Open(
            string sourcePath,
            string masterRevisionId,
            string expectedByteContentId,
            CancellationToken cancellationToken)
        {
            string fullPath;
            try
            {
                if (string.IsNullOrWhiteSpace(sourcePath))
                    throw new ArgumentException("A master source path is required.", nameof(sourcePath));
                fullPath = Path.GetFullPath(sourcePath);
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is NotSupportedException ||
                exception is PathTooLongException ||
                exception is SecurityException)
            {
                throw new MergeOpenXmlMasterPreflightException(
                    MergeOpenXmlMasterPreflightCodes.InvalidPath,
                    "The master source path is invalid: " + exception.Message,
                    "Select the exact analyzed master workbook and retry.",
                    sourcePath ?? "",
                    innerException: exception);
            }
            if (Directory.Exists(fullPath))
            {
                throw new MergeOpenXmlMasterPreflightException(
                    MergeOpenXmlMasterPreflightCodes.InvalidPath,
                    "The selected master source is a directory, not a workbook.",
                    "Select the exact analyzed master workbook.",
                    fullPath);
            }

            FileStream stream = null;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                stream = new FileStream(
                    fullPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    1024 * 1024,
                    FileOptions.SequentialScan);
                long length = stream.Length;
                if (length < 0 ||
                    length > OpenXmlWorkbookSnapshotAdapter.MaxSourcePackageBytes ||
                    length > int.MaxValue)
                {
                    throw new MergeOpenXmlMasterPreflightException(
                        MergeOpenXmlMasterPreflightCodes.TooLarge,
                        "The master length exceeds the bounded OOXML lease limit.",
                        "Use a smaller workbook or the later isolated streaming path.",
                        fullPath);
                }

                var bytes = new byte[(int)length];
                int total = 0;
                using (SHA256 sha = SHA256.Create())
                {
                    while (total < bytes.Length)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        int read = stream.Read(
                            bytes,
                            total,
                            Math.Min(1024 * 1024, bytes.Length - total));
                        if (read == 0)
                        {
                            throw new IOException(
                                "The retained master ended before its opened length.");
                        }
                        sha.TransformBlock(bytes, total, read, bytes, total);
                        total += read;
                    }
                    if (stream.ReadByte() != -1)
                    {
                        throw new IOException(
                            "The retained master grew beyond its opened length.");
                    }
                    sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                    string hash = ToHex(sha.Hash);
                    if (!string.Equals(
                            hash,
                            expectedByteContentId,
                            StringComparison.Ordinal))
                    {
                        throw new MergeOpenXmlMasterPreflightException(
                            MergeOpenXmlMasterPreflightCodes.ByteIdentityMismatch,
                            "The retained master SHA-256 does not match the frozen plan.",
                            "Restore/select the exact analyzed master or rebuild the comparison and decisions.",
                            fullPath);
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    OpenXmlWorkbookSnapshot snapshot =
                        OpenXmlWorkbookSnapshotAdapter.FromBytes(
                            masterRevisionId,
                            bytes,
                            cancellationToken);
                    if (!string.Equals(
                            snapshot.ByteContentId,
                            hash,
                            StringComparison.Ordinal) ||
                        !string.Equals(
                            snapshot.RevisionId,
                            masterRevisionId,
                            StringComparison.Ordinal))
                    {
                        throw new InvalidDataException(
                            "The parsed snapshot identity disagrees with the retained bytes.");
                    }
                    var result = new RetainedOpenXmlMasterLease(
                        fullPath,
                        masterRevisionId,
                        stream,
                        bytes,
                        hash,
                        snapshot);
                    stream = null;
                    return result;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            finally
            {
                stream?.Dispose();
            }
        }

        internal void CopyOwnedBytesTo(
            Stream destination,
            CancellationToken cancellationToken)
        {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite)
                throw new ArgumentException("The destination stream is not writable.", nameof(destination));
            lock (_gate)
            {
                EnsureActive();
                for (int offset = 0; offset < _ownedBytes.Length; offset += 1024 * 1024)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int count = Math.Min(1024 * 1024, _ownedBytes.Length - offset);
                    destination.Write(_ownedBytes, offset, count);
                }
            }
        }

        public void Dispose()
        {
            lock (_gate)
            {
                _stream?.Dispose();
                _stream = null;
                _ownedBytes = null;
            }
        }

        private void EnsureActive()
        {
            if (_stream == null || _ownedBytes == null)
            {
                throw new MergeOpenXmlMasterPreflightException(
                    MergeOpenXmlMasterPreflightCodes.LeaseDisposed,
                    "The retained master lease is no longer active.",
                    "Reopen and revalidate the exact master before copying base bytes.",
                    FullPath);
            }
        }

        private static string ToHex(byte[] bytes)
        {
            var characters = new char[bytes.Length * 2];
            const string hex = "0123456789ABCDEF";
            for (int index = 0; index < bytes.Length; index++)
            {
                characters[index * 2] = hex[bytes[index] >> 4];
                characters[index * 2 + 1] = hex[bytes[index] & 0x0F];
            }
            return new string(characters);
        }
    }
}
