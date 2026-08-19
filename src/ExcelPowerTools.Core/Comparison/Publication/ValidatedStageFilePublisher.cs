using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace ExcelPowerTools.Core.Comparison.Publication
{
    /// <summary>
    /// Publishes an already-created, already-semantically-validated stage file without replacing a destination.
    /// This primitive is deliberately limited to one existing, local, fixed NTFS directory on Windows.
    /// It inventories the unnamed data stream and hard-link count at protected checkpoints, but does not claim
    /// to exclude a coordinated same-user or privileged process that changes streams, links, or the namespace
    /// between those checkpoints. The retained handles protect against ordinary rename/substitution races.
    /// </summary>
    public static class ValidatedStageFilePublisher
    {
        private const uint GenericRead = 0x80000000;
        private const uint DeleteAccess = 0x00010000;
        private const uint FileListDirectory = 0x00000001;
        private const uint FileReadAttributes = 0x00000080;
        private const uint FileShareRead = 0x00000001;
        private const uint FileShareWrite = 0x00000002;
        private const uint OpenExisting = 3;
        private const uint FileFlagBackupSemantics = 0x02000000;
        private const uint FileFlagOpenReparsePoint = 0x00200000;
        private const uint FileFlagSequentialScan = 0x08000000;
        private const uint FileFlagWriteThrough = 0x80000000;
        private const uint FileAttributeDirectory = 0x00000010;
        private const uint FileAttributeReparsePoint = 0x00000400;
        private const uint DriveFixed = 3;
        private const int ErrorFileNotFound = 2;
        private const int ErrorPathNotFound = 3;
        private const int ErrorAccessDenied = 5;
        private const int ErrorInvalidData = 13;
        private const int ErrorFileExists = 80;
        private const int ErrorAlreadyExists = 183;
        private const int ErrorMoreData = 234;
        private const int FileStreamInfo = 7;
        private const int FileRenameInfo = 3;
        private const int FileCaseSensitiveInfo = 23;
        private const uint FileCaseSensitiveDirectory = 0x00000001;
        private const uint VolumeNameDos = 0;
        private const uint VolumeNameGuid = 1;
        private const int HashBufferSize = 64 * 1024;
        private const int StreamInfoBufferSize = 64 * 1024;
        private const int StreamInfoNameOffset = 24;
        private const string UnnamedDataStreamName = "::$DATA";

        public const long MaxStageBytes = 512L * 1024 * 1024;
        public const int MaxAncestorHandles = 64;

        private static readonly ValidatedStagePublicationCapability SupportedCapability =
            new ValidatedStagePublicationCapability(
                ValidatedStagePublicationCapabilityCode.Supported,
                "The stage and destination support path-chain-locked NTFS publication.",
                "No corrective action is required.");

        private static readonly ValidatedStagePublicationCapability NotEvaluatedCapability =
            new ValidatedStagePublicationCapability(
                ValidatedStagePublicationCapabilityCode.NotEvaluated,
                "Publication capability was not evaluated because the expected identity is invalid.",
                "Provide a bounded byte length and a 64-digit SHA-256 digest, then retry.");

        /// <summary>
        /// Performs a non-mutating capability check. Publish repeats every check while retaining all protective
        /// handles; a successful probe is advisory and is not a lease.
        /// </summary>
        public static ValidatedStagePublicationCapability CheckCapability(
            string stagePath,
            string destinationPath)
        {
            PublicationContext context;
            ValidatedStagePublicationCapability capability;
            if (!TryAcquireContext(stagePath, destinationPath, out context, out capability))
                return capability;
            using (context) return SupportedCapability;
        }

        public static ValidatedStagePublicationResult Publish(
            string stagePath,
            string destinationPath,
            long expectedByteLength,
            string expectedSha256)
        {
            return PublishCore(
                stagePath,
                destinationPath,
                expectedByteLength,
                expectedSha256,
                null);
        }

        // A deterministic test seam pauses only after the same protected context used by production exists.
        // It is private so callers cannot execute arbitrary code while native handles are retained.
        private static ValidatedStagePublicationResult PublishObserved(
            string stagePath,
            string destinationPath,
            long expectedByteLength,
            string expectedSha256,
            Action protectedContextAcquired)
        {
            if (protectedContextAcquired == null)
                throw new ArgumentNullException(nameof(protectedContextAcquired));
            return PublishCore(
                stagePath,
                destinationPath,
                expectedByteLength,
                expectedSha256,
                protectedContextAcquired);
        }

        private static ValidatedStagePublicationResult PublishCore(
            string stagePath,
            string destinationPath,
            long expectedByteLength,
            string expectedSha256,
            Action protectedContextAcquired)
        {
            string normalizedExpectedSha256;
            if (expectedByteLength < 0 || expectedByteLength > MaxStageBytes ||
                !TryNormalizeSha256(expectedSha256, out normalizedExpectedSha256))
            {
                return Failure(
                    ValidatedStagePublicationStatus.InvalidExpectedIdentity,
                    NotEvaluatedCapability,
                    "The expected stage identity must contain a byte length from zero through " +
                    MaxStageBytes + " and a 64-digit SHA-256 digest.",
                    "Recompute the length and SHA-256 after semantic validation, then retry.");
            }

            PublicationContext context;
            ValidatedStagePublicationCapability capability;
            if (!TryAcquireContext(stagePath, destinationPath, out context, out capability))
            {
                return Failure(
                    ValidatedStagePublicationStatus.CapabilityRejected,
                    capability,
                    capability.Message,
                    capability.CorrectiveAction,
                    capability.NativeErrorCode);
            }

            using (context)
            {
                protectedContextAcquired?.Invoke();
                bool renamed = false;
                HandleIdentity before = default(HandleIdentity);
                try
                {
                    before = ReadIdentity(context.StageHandle, "stage");
                    if (!IsRegularSingleLinkFile(before) || before.ByteLength != expectedByteLength)
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.StageValidationFailed,
                            SupportedCapability,
                            "The protected stage attributes, link count, or length no longer match the validated stage.",
                            "Discard this stage, create and semantically validate a new stage, then retry.");
                    }

                    string beforeSha256 = ComputeSha256(context.StageHandle);
                    if (!string.Equals(beforeSha256, normalizedExpectedSha256, StringComparison.Ordinal))
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.StageValidationFailed,
                            SupportedCapability,
                            "The protected stage SHA-256 no longer matches the validated digest.",
                            "Discard this stage, create and semantically validate a new stage, then retry.");
                    }

                    int existenceError;
                    if (PathEntryExists(context.RenameDestinationPath, out existenceError))
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.DestinationAlreadyExists,
                            SupportedCapability,
                            "The destination already exists and was not changed.",
                            "Choose a new revision name or explicitly resolve the existing artifact.");
                    }
                    if (existenceError != 0 && existenceError != ErrorFileNotFound && existenceError != ErrorPathNotFound)
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.PublicationFailed,
                            SupportedCapability,
                            "The destination absence could not be proven before publication.",
                            "Check destination permissions and retry with a new revision name.",
                            existenceError);
                    }

                    bool singleUnnamedStream;
                    int streamError;
                    if (!TryProveSingleUnnamedDataStream(
                            context.StageHandle,
                            out singleUnnamedStream,
                            out streamError))
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.PublicationFailed,
                            SupportedCapability,
                            "The retained stage stream inventory could not be proven immediately before rename.",
                            "Keep the stage for diagnosis and retry only after correcting the Windows error.",
                            streamError);
                    }
                    if (!singleUnnamedStream)
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.StageValidationFailed,
                            SupportedCapability,
                            "The retained stage acquired a named NTFS alternate data stream before publication.",
                            "Remove the stage and create a new semantically validated single-stream stage.");
                    }

                    HandleIdentity immediatelyBeforeRename = ReadIdentity(
                        context.StageHandle,
                        "stage immediately before rename");
                    VerifyDirectoryChain(context);
                    string beforeRenameDosPath = GetFinalPath(context.StageHandle, VolumeNameDos);
                    string beforeRenameGuidPath = GetFinalPath(context.StageHandle, VolumeNameGuid);
                    if (!before.SameOccurrence(immediatelyBeforeRename) ||
                        !IsRegularSingleLinkFile(immediatelyBeforeRename) ||
                        immediatelyBeforeRename.Attributes != before.Attributes ||
                        immediatelyBeforeRename.ByteLength != expectedByteLength ||
                        !string.Equals(
                            beforeRenameDosPath,
                            context.StageDosFinalPath,
                            StringComparison.Ordinal) ||
                        !string.Equals(
                            beforeRenameGuidPath,
                            context.StageGuidFinalPath,
                            StringComparison.Ordinal))
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.StageValidationFailed,
                            SupportedCapability,
                            "The retained stage identity, attributes, hard-link count, or path changed before rename.",
                            "Discard this stage, create and semantically validate a new stage, then retry.");
                    }

                    int renameError;
                    if (!RenameAbsoluteCreateNew(
                            context.StageHandle,
                            context.RenameDestinationPath,
                            out renameError))
                    {
                        int observedError;
                        bool destinationExists = PathEntryExists(context.RenameDestinationPath, out observedError);
                        bool race = renameError == ErrorAlreadyExists || renameError == ErrorFileExists ||
                                    (renameError == ErrorAccessDenied && destinationExists);
                        return Failure(
                            race
                                ? ValidatedStagePublicationStatus.DestinationRaceLost
                                : ValidatedStagePublicationStatus.PublicationFailed,
                            SupportedCapability,
                            race
                                ? "Another publisher won the destination name; its file was left unchanged."
                                : "Windows rejected the handle-renamed, path-chain-locked create-new publication.",
                            race
                                ? "Choose a new revision name or inspect the winning artifact before retrying."
                                : "Keep the stage for diagnosis and retry only after correcting the reported Windows error.",
                            renameError);
                    }
                    renamed = true;

                    HandleIdentity after = ReadIdentity(context.StageHandle, "published stage");
                    string afterSha256 = ComputeSha256(context.StageHandle);
                    HandleIdentity directoryAfter = ReadIdentity(context.DirectoryHandle, "publication directory");
                    VerifyDirectoryChain(context);
                    if (!TryProveSingleUnnamedDataStream(
                            context.StageHandle,
                            out singleUnnamedStream,
                            out streamError) ||
                        !singleUnnamedStream)
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.PostPublicationVerificationFailed,
                            SupportedCapability,
                            streamError == 0
                                ? "The published occurrence has a named NTFS alternate data stream."
                                : "The final published stream inventory could not be proven.",
                            "Quarantine the destination; do not use or delete it automatically.",
                            streamError,
                            true);
                    }
                    string finalStageDosPath = GetFinalPath(context.StageHandle, VolumeNameDos);
                    string finalStageGuidPath = GetFinalPath(context.StageHandle, VolumeNameGuid);
                    string finalParent = Path.GetDirectoryName(finalStageDosPath);
                    string finalLeaf = Path.GetFileName(finalStageDosPath);

                    if (!before.SameOccurrence(after) ||
                        !context.DirectoryIdentity.SameOccurrence(directoryAfter) ||
                        before.NumberOfLinks != 1 ||
                        after.NumberOfLinks != 1 ||
                        after.Attributes != before.Attributes ||
                        after.ByteLength != expectedByteLength ||
                        !string.Equals(afterSha256, normalizedExpectedSha256, StringComparison.Ordinal) ||
                        !string.Equals(finalParent, context.DirectoryDosFinalPath, StringComparison.Ordinal) ||
                        !string.Equals(finalLeaf, context.DestinationLeaf, StringComparison.Ordinal) ||
                        !string.Equals(
                            finalStageGuidPath,
                            context.RenameDestinationPath,
                            StringComparison.Ordinal))
                    {
                        return Failure(
                            ValidatedStagePublicationStatus.PostPublicationVerificationFailed,
                            SupportedCapability,
                            "The create-new rename completed, but the mandatory post-publication identity proof failed.",
                            "Do not use or delete the destination automatically; quarantine it and inspect the verbose log.",
                            0,
                            true);
                    }

                    var metadata = new ValidatedStagePublicationMetadata(
                        context.StagePath,
                        context.DestinationPath,
                        after.ByteLength,
                        afterSha256,
                        before.OccurrenceToken,
                        after.OccurrenceToken,
                        directoryAfter.OccurrenceToken,
                        context.RenameDestinationPath,
                        true,
                        DateTimeOffset.UtcNow);
                    return new ValidatedStagePublicationResult(
                        ValidatedStagePublicationStatus.Published,
                        SupportedCapability,
                        metadata,
                        "The validated stage was published without replacing an existing destination.",
                        "No corrective action is required.",
                        0,
                        false);
                }
                catch (Exception exception) when (IsOperationalException(exception))
                {
                    return Failure(
                        renamed
                            ? ValidatedStagePublicationStatus.PostPublicationVerificationFailed
                            : ValidatedStagePublicationStatus.PublicationFailed,
                        SupportedCapability,
                        renamed
                            ? "Publication completed, but post-publication verification could not be completed: " + exception.Message
                            : "Publication failed before the destination rename: " + exception.Message,
                        renamed
                            ? "Do not use or delete the destination automatically; quarantine it and inspect the verbose log."
                            : "Keep the stage for diagnosis and correct the reported condition before retrying.",
                        NativeError(exception),
                        renamed);
                }
            }
        }

        internal static int RenameInfoRootDirectoryOffset(int pointerSize)
        {
            if (pointerSize == 4) return 4;
            if (pointerSize == 8) return 8;
            throw new ArgumentOutOfRangeException(nameof(pointerSize));
        }

        internal static int RenameInfoFileNameLengthOffset(int pointerSize)
        {
            return RenameInfoRootDirectoryOffset(pointerSize) + pointerSize;
        }

        internal static int RenameInfoFileNameOffset(int pointerSize)
        {
            return RenameInfoFileNameLengthOffset(pointerSize) + sizeof(int);
        }

        internal static int RenameInfoStructureSize(int pointerSize)
        {
            if (pointerSize == 4) return 16;
            if (pointerSize == 8) return 24;
            throw new ArgumentOutOfRangeException(nameof(pointerSize));
        }

        private static bool TryAcquireContext(
            string stagePath,
            string destinationPath,
            out PublicationContext context,
            out ValidatedStagePublicationCapability capability)
        {
            context = null;
            capability = null;
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.UnsupportedOperatingSystem,
                    "Path-chain-locked NTFS publication is available only on Windows.",
                    "Create the final artifact on a supported Windows host.");
                return false;
            }

            string stage;
            string destination;
            string directory;
            string destinationLeaf;
            try
            {
                if (string.IsNullOrWhiteSpace(stagePath) || string.IsNullOrWhiteSpace(destinationPath))
                    throw new ArgumentException("Both stage and destination paths are required.");
                if (!IsSafeLeaf(Path.GetFileName(stagePath)) ||
                    !IsSafeLeaf(Path.GetFileName(destinationPath)))
                    throw new ArgumentException(
                        "Stage and destination must use ordinary non-device file names.");
                stage = Path.GetFullPath(stagePath);
                destination = Path.GetFullPath(destinationPath);
                if (!string.Equals(stagePath, stage, StringComparison.Ordinal) ||
                    !string.Equals(destinationPath, destination, StringComparison.Ordinal))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.PathAlias,
                        "Stage and destination paths must already be exact absolute normalized DOS paths.",
                        "Use the full long path with Windows-native separators and exact casing.");
                    return false;
                }
                if (IsNetworkOrExtendedPath(stage) || IsNetworkOrExtendedPath(destination))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.NetworkPath,
                        "Network and extended-device paths are outside the proven publication boundary.",
                        "Move the stage to a normal path on a local fixed NTFS drive.");
                    return false;
                }
                string stageDirectory = Path.GetDirectoryName(stage);
                directory = Path.GetDirectoryName(destination);
                destinationLeaf = Path.GetFileName(destination);
                if (string.IsNullOrEmpty(stageDirectory) || string.IsNullOrEmpty(directory) ||
                    !IsSafeLeaf(Path.GetFileName(stage)) || !IsSafeLeaf(destinationLeaf))
                    throw new ArgumentException("Stage and destination must use ordinary file names in an existing directory.");
                if (!string.Equals(stageDirectory, directory, StringComparison.Ordinal))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.DifferentDirectory,
                        "The stage is not a direct child of the destination directory.",
                        "Create the validated stage directly in the final destination directory.");
                    return false;
                }
                if (string.Equals(stage, destination, StringComparison.Ordinal))
                    throw new ArgumentException("Stage and destination must be different file names.");
            }
            catch (Exception exception) when (IsPathException(exception))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.InvalidPath,
                    "The publication paths are invalid: " + exception.Message,
                    "Choose two ordinary file paths in the same existing local directory.");
                return false;
            }

            string root = Path.GetPathRoot(directory);
            if (string.IsNullOrEmpty(root) || root.Length < 3 || root[1] != ':')
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.NetworkPath,
                    "The publication directory is not on a normal local drive path.",
                    "Choose a directory on a local fixed NTFS drive.");
                return false;
            }
            uint driveType = GetDriveTypeW(root);
            if (driveType != DriveFixed)
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.NonFixedVolume,
                    "The publication directory is not on a local fixed drive.",
                    "Move the stage to a directory on a local fixed NTFS drive.");
                return false;
            }
            string driveDeviceTarget;
            int driveDeviceError;
            if (!TryGetDriveDeviceTarget(root, out driveDeviceTarget, out driveDeviceError))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                    "The local drive mapping could not be proven.",
                    "Use a direct local fixed NTFS drive path and inspect the Windows error.",
                    driveDeviceError);
                return false;
            }
            if (driveDeviceTarget.StartsWith(@"\??\", StringComparison.Ordinal) ||
                driveDeviceTarget.StartsWith(@"\DosDevices\", StringComparison.Ordinal))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.PathAlias,
                    "The publication path uses a substituted DOS drive mapping.",
                    "Use the exact path on the underlying local fixed NTFS drive.");
                return false;
            }
            string longDirectoryPath;
            int longDirectoryError;
            if (!TryGetLongPath(directory, out longDirectoryPath, out longDirectoryError))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.DirectoryUnavailable,
                    "The publication directory could not be resolved to its long DOS path.",
                    "Use an existing direct local NTFS directory and inspect the Windows error.",
                    longDirectoryError);
                return false;
            }
            if (!string.Equals(directory, longDirectoryPath, StringComparison.Ordinal))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.PathAlias,
                    "The publication directory uses an 8.3 short-name or casing alias.",
                    "Use the exact long DOS path reported by Windows.");
                return false;
            }

            IList<string> ancestorPaths;
            try
            {
                ancestorPaths = BuildAncestorPaths(directory);
            }
            catch (Exception exception) when (IsPathException(exception))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.InvalidPath,
                    "The publication ancestor chain is invalid: " + exception.Message,
                    "Choose a direct directory path on a local fixed NTFS drive.");
                return false;
            }
            if (ancestorPaths.Count > MaxAncestorHandles)
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.PathTooDeep,
                    "The publication path requires more than " + MaxAncestorHandles + " retained ancestor handles.",
                    "Choose a shallower direct directory path for the validated stage.");
                return false;
            }

            var handles = new List<SafeFileHandle>();
            var handleIdentities = new List<HandleIdentity>();
            var handleDosFinalPaths = new List<string>();
            var handleGuidFinalPaths = new List<string>();
            SafeFileHandle stageHandle = null;
            try
            {
                foreach (string ancestor in ancestorPaths)
                {
                    int error;
                    SafeFileHandle handle = OpenDirectoryProtected(ancestor, out error);
                    if (handle == null)
                    {
                        capability = Unsupported(
                            ValidatedStagePublicationCapabilityCode.DirectoryUnavailable,
                            "A publication directory ancestor could not be locked against substitution.",
                            "Close software holding rename/delete access, verify permissions, and retry.",
                            error);
                        return false;
                    }
                    handles.Add(handle);
                    HandleIdentity identity = ReadIdentity(handle, "publication directory ancestor");
                    if ((identity.Attributes & FileAttributeDirectory) == 0 ||
                        (identity.Attributes & FileAttributeReparsePoint) != 0)
                    {
                        capability = Unsupported(
                            ValidatedStagePublicationCapabilityCode.ReparsePoint,
                            "A publication directory or ancestor is a reparse point.",
                            "Choose a direct directory path with no junctions, symbolic links, or mount points.");
                        return false;
                    }
                    uint caseSensitiveFlags;
                    int caseSensitiveError;
                    if (!TryGetCaseSensitiveFlags(
                            handle,
                            out caseSensitiveFlags,
                            out caseSensitiveError))
                    {
                        capability = Unsupported(
                            ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                            "The case-sensitivity mode of a retained directory handle could not be proven.",
                            "Use a standard case-insensitive NTFS directory on a supported Windows version.",
                            caseSensitiveError);
                        return false;
                    }
                    if ((caseSensitiveFlags & FileCaseSensitiveDirectory) != 0)
                    {
                        capability = Unsupported(
                            ValidatedStagePublicationCapabilityCode.CaseSensitiveDirectory,
                            "A publication directory or ancestor has NTFS per-directory case sensitivity enabled.",
                            "Use a standard case-insensitive NTFS directory for publication.");
                        return false;
                    }
                    string handleDosFinalPath = GetFinalPath(handle, VolumeNameDos);
                    string handleGuidFinalPath = GetFinalPath(handle, VolumeNameGuid);
                    if (!IsVolumeGuidPath(handleGuidFinalPath))
                    {
                        capability = Unsupported(
                            ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                            "Windows did not return a volume-GUID path for a retained directory handle.",
                            "Choose a direct directory on a local fixed NTFS volume.");
                        return false;
                    }
                    if (!string.Equals(ancestor, handleDosFinalPath, StringComparison.Ordinal))
                    {
                        capability = Unsupported(
                            ValidatedStagePublicationCapabilityCode.PathAlias,
                            "A publication ancestor uses a DOS alias, short name, substituted drive, or casing variant.",
                            "Use the exact normalized local DOS path reported by Windows.");
                        return false;
                    }
                    if (handleDosFinalPaths.Count != 0 && (!string.Equals(
                            Path.GetDirectoryName(handleDosFinalPath),
                            handleDosFinalPaths[handleDosFinalPaths.Count - 1],
                            StringComparison.Ordinal) ||
                        !string.Equals(
                            Path.GetDirectoryName(handleGuidFinalPath),
                            handleGuidFinalPaths[handleGuidFinalPaths.Count - 1],
                            StringComparison.Ordinal)))
                    {
                        capability = Unsupported(
                            ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                            "The retained directory handles do not form one direct ancestor chain.",
                            "Choose a direct local directory path with no aliases or redirections.");
                        return false;
                    }
                    handleIdentities.Add(identity);
                    handleDosFinalPaths.Add(handleDosFinalPath);
                    handleGuidFinalPaths.Add(handleGuidFinalPath);
                }

                SafeFileHandle directoryHandle = handles[handles.Count - 1];
                string fileSystem;
                int volumeError;
                if (!TryGetFileSystem(directoryHandle, out fileSystem, out volumeError))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.UnsupportedFileSystem,
                        "The filesystem could not be proven from the retained directory handle.",
                        "Choose a local fixed NTFS directory and retry.",
                        volumeError);
                    return false;
                }
                if (!string.Equals(fileSystem, "NTFS", StringComparison.OrdinalIgnoreCase))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.UnsupportedFileSystem,
                        "The retained publication directory is on " + fileSystem + ", not NTFS.",
                        "Move the stage to a local fixed NTFS directory.");
                    return false;
                }

                HandleIdentity directoryIdentity = ReadIdentity(directoryHandle, "publication directory");
                string directoryDosFinalPath = handleDosFinalPaths[handleDosFinalPaths.Count - 1];
                string directoryGuidFinalPath = handleGuidFinalPaths[handleGuidFinalPaths.Count - 1];
                int stageError;
                stageHandle = OpenStageExclusive(stage, out stageError);
                if (stageHandle == null)
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.StageUnavailable,
                        "The stage could not be opened with an exclusive retained handle.",
                        "Ensure the validated stage exists and close every process using it, then retry.",
                        stageError);
                    return false;
                }
                HandleIdentity stageIdentity = ReadIdentity(stageHandle, "stage");
                if ((stageIdentity.Attributes & (FileAttributeDirectory | FileAttributeReparsePoint)) != 0)
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.ReparsePoint,
                        "The stage is a directory or reparse point, not a direct regular file.",
                        "Create a regular validated stage file directly in the destination directory.");
                    return false;
                }
                string longStagePath;
                int longStageError;
                if (!TryGetLongPath(stage, out longStagePath, out longStageError))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                        "The retained stage could not be resolved to its long DOS path.",
                        "Use the exact long DOS stage path and inspect the Windows error.",
                        longStageError);
                    return false;
                }
                if (!string.Equals(stage, longStagePath, StringComparison.Ordinal))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.PathAlias,
                        "The stage uses an 8.3 short-name or casing alias.",
                        "Use the exact long DOS stage path reported by Windows.");
                    return false;
                }
                if (stageIdentity.NumberOfLinks != 1)
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.StageAliased,
                        "The stage has more than one NTFS hard-link name and is not exclusively owned.",
                        "Copy the validated bytes into a new single-link stage in the destination directory.");
                    return false;
                }
                if (stageIdentity.ByteLength > MaxStageBytes)
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.StageTooLarge,
                        "The stage exceeds the " + MaxStageBytes + "-byte publication bound.",
                        "Create a smaller validated stage or use a future large-artifact publication workflow.");
                    return false;
                }
                bool singleUnnamedStream;
                int streamError;
                if (!TryProveSingleUnnamedDataStream(
                        stageHandle,
                        out singleUnnamedStream,
                        out streamError))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                        "The retained stage stream inventory could not be proven.",
                        "Use a regular single-stream NTFS stage and inspect the Windows error.",
                        streamError);
                    return false;
                }
                if (!singleUnnamedStream)
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.StageAlternateDataStreams,
                        "The stage has one or more named NTFS alternate data streams.",
                        "Create a new semantically validated stage containing only the unnamed data stream.");
                    return false;
                }
                if (stageIdentity.VolumeSerialNumber != directoryIdentity.VolumeSerialNumber)
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.DifferentDirectory,
                        "The retained stage and destination directory are not on the same volume occurrence.",
                        "Create the validated stage directly in the final destination directory.");
                    return false;
                }

                string stageDosFinalPath = GetFinalPath(stageHandle, VolumeNameDos);
                string stageGuidFinalPath = GetFinalPath(stageHandle, VolumeNameGuid);
                if (!IsVolumeGuidPath(stageGuidFinalPath))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                        "Windows did not return a volume-GUID path for the retained stage handle.",
                        "Create the stage on a direct local fixed NTFS path.");
                    return false;
                }
                if (!string.Equals(stage, stageDosFinalPath, StringComparison.Ordinal))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.PathAlias,
                        "The stage path uses a DOS alias, short name, substituted drive, or casing variant.",
                        "Use the exact normalized local DOS path reported by Windows.");
                    return false;
                }
                if (!string.Equals(
                        Path.GetDirectoryName(stageDosFinalPath),
                        directoryDosFinalPath,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        Path.GetDirectoryName(stageGuidFinalPath),
                        directoryGuidFinalPath,
                        StringComparison.Ordinal))
                {
                    capability = Unsupported(
                        ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                        "The retained stage handle could not be proven to be a direct child of the retained directory.",
                        "Recreate the stage directly in the destination directory and retry.");
                    return false;
                }

                context = new PublicationContext(
                    handles,
                    handleIdentities,
                    handleDosFinalPaths,
                    handleGuidFinalPaths,
                    stageHandle,
                    stage,
                    destination,
                    destinationLeaf,
                    directoryHandle,
                    directoryIdentity,
                    directoryDosFinalPath,
                    directoryGuidFinalPath,
                    stageDosFinalPath,
                    stageGuidFinalPath,
                    Path.Combine(directoryGuidFinalPath, destinationLeaf));
                handles = null;
                stageHandle = null;
                capability = SupportedCapability;
                return true;
            }
            catch (Exception exception) when (IsOperationalException(exception))
            {
                capability = Unsupported(
                    ValidatedStagePublicationCapabilityCode.IdentityUnavailable,
                    "A required path-chain and handle identity proof failed: " + exception.Message,
                    "Keep the stage unchanged, inspect the Windows error, and retry on a local fixed NTFS directory.",
                    NativeError(exception));
                return false;
            }
            finally
            {
                if (stageHandle != null) stageHandle.Dispose();
                if (handles != null)
                {
                    for (int index = handles.Count - 1; index >= 0; index--)
                        handles[index].Dispose();
                }
            }
        }

        private static IList<string> BuildAncestorPaths(string directory)
        {
            var stack = new Stack<string>();
            DirectoryInfo current = new DirectoryInfo(directory);
            while (current != null)
            {
                stack.Push(current.FullName);
                current = current.Parent;
            }
            return new List<string>(stack);
        }

        private static SafeFileHandle OpenDirectoryProtected(string path, out int error)
        {
            SafeFileHandle handle = CreateFileW(
                ToNativePath(path),
                FileListDirectory | FileReadAttributes,
                FileShareRead | FileShareWrite,
                IntPtr.Zero,
                OpenExisting,
                FileFlagBackupSemantics | FileFlagOpenReparsePoint,
                IntPtr.Zero);
            if (!handle.IsInvalid)
            {
                error = 0;
                return handle;
            }
            error = Marshal.GetLastWin32Error();
            handle.Dispose();
            return null;
        }

        private static SafeFileHandle OpenStageExclusive(string path, out int error)
        {
            SafeFileHandle handle = CreateFileW(
                ToNativePath(path),
                GenericRead | DeleteAccess | FileReadAttributes,
                0,
                IntPtr.Zero,
                OpenExisting,
                // WRITE_THROUGH asks NTFS to flush metadata-changing operations such as rename. It does not
                // claim hardware power-loss durability, which remains outside this bounded primitive.
                FileFlagBackupSemantics | FileFlagOpenReparsePoint |
                FileFlagSequentialScan | FileFlagWriteThrough,
                IntPtr.Zero);
            if (!handle.IsInvalid)
            {
                error = 0;
                return handle;
            }
            error = Marshal.GetLastWin32Error();
            handle.Dispose();
            return null;
        }

        private static bool RenameAbsoluteCreateNew(
            SafeFileHandle stageHandle,
            string absoluteDestinationPath,
            out int error)
        {
            char[] name = absoluteDestinationPath.ToCharArray();
            int byteLength = checked(name.Length * sizeof(char));
            int nameOffset = RenameInfoFileNameOffset(IntPtr.Size);
            // FILE_RENAME_INFO ends in WCHAR FileName[1] and is padded to native pointer alignment. Windows
            // validates the complete native structure size, not merely the offset of that trailing array.
            int bufferSize = checked(RenameInfoStructureSize(IntPtr.Size) + byteLength);
            IntPtr buffer = Marshal.AllocHGlobal(bufferSize);
            try
            {
                for (int index = 0; index < bufferSize; index++) Marshal.WriteByte(buffer, index, 0);
                Marshal.WriteInt32(buffer, 0, 0); // ReplaceIfExists/Flags: create-new only.
                // The target is the absolute VOLUME_NAME_GUID path obtained from the retained directory
                // occurrence. A null RootDirectory with a leaf would resolve against process CWD, so a leaf is
                // never accepted here. Every verified ancestor also remains open without FILE_SHARE_DELETE.
                Marshal.WriteIntPtr(buffer, RenameInfoRootDirectoryOffset(IntPtr.Size), IntPtr.Zero);
                Marshal.WriteInt32(buffer, RenameInfoFileNameLengthOffset(IntPtr.Size), byteLength);
                Marshal.Copy(name, 0, IntPtr.Add(buffer, nameOffset), name.Length);
                if (SetFileInformationByHandle(
                        stageHandle,
                        FileRenameInfo,
                        buffer,
                        checked((uint)bufferSize)))
                {
                    error = 0;
                    return true;
                }
                error = Marshal.GetLastWin32Error();
                return false;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static HandleIdentity ReadIdentity(SafeFileHandle handle, string subject)
        {
            ByHandleFileInformation information;
            if (!GetFileInformationByHandle(handle, out information))
                throw WindowsIOException("Windows could not identify the retained " + subject + " handle");
            ulong fileIndex = ((ulong)information.FileIndexHigh << 32) | information.FileIndexLow;
            if (fileIndex == 0)
                throw new IOException("Windows returned no stable occurrence identity for the retained " + subject + ".");
            ulong unsignedLength = ((ulong)information.FileSizeHigh << 32) | information.FileSizeLow;
            if (unsignedLength > long.MaxValue)
                throw new IOException("Windows returned an invalid length for the retained " + subject + ".");
            return new HandleIdentity(
                information.VolumeSerialNumber,
                fileIndex,
                (long)unsignedLength,
                information.FileAttributes,
                information.NumberOfLinks);
        }

        private static bool IsRegularSingleLinkFile(HandleIdentity identity)
        {
            return (identity.Attributes & (FileAttributeDirectory | FileAttributeReparsePoint)) == 0 &&
                   identity.NumberOfLinks == 1;
        }

        private static void VerifyDirectoryChain(PublicationContext context)
        {
            for (int index = 0; index < context.DirectoryHandles.Count; index++)
            {
                HandleIdentity current = ReadIdentity(
                    context.DirectoryHandles[index],
                    "publication directory ancestor");
                uint caseSensitiveFlags;
                int caseSensitiveError;
                bool caseSensitivityKnown = TryGetCaseSensitiveFlags(
                    context.DirectoryHandles[index],
                    out caseSensitiveFlags,
                    out caseSensitiveError);
                if (!context.DirectoryIdentities[index].SameOccurrence(current) ||
                    current.Attributes != context.DirectoryIdentities[index].Attributes ||
                    (current.Attributes & FileAttributeDirectory) == 0 ||
                    (current.Attributes & FileAttributeReparsePoint) != 0 ||
                    !caseSensitivityKnown ||
                    (caseSensitiveFlags & FileCaseSensitiveDirectory) != 0 ||
                    !string.Equals(
                        context.DirectoryDosFinalPaths[index],
                        GetFinalPath(context.DirectoryHandles[index], VolumeNameDos),
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        context.DirectoryGuidFinalPaths[index],
                        GetFinalPath(context.DirectoryHandles[index], VolumeNameGuid),
                        StringComparison.Ordinal))
                    throw new IOException(
                        "The retained publication directory chain changed during rename" +
                        (caseSensitiveError == 0 ? "." : " (Win32 " + caseSensitiveError + ")."));
            }
        }

        private static bool TryGetCaseSensitiveFlags(
            SafeFileHandle directoryHandle,
            out uint flags,
            out int error)
        {
            FileCaseSensitiveInformation information;
            bool success = GetFileInformationByHandleEx(
                directoryHandle,
                FileCaseSensitiveInfo,
                out information,
                (uint)Marshal.SizeOf(typeof(FileCaseSensitiveInformation)));
            flags = success ? information.Flags : 0;
            error = success ? 0 : Marshal.GetLastWin32Error();
            return success;
        }

        private static bool TryProveSingleUnnamedDataStream(
            SafeFileHandle stageHandle,
            out bool singleUnnamedDataStream,
            out int error)
        {
            singleUnnamedDataStream = false;
            IntPtr buffer = Marshal.AllocHGlobal(StreamInfoBufferSize);
            try
            {
                if (!GetFileInformationByHandleEx(
                        stageHandle,
                        FileStreamInfo,
                        buffer,
                        StreamInfoBufferSize))
                {
                    error = Marshal.GetLastWin32Error();
                    if (error == ErrorMoreData)
                    {
                        // The one permitted record is tiny; exceeding the bound proves extra stream metadata.
                        singleUnnamedDataStream = false;
                        error = 0;
                        return true;
                    }
                    return false;
                }

                int offset = 0;
                int records = 0;
                while (true)
                {
                    if (offset < 0 || offset > StreamInfoBufferSize - StreamInfoNameOffset)
                    {
                        error = ErrorInvalidData;
                        return false;
                    }
                    uint nextOffset = unchecked((uint)Marshal.ReadInt32(buffer, offset));
                    uint nameByteLength = unchecked((uint)Marshal.ReadInt32(buffer, offset + sizeof(uint)));
                    if ((nameByteLength & 1) != 0 ||
                        nameByteLength > StreamInfoBufferSize - offset - StreamInfoNameOffset)
                    {
                        error = ErrorInvalidData;
                        return false;
                    }
                    string streamName = Marshal.PtrToStringUni(
                        IntPtr.Add(buffer, offset + StreamInfoNameOffset),
                        checked((int)nameByteLength / sizeof(char))) ?? "";
                    records++;
                    if (records != 1 ||
                        !string.Equals(streamName, UnnamedDataStreamName, StringComparison.Ordinal))
                    {
                        singleUnnamedDataStream = false;
                        error = 0;
                        return true;
                    }
                    if (nextOffset == 0)
                    {
                        singleUnnamedDataStream = true;
                        error = 0;
                        return true;
                    }
                    if ((nextOffset & 7) != 0 ||
                        nextOffset < StreamInfoNameOffset + nameByteLength ||
                        nextOffset > StreamInfoBufferSize - offset - StreamInfoNameOffset)
                    {
                        error = ErrorInvalidData;
                        return false;
                    }
                    offset = checked(offset + (int)nextOffset);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static string ComputeSha256(SafeFileHandle handle)
        {
            long ignored;
            if (!SetFilePointerEx(handle, 0, out ignored, 0))
                throw WindowsIOException("Windows could not rewind the retained stage handle");
            using (SHA256 sha256 = SHA256.Create())
            {
                var buffer = new byte[HashBufferSize];
                while (true)
                {
                    int read;
                    if (!ReadFile(handle, buffer, buffer.Length, out read, IntPtr.Zero))
                        throw WindowsIOException("Windows could not read the retained stage handle");
                    if (read == 0) break;
                    sha256.TransformBlock(buffer, 0, read, null, 0);
                }
                sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                return BytesToHex(sha256.Hash);
            }
        }

        private static string GetFinalPath(SafeFileHandle handle, uint volumeNameFormat)
        {
            int capacity = 512;
            while (capacity <= 32768)
            {
                var buffer = new StringBuilder(capacity);
                uint length = GetFinalPathNameByHandleW(
                    handle,
                    buffer,
                    (uint)buffer.Capacity,
                    volumeNameFormat);
                if (length == 0)
                    throw WindowsIOException("Windows could not resolve a retained handle path");
                if (length < buffer.Capacity)
                    return volumeNameFormat == VolumeNameDos
                        ? FromNativeFinalPath(buffer.ToString())
                        : buffer.ToString();
                capacity = checked((int)length + 1);
            }
            throw new IOException("A retained handle path exceeds the supported Windows path length.");
        }

        private static bool TryGetFileSystem(
            SafeFileHandle directoryHandle,
            out string fileSystem,
            out int error)
        {
            var buffer = new StringBuilder(64);
            uint serial;
            uint maximumComponentLength;
            uint flags;
            bool success = GetVolumeInformationByHandleW(
                directoryHandle,
                null,
                0,
                out serial,
                out maximumComponentLength,
                out flags,
                buffer,
                (uint)buffer.Capacity);
            error = success ? 0 : Marshal.GetLastWin32Error();
            fileSystem = success ? buffer.ToString() : "";
            return success;
        }

        private static bool TryGetDriveDeviceTarget(
            string root,
            out string deviceTarget,
            out int error)
        {
            deviceTarget = "";
            if (string.IsNullOrEmpty(root) || root.Length < 2)
            {
                error = ErrorInvalidData;
                return false;
            }
            var buffer = new StringBuilder(32768);
            uint length = QueryDosDeviceW(root.Substring(0, 2), buffer, (uint)buffer.Capacity);
            if (length == 0)
            {
                error = Marshal.GetLastWin32Error();
                return false;
            }
            deviceTarget = buffer.ToString();
            error = 0;
            return !string.IsNullOrEmpty(deviceTarget);
        }

        private static bool TryGetLongPath(
            string path,
            out string longPath,
            out int error)
        {
            var buffer = new StringBuilder(32768);
            uint length = GetLongPathNameW(path, buffer, (uint)buffer.Capacity);
            if (length == 0 || length >= buffer.Capacity)
            {
                longPath = "";
                error = length == 0 ? Marshal.GetLastWin32Error() : ErrorInvalidData;
                return false;
            }
            longPath = buffer.ToString();
            error = 0;
            return true;
        }

        private static bool PathEntryExists(string path, out int error)
        {
            uint attributes = GetFileAttributesW(ToNativePath(path));
            if (attributes != uint.MaxValue)
            {
                error = 0;
                return true;
            }
            error = Marshal.GetLastWin32Error();
            return false;
        }

        private static string ToNativePath(string path)
        {
            return path.StartsWith(@"\\?\", StringComparison.Ordinal) ? path : @"\\?\" + path;
        }

        private static string FromNativeFinalPath(string path)
        {
            if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
                return @"\\" + path.Substring(8);
            if (path.StartsWith(@"\\?\", StringComparison.Ordinal)) return path.Substring(4);
            return path;
        }

        private static bool IsVolumeGuidPath(string path)
        {
            return !string.IsNullOrEmpty(path) &&
                   path.StartsWith(@"\\?\Volume{", StringComparison.Ordinal) &&
                   path.IndexOf('}', 10) > 10;
        }

        private static bool IsNetworkOrExtendedPath(string path)
        {
            return path.StartsWith(@"\\", StringComparison.Ordinal);
        }

        private static bool IsSafeLeaf(string leaf)
        {
            if (string.IsNullOrWhiteSpace(leaf) || leaf == "." || leaf == ".." ||
                leaf.EndsWith(" ", StringComparison.Ordinal) || leaf.EndsWith(".", StringComparison.Ordinal))
                return false;
            return leaf.IndexOfAny(Path.GetInvalidFileNameChars()) < 0 &&
                   !IsReservedDosDeviceLeaf(leaf);
        }

        private static bool IsReservedDosDeviceLeaf(string leaf)
        {
            int extension = leaf.IndexOf('.');
            string baseName = (extension < 0 ? leaf : leaf.Substring(0, extension)).TrimEnd(' ', '.');
            if (baseName.Equals("CON", StringComparison.OrdinalIgnoreCase) ||
                baseName.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
                baseName.Equals("AUX", StringComparison.OrdinalIgnoreCase) ||
                baseName.Equals("NUL", StringComparison.OrdinalIgnoreCase) ||
                baseName.Equals("CLOCK$", StringComparison.OrdinalIgnoreCase) ||
                baseName.Equals("CONIN$", StringComparison.OrdinalIgnoreCase) ||
                baseName.Equals("CONOUT$", StringComparison.OrdinalIgnoreCase))
                return true;

            if (baseName.Length != 4 ||
                !(baseName.StartsWith("COM", StringComparison.OrdinalIgnoreCase) ||
                  baseName.StartsWith("LPT", StringComparison.OrdinalIgnoreCase)))
                return false;

            char suffix = baseName[3];
            return (suffix >= '1' && suffix <= '9') ||
                   suffix == '\u00b9' || suffix == '\u00b2' || suffix == '\u00b3';
        }

        private static bool TryNormalizeSha256(string value, out string normalized)
        {
            normalized = "";
            if (value == null || value.Length != 64) return false;
            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];
                if (!((current >= '0' && current <= '9') ||
                      (current >= 'a' && current <= 'f') ||
                      (current >= 'A' && current <= 'F')))
                    return false;
            }
            normalized = value.ToLowerInvariant();
            return true;
        }

        private static string BytesToHex(byte[] bytes)
        {
            if (bytes == null || bytes.Length != 32)
                throw new IOException("A complete SHA-256 digest was not produced.");
            const string alphabet = "0123456789abcdef";
            var chars = new char[64];
            for (int index = 0; index < bytes.Length; index++)
            {
                chars[index * 2] = alphabet[bytes[index] >> 4];
                chars[index * 2 + 1] = alphabet[bytes[index] & 0x0f];
            }
            return new string(chars);
        }

        private static ValidatedStagePublicationCapability Unsupported(
            ValidatedStagePublicationCapabilityCode code,
            string message,
            string correctiveAction,
            int nativeErrorCode = 0)
        {
            return new ValidatedStagePublicationCapability(code, message, correctiveAction, nativeErrorCode);
        }

        private static ValidatedStagePublicationResult Failure(
            ValidatedStagePublicationStatus status,
            ValidatedStagePublicationCapability capability,
            string message,
            string correctiveAction,
            int nativeErrorCode = 0,
            bool destinationMayContainPublishedStage = false)
        {
            return new ValidatedStagePublicationResult(
                status,
                capability,
                null,
                message,
                correctiveAction,
                nativeErrorCode,
                destinationMayContainPublishedStage);
        }

        private static bool IsPathException(Exception exception)
        {
            return exception is ArgumentException || exception is NotSupportedException ||
                   exception is PathTooLongException;
        }

        private static bool IsOperationalException(Exception exception)
        {
            return exception is IOException || exception is UnauthorizedAccessException ||
                   exception is ArgumentException || exception is NotSupportedException;
        }

        private static int NativeError(Exception exception)
        {
            Exception current = exception;
            while (current != null)
            {
                if (current is Win32Exception win32) return win32.NativeErrorCode;
                current = current.InnerException;
            }
            return 0;
        }

        private static IOException WindowsIOException(string message)
        {
            int error = Marshal.GetLastWin32Error();
            return new IOException(
                message + " (Win32 " + error + ").",
                new Win32Exception(error));
        }

        private sealed class PublicationContext : IDisposable
        {
            private List<SafeFileHandle> _directoryHandles;
            private SafeFileHandle _stageHandle;

            public PublicationContext(
                List<SafeFileHandle> directoryHandles,
                List<HandleIdentity> directoryIdentities,
                List<string> directoryDosFinalPaths,
                List<string> directoryGuidFinalPaths,
                SafeFileHandle stageHandle,
                string stagePath,
                string destinationPath,
                string destinationLeaf,
                SafeFileHandle directoryHandle,
                HandleIdentity directoryIdentity,
                string directoryDosFinalPath,
                string directoryGuidFinalPath,
                string stageDosFinalPath,
                string stageGuidFinalPath,
                string renameDestinationPath)
            {
                _directoryHandles = directoryHandles;
                DirectoryIdentities = directoryIdentities;
                DirectoryDosFinalPaths = directoryDosFinalPaths;
                DirectoryGuidFinalPaths = directoryGuidFinalPaths;
                _stageHandle = stageHandle;
                StagePath = stagePath;
                DestinationPath = destinationPath;
                DestinationLeaf = destinationLeaf;
                DirectoryHandle = directoryHandle;
                DirectoryIdentity = directoryIdentity;
                DirectoryDosFinalPath = directoryDosFinalPath;
                DirectoryGuidFinalPath = directoryGuidFinalPath;
                StageDosFinalPath = stageDosFinalPath;
                StageGuidFinalPath = stageGuidFinalPath;
                RenameDestinationPath = renameDestinationPath;
            }

            public SafeFileHandle StageHandle => _stageHandle;
            public string StagePath { get; }
            public string DestinationPath { get; }
            public string DestinationLeaf { get; }
            public SafeFileHandle DirectoryHandle { get; }
            public IList<SafeFileHandle> DirectoryHandles => _directoryHandles;
            public IList<HandleIdentity> DirectoryIdentities { get; }
            public IList<string> DirectoryDosFinalPaths { get; }
            public IList<string> DirectoryGuidFinalPaths { get; }
            public HandleIdentity DirectoryIdentity { get; }
            public string DirectoryDosFinalPath { get; }
            public string DirectoryGuidFinalPath { get; }
            public string StageDosFinalPath { get; }
            public string StageGuidFinalPath { get; }
            public string RenameDestinationPath { get; }

            public void Dispose()
            {
                SafeFileHandle stage = _stageHandle;
                _stageHandle = null;
                if (stage != null) stage.Dispose();
                List<SafeFileHandle> directories = _directoryHandles;
                _directoryHandles = null;
                if (directories == null) return;
                for (int index = directories.Count - 1; index >= 0; index--)
                    directories[index].Dispose();
            }
        }

        private struct HandleIdentity
        {
            public HandleIdentity(
                uint volumeSerialNumber,
                ulong fileIndex,
                long byteLength,
                uint attributes,
                uint numberOfLinks)
            {
                VolumeSerialNumber = volumeSerialNumber;
                FileIndex = fileIndex;
                ByteLength = byteLength;
                Attributes = attributes;
                NumberOfLinks = numberOfLinks;
            }

            public uint VolumeSerialNumber { get; }
            public ulong FileIndex { get; }
            public long ByteLength { get; }
            public uint Attributes { get; }
            public uint NumberOfLinks { get; }
            public string OccurrenceToken =>
                VolumeSerialNumber.ToString("x8") + ":" + FileIndex.ToString("x16");

            public bool SameOccurrence(HandleIdentity other)
            {
                return VolumeSerialNumber == other.VolumeSerialNumber && FileIndex == other.FileIndex;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ByHandleFileInformation
        {
            public uint FileAttributes;
            public uint CreationTimeLow;
            public uint CreationTimeHigh;
            public uint LastAccessTimeLow;
            public uint LastAccessTimeHigh;
            public uint LastWriteTimeLow;
            public uint LastWriteTimeHigh;
            public uint VolumeSerialNumber;
            public uint FileSizeHigh;
            public uint FileSizeLow;
            public uint NumberOfLinks;
            public uint FileIndexHigh;
            public uint FileIndexLow;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FileCaseSensitiveInformation
        {
            public uint Flags;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(
            string fileName,
            uint desiredAccess,
            uint shareMode,
            IntPtr securityAttributes,
            uint creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(
            SafeFileHandle fileHandle,
            out ByHandleFileInformation fileInformation);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(
            SafeFileHandle fileHandle,
            int fileInformationClass,
            out FileCaseSensitiveInformation fileInformation,
            uint bufferSize);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(
            SafeFileHandle fileHandle,
            int fileInformationClass,
            IntPtr fileInformation,
            int bufferSize);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandleW(
            SafeFileHandle fileHandle,
            StringBuilder filePath,
            uint filePathCharacters,
            uint flags);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetVolumeInformationByHandleW(
            SafeFileHandle fileHandle,
            StringBuilder volumeNameBuffer,
            uint volumeNameSize,
            out uint volumeSerialNumber,
            out uint maximumComponentLength,
            out uint fileSystemFlags,
            StringBuilder fileSystemNameBuffer,
            uint fileSystemNameSize);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern uint GetDriveTypeW(string rootPathName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint QueryDosDeviceW(
            string deviceName,
            StringBuilder targetPath,
            uint targetPathCharacters);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetLongPathNameW(
            string shortPath,
            StringBuilder longPath,
            uint longPathCharacters);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFileAttributesW(string fileName);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetFileInformationByHandle(
            SafeFileHandle fileHandle,
            int fileInformationClass,
            IntPtr fileInformation,
            uint bufferSize);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetFilePointerEx(
            SafeFileHandle fileHandle,
            long distanceToMove,
            out long newFilePointer,
            uint moveMethod);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ReadFile(
            SafeFileHandle fileHandle,
            byte[] buffer,
            int bytesToRead,
            out int bytesRead,
            IntPtr overlapped);
    }
}
