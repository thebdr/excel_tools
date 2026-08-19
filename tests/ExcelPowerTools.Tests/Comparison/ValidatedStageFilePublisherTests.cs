using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerTools.Core.Comparison.Publication;
using Microsoft.Win32.SafeHandles;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ValidatedStageFilePublisherTests : IDisposable
    {
        private readonly string _directory = Path.Combine(
            Path.GetTempPath(),
            "ExcelPowerTools.Publication.Tests",
            Guid.NewGuid().ToString("N"));

        public ValidatedStageFilePublisherTests()
        {
            Directory.CreateDirectory(_directory);
        }

        [Fact]
        public void FileRenameInfo_offsets_are_explicit_for_x86_and_x64()
        {
            MethodInfo root = LayoutMethod("RenameInfoRootDirectoryOffset");
            MethodInfo length = LayoutMethod("RenameInfoFileNameLengthOffset");
            MethodInfo name = LayoutMethod("RenameInfoFileNameOffset");
            MethodInfo size = LayoutMethod("RenameInfoStructureSize");

            Assert.Equal(4, InvokeOffset(root, 4));
            Assert.Equal(8, InvokeOffset(length, 4));
            Assert.Equal(12, InvokeOffset(name, 4));
            Assert.Equal(16, InvokeOffset(size, 4));
            Assert.Equal(8, InvokeOffset(root, 8));
            Assert.Equal(16, InvokeOffset(length, 8));
            Assert.Equal(20, InvokeOffset(name, 8));
            Assert.Equal(24, InvokeOffset(size, 8));
        }

        [Fact]
        public void Publish_renames_exact_stage_create_new_and_retains_identity()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] bytes = Encoding.UTF8.GetBytes("validated-stage-" + Guid.NewGuid().ToString("N"));
            string stage = CreateStage("exact.stage", bytes);
            string destination = Path.Combine(_directory, "review-r2.xlsx");

            ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                stage,
                destination,
                bytes.LongLength,
                Sha256(bytes));

            Assert.True(
                result.Status == ValidatedStagePublicationStatus.Published,
                result.Status + " / Win32=" + result.NativeErrorCode + " / " + result.Message);
            Assert.True(result.IsPublished);
            Assert.NotNull(result.Metadata);
            Assert.False(result.DestinationMayContainPublishedStage);
            Assert.False(File.Exists(stage));
            Assert.Equal(bytes, File.ReadAllBytes(destination));
            Assert.Equal(Path.GetFullPath(stage), result.Metadata.StagePath);
            Assert.Equal(Path.GetFullPath(destination), result.Metadata.DestinationPath);
            Assert.Equal(bytes.LongLength, result.Metadata.ByteLength);
            Assert.Equal(Sha256(bytes), result.Metadata.Sha256);
            Assert.Equal(result.Metadata.StageOccurrenceToken, result.Metadata.PublishedOccurrenceToken);
            Assert.True(result.Metadata.SingleUnnamedDataStreamObserved);
            Assert.StartsWith(@"\\?\Volume{", result.Metadata.DestinationVolumeGuidPath, StringComparison.Ordinal);
            Assert.Equal(bytes, File.ReadAllBytes(result.Metadata.DestinationVolumeGuidPath));
            Assert.Equal(ReadOccurrenceToken(destination), result.Metadata.PublishedOccurrenceToken);
            Assert.Equal(new[] { destination }, Directory.GetFiles(_directory));
        }

        [Fact]
        public void Existing_destination_is_never_overwritten_and_stage_remains()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] stageBytes = Encoding.UTF8.GetBytes("candidate");
            byte[] winnerBytes = Encoding.UTF8.GetBytes("existing-winner");
            string stage = CreateStage("existing.stage", stageBytes);
            string destination = Path.Combine(_directory, "review-r3.xlsx");
            File.WriteAllBytes(destination, winnerBytes);

            ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                stage,
                destination,
                stageBytes.LongLength,
                Sha256(stageBytes));

            Assert.Equal(ValidatedStagePublicationStatus.DestinationAlreadyExists, result.Status);
            Assert.False(result.IsPublished);
            Assert.Equal(winnerBytes, File.ReadAllBytes(destination));
            Assert.Equal(stageBytes, File.ReadAllBytes(stage));
            Assert.Equal(2, Directory.GetFiles(_directory).Length);
        }

        [Fact]
        public void Mutated_or_substituted_stage_fails_identity_before_publication()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] validated = Encoding.UTF8.GetBytes("validated-original");
            string stage = CreateStage("mutated.stage", validated);
            string destination = Path.Combine(_directory, "mutated.xlsx");
            File.WriteAllBytes(stage, Encoding.UTF8.GetBytes("mutated-content!"));

            ValidatedStagePublicationResult mutated = ValidatedStageFilePublisher.Publish(
                stage,
                destination,
                validated.LongLength,
                Sha256(validated));

            Assert.Equal(ValidatedStagePublicationStatus.StageValidationFailed, mutated.Status);
            Assert.False(File.Exists(destination));

            File.Delete(stage);
            byte[] substitute = Encoding.UTF8.GetBytes("same-byte-length!!");
            Assert.Equal(validated.Length, substitute.Length);
            File.WriteAllBytes(stage, substitute);

            ValidatedStagePublicationResult substituted = ValidatedStageFilePublisher.Publish(
                stage,
                destination,
                validated.LongLength,
                Sha256(validated));

            Assert.Equal(ValidatedStagePublicationStatus.StageValidationFailed, substituted.Status);
            Assert.Equal(substitute, File.ReadAllBytes(stage));
            Assert.False(File.Exists(destination));
        }

        [Fact]
        public void Stage_must_be_a_direct_child_of_the_destination_directory()
        {
            if (!IsSupportedTestVolume()) return;
            string other = Path.Combine(_directory, "other");
            Directory.CreateDirectory(other);
            byte[] bytes = Encoding.UTF8.GetBytes("different-directory");
            string stage = Path.Combine(other, "stage.tmp");
            File.WriteAllBytes(stage, bytes);

            ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                stage,
                Path.Combine(_directory, "result.xlsx"),
                bytes.LongLength,
                Sha256(bytes));

            Assert.Equal(ValidatedStagePublicationStatus.CapabilityRejected, result.Status);
            Assert.Equal(ValidatedStagePublicationCapabilityCode.DifferentDirectory, result.Capability.Code);
            Assert.False(File.Exists(Path.Combine(_directory, "result.xlsx")));
        }

        [Fact]
        public void Network_paths_return_a_corrective_capability_failure_without_io()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

            ValidatedStagePublicationCapability capability =
                ValidatedStageFilePublisher.CheckCapability(
                    @"\\publication-test.invalid\share\stage.tmp",
                    @"\\publication-test.invalid\share\result.xlsx");

            Assert.False(capability.IsSupported);
            Assert.Equal(ValidatedStagePublicationCapabilityCode.NetworkPath, capability.Code);
            Assert.NotEmpty(capability.Message);
            Assert.NotEmpty(capability.CorrectiveAction);
        }

        [Fact]
        public void Invalid_expected_identity_reports_capability_not_evaluated()
        {
            ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                "not-inspected.stage",
                "not-inspected.xlsx",
                -1,
                "invalid");

            Assert.Equal(ValidatedStagePublicationStatus.InvalidExpectedIdentity, result.Status);
            Assert.False(result.Capability.IsSupported);
            Assert.Equal(ValidatedStagePublicationCapabilityCode.NotEvaluated, result.Capability.Code);
        }

        [Fact]
        public void Reserved_dos_device_leaves_are_rejected_before_capability_io()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            string stage = CreateStage("ordinary.stage", new byte[] { 1 });
            string[] reservedLeaves =
            {
                "CON.xlsx",
                "nul",
                "LPT1.report",
                "com9",
                "COM\u00b9.xlsx",
                "LPT\u00b2.txt",
                "CLOCK$",
                "CONIN$.log",
                "CONOUT$",
            };

            foreach (string leaf in reservedLeaves)
            {
                ValidatedStagePublicationCapability capability =
                    ValidatedStageFilePublisher.CheckCapability(
                        stage,
                        Path.Combine(_directory, leaf));

                Assert.False(capability.IsSupported);
                Assert.True(
                    capability.Code == ValidatedStagePublicationCapabilityCode.InvalidPath,
                    leaf + " returned " + capability.Code + ": " + capability.Message);
            }

            Assert.Equal(new[] { stage }, Directory.GetFiles(_directory));
        }

        [Fact]
        public void Stage_size_bound_is_checked_before_hashing()
        {
            if (!IsSupportedTestVolume()) return;
            string stage = Path.Combine(_directory, "oversize.stage");
            using (var stream = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                stream.SetLength(ValidatedStageFilePublisher.MaxStageBytes + 1);

            ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                stage,
                Path.Combine(_directory, "oversize.xlsx"),
                0,
                Sha256(Array.Empty<byte>()));

            Assert.Equal(ValidatedStagePublicationStatus.CapabilityRejected, result.Status);
            Assert.Equal(ValidatedStagePublicationCapabilityCode.StageTooLarge, result.Capability.Code);
            Assert.False(File.Exists(Path.Combine(_directory, "oversize.xlsx")));
            Assert.Equal(ValidatedStageFilePublisher.MaxStageBytes + 1, new FileInfo(stage).Length);
        }

        [Fact]
        public void Ancestor_handle_bound_rejects_deep_path_before_publication()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            string root = Path.Combine(Path.GetTempPath(), "ept-depth-" + Guid.NewGuid().ToString("N"));
            string directory = root;
            try
            {
                for (int index = 0; index < ValidatedStageFilePublisher.MaxAncestorHandles; index++)
                    directory = Path.Combine(directory, "d");
                Directory.CreateDirectory(directory);
                string stage = Path.Combine(directory, "s.tmp");
                File.WriteAllBytes(stage, new byte[] { 1 });

                ValidatedStagePublicationCapability capability =
                    ValidatedStageFilePublisher.CheckCapability(
                        stage,
                        Path.Combine(directory, "r.xlsx"));

                Assert.Equal(ValidatedStagePublicationCapabilityCode.PathTooDeep, capability.Code);
                Assert.False(File.Exists(Path.Combine(directory, "r.xlsx")));
            }
            finally
            {
                try { if (Directory.Exists(root)) Directory.Delete(root, true); }
                catch { }
            }
        }

        [Fact]
        public void Substituted_drive_and_short_name_aliases_are_rejected()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] bytes = Encoding.UTF8.GetBytes("alias-stage");
            string stage = CreateStage("alias.stage", bytes);
            string drive = FindUnusedDriveLetter();
            Assert.True(RunProcess("subst.exe", drive + " \"" + _directory + "\"") == 0,
                "Could not create the temporary SUBST drive.");
            try
            {
                ValidatedStagePublicationCapability capability =
                    ValidatedStageFilePublisher.CheckCapability(
                        drive + "\\alias.stage",
                        drive + "\\alias-result.xlsx");
                Assert.Equal(ValidatedStagePublicationCapabilityCode.PathAlias, capability.Code);
            }
            finally
            {
                RunProcess("subst.exe", drive + " /d");
            }

            string shortDirectory = GetShortPath(_directory);
            if (!string.IsNullOrEmpty(shortDirectory) &&
                !string.Equals(shortDirectory, _directory, StringComparison.Ordinal))
            {
                ValidatedStagePublicationCapability shortCapability =
                    ValidatedStageFilePublisher.CheckCapability(
                        Path.Combine(shortDirectory, Path.GetFileName(stage)),
                        Path.Combine(shortDirectory, "short-result.xlsx"));
                Assert.Equal(ValidatedStagePublicationCapabilityCode.PathAlias, shortCapability.Code);
            }
        }

        [Fact]
        public void Named_ntfs_stream_added_after_capability_probe_is_rejected_before_rename()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] bytes = Encoding.UTF8.GetBytes("single-stream-stage");
            string stage = CreateStage("stream.stage", bytes);
            string destination = Path.Combine(_directory, "stream-result.xlsx");
            Assert.True(ValidatedStageFilePublisher.CheckCapability(stage, destination).IsSupported);
            File.WriteAllText(stage + ":unvalidated", "named-stream-content");

            ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                stage,
                destination,
                bytes.LongLength,
                Sha256(bytes));

            Assert.Equal(ValidatedStagePublicationStatus.CapabilityRejected, result.Status);
            Assert.Equal(
                ValidatedStagePublicationCapabilityCode.StageAlternateDataStreams,
                result.Capability.Code);
            Assert.False(File.Exists(destination));
            Assert.Equal(bytes, File.ReadAllBytes(stage));
        }

        [Fact]
        public void Casing_aliases_are_rejected_or_observed_as_existing_without_rename()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] bytes = Encoding.UTF8.GetBytes("case-alias-stage");
            string stage = CreateStage("CaseAlias.stage", bytes);
            string differentlyCasedDirectory = _directory.ToUpperInvariant();
            if (!string.Equals(differentlyCasedDirectory, _directory, StringComparison.Ordinal))
            {
                ValidatedStagePublicationCapability bothPathsAliased =
                    ValidatedStageFilePublisher.CheckCapability(
                        Path.Combine(differentlyCasedDirectory, Path.GetFileName(stage)),
                        Path.Combine(differentlyCasedDirectory, "both-aliased-result.xlsx"));
                Assert.Equal(
                    ValidatedStagePublicationCapabilityCode.PathAlias,
                    bothPathsAliased.Code);

                ValidatedStagePublicationResult directoryAlias = ValidatedStageFilePublisher.Publish(
                    stage,
                    Path.Combine(differentlyCasedDirectory, "result.xlsx"),
                    bytes.LongLength,
                    Sha256(bytes));
                Assert.Equal(ValidatedStagePublicationStatus.CapabilityRejected, directoryAlias.Status);
                Assert.Equal(
                    ValidatedStagePublicationCapabilityCode.DifferentDirectory,
                    directoryAlias.Capability.Code);
            }

            ValidatedStagePublicationResult leafAlias = ValidatedStageFilePublisher.Publish(
                stage,
                Path.Combine(_directory, "casealias.stage"),
                bytes.LongLength,
                Sha256(bytes));
            Assert.Equal(ValidatedStagePublicationStatus.DestinationAlreadyExists, leafAlias.Status);
            Assert.Equal(bytes, File.ReadAllBytes(stage));
        }

        [Fact]
        public void Case_sensitive_ntfs_directory_is_an_explicit_capability_failure_when_host_allows_it()
        {
            if (!IsSupportedTestVolume()) return;
            string caseSensitiveDirectory = Path.Combine(_directory, "case-sensitive");
            Directory.CreateDirectory(caseSensitiveDirectory);
            if (!TrySetCaseSensitiveInfo(caseSensitiveDirectory, true)) return;
            try
            {
                byte[] bytes = Encoding.UTF8.GetBytes("case-sensitive-stage");
                string stage = Path.Combine(caseSensitiveDirectory, "Stage.tmp");
                File.WriteAllBytes(stage, bytes);

                ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                    stage,
                    Path.Combine(caseSensitiveDirectory, "Result.xlsx"),
                    bytes.LongLength,
                    Sha256(bytes));

                Assert.Equal(ValidatedStagePublicationStatus.CapabilityRejected, result.Status);
                Assert.Equal(
                    ValidatedStagePublicationCapabilityCode.CaseSensitiveDirectory,
                    result.Capability.Code);
                Assert.False(File.Exists(Path.Combine(caseSensitiveDirectory, "Result.xlsx")));
                Assert.Equal(bytes, File.ReadAllBytes(stage));
            }
            finally
            {
                TrySetCaseSensitiveInfo(caseSensitiveDirectory, false);
            }
        }

        [Fact]
        public async Task Two_publishers_never_overwrite_the_winning_destination()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] firstBytes = RepeatedBytes(4 * 1024 * 1024, 0x31);
            byte[] secondBytes = RepeatedBytes(4 * 1024 * 1024, 0x72);
            string firstStage = CreateStage("race-a.stage", firstBytes);
            string secondStage = CreateStage("race-b.stage", secondBytes);
            string destination = Path.Combine(_directory, "race.xlsx");
            using var start = new ManualResetEventSlim(false);

            Task<ValidatedStagePublicationResult> first = Task.Run(() =>
            {
                start.Wait();
                return ValidatedStageFilePublisher.Publish(
                    firstStage, destination, firstBytes.LongLength, Sha256(firstBytes));
            });
            Task<ValidatedStagePublicationResult> second = Task.Run(() =>
            {
                start.Wait();
                return ValidatedStageFilePublisher.Publish(
                    secondStage, destination, secondBytes.LongLength, Sha256(secondBytes));
            });
            start.Set();
            await Task.WhenAll(first, second);

            ValidatedStagePublicationResult[] results = { first.Result, second.Result };
            Assert.Single(results, result => result.Status == ValidatedStagePublicationStatus.Published);
            Assert.Single(results, result =>
                result.Status == ValidatedStagePublicationStatus.DestinationRaceLost ||
                result.Status == ValidatedStagePublicationStatus.DestinationAlreadyExists);
            ValidatedStagePublicationResult winner = results.Single(result => result.IsPublished);
            Assert.Equal(winner.Metadata.Sha256, Sha256(File.ReadAllBytes(destination)));
            string losingStage = winner.Metadata.StagePath == Path.GetFullPath(firstStage)
                ? secondStage
                : firstStage;
            Assert.True(File.Exists(losingStage));
            Assert.Equal(2, Directory.GetFiles(_directory).Length);
        }

        [Fact]
        public async Task Destination_race_stress_has_exactly_one_winner_per_name()
        {
            if (!IsSupportedTestVolume()) return;
            for (int iteration = 0; iteration < 12; iteration++)
            {
                string raceDirectory = Path.Combine(_directory, "race-" + iteration);
                Directory.CreateDirectory(raceDirectory);
                string destination = Path.Combine(raceDirectory, "result.xlsx");
                byte[][] payloads = Enumerable.Range(0, 4)
                    .Select(index => RepeatedBytes(256 * 1024, (byte)(iteration * 7 + index + 1)))
                    .ToArray();
                string[] stages = payloads.Select((bytes, index) =>
                {
                    string path = Path.Combine(raceDirectory, "stage-" + index + ".tmp");
                    File.WriteAllBytes(path, bytes);
                    return path;
                }).ToArray();
                using var start = new ManualResetEventSlim(false);
                Task<ValidatedStagePublicationResult>[] tasks = stages.Select((stage, index) => Task.Run(() =>
                {
                    start.Wait();
                    return ValidatedStageFilePublisher.Publish(
                        stage, destination, payloads[index].LongLength, Sha256(payloads[index]));
                })).ToArray();
                start.Set();
                await Task.WhenAll(tasks);

                ValidatedStagePublicationResult[] results = tasks.Select(task => task.Result).ToArray();
                Assert.Single(results, result => result.IsPublished);
                Assert.All(results.Where(result => !result.IsPublished), result => Assert.True(
                    result.Status == ValidatedStagePublicationStatus.DestinationRaceLost ||
                    result.Status == ValidatedStagePublicationStatus.DestinationAlreadyExists,
                    "Unexpected losing status: " + result.Status + " / " + result.Message));
                Assert.Equal(
                    results.Single(result => result.IsPublished).Metadata.Sha256,
                    Sha256(File.ReadAllBytes(destination)));
                Assert.Equal(4, Directory.GetFiles(raceDirectory).Length);
            }
        }

        [Fact]
        public void Stage_and_directory_reparse_points_are_capability_failures()
        {
            if (!IsSupportedTestVolume()) return;
            string real = Path.Combine(_directory, "real");
            Directory.CreateDirectory(real);
            byte[] bytes = Encoding.UTF8.GetBytes("reparse-test");
            string realStage = Path.Combine(real, "stage.tmp");
            File.WriteAllBytes(realStage, bytes);

            string linkedDirectory = Path.Combine(_directory, "linked-directory");
            CreateJunction(linkedDirectory, real);
            ValidatedStagePublicationCapability directoryCapability =
                ValidatedStageFilePublisher.CheckCapability(
                    Path.Combine(linkedDirectory, "stage.tmp"),
                    Path.Combine(linkedDirectory, "result.xlsx"));
            Assert.Equal(ValidatedStagePublicationCapabilityCode.ReparsePoint, directoryCapability.Code);

            string stageReparse = Path.Combine(_directory, "stage-link.tmp");
            CreateJunction(stageReparse, real);
            ValidatedStagePublicationCapability stageCapability =
                ValidatedStageFilePublisher.CheckCapability(
                    stageReparse,
                    Path.Combine(_directory, "result.xlsx"));
            Assert.Equal(ValidatedStagePublicationCapabilityCode.ReparsePoint, stageCapability.Code);

            string realAncestor = Path.Combine(_directory, "real-ancestor");
            string realChild = Path.Combine(realAncestor, "child");
            Directory.CreateDirectory(realChild);
            File.WriteAllBytes(Path.Combine(realChild, "stage.tmp"), bytes);
            string linkedAncestor = Path.Combine(_directory, "linked-ancestor");
            CreateJunction(linkedAncestor, realAncestor);
            ValidatedStagePublicationCapability ancestorCapability =
                ValidatedStageFilePublisher.CheckCapability(
                    Path.Combine(linkedAncestor, "child", "stage.tmp"),
                    Path.Combine(linkedAncestor, "child", "result.xlsx"));
            Assert.Equal(ValidatedStagePublicationCapabilityCode.ReparsePoint, ancestorCapability.Code);
            Assert.False(File.Exists(Path.Combine(_directory, "result.xlsx")));
        }

        [Fact]
        public void Hard_linked_stage_is_rejected_as_not_exclusively_owned()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] bytes = Encoding.UTF8.GetBytes("hard-link-stage");
            string stage = CreateStage("hard-link.stage", bytes);
            string alias = Path.Combine(_directory, "hard-link-alias.stage");
            Assert.True(CreateHardLinkW(alias, stage, IntPtr.Zero),
                "CreateHardLink failed with Win32 " + Marshal.GetLastWin32Error());

            ValidatedStagePublicationResult result = ValidatedStageFilePublisher.Publish(
                stage,
                Path.Combine(_directory, "hard-link-result.xlsx"),
                bytes.LongLength,
                Sha256(bytes));

            Assert.Equal(ValidatedStagePublicationStatus.CapabilityRejected, result.Status);
            Assert.Equal(ValidatedStagePublicationCapabilityCode.StageAliased, result.Capability.Code);
            Assert.Equal(bytes, File.ReadAllBytes(stage));
            Assert.Equal(bytes, File.ReadAllBytes(alias));
            Assert.False(File.Exists(Path.Combine(_directory, "hard-link-result.xlsx")));
        }

        [Fact]
        public void Hard_link_added_after_context_acquisition_fails_before_destination_creation()
        {
            if (!IsSupportedTestVolume()) return;
            byte[] bytes = Encoding.UTF8.GetBytes("late-hard-link-stage");
            string stage = CreateStage("late-hard-link.stage", bytes);
            string alias = Path.Combine(_directory, "late-hard-link-alias.stage");
            string destination = Path.Combine(_directory, "late-hard-link-result.xlsx");

            ValidatedStagePublicationResult result = InvokeObservedPublish(
                stage,
                destination,
                bytes.LongLength,
                Sha256(bytes),
                () => Assert.True(
                    CreateHardLinkW(alias, stage, IntPtr.Zero),
                    "CreateHardLink failed with Win32 " + Marshal.GetLastWin32Error()));

            Assert.Equal(ValidatedStagePublicationStatus.StageValidationFailed, result.Status);
            Assert.False(File.Exists(destination));
            Assert.True(File.Exists(stage));
            Assert.True(File.Exists(alias));
        }

        [Fact]
        public void Retained_directory_chain_denies_ancestor_rename_until_publication_disposes_handles()
        {
            if (!IsSupportedTestVolume()) return;
            string ancestor = Path.Combine(_directory, "locked-ancestor");
            string publicationDirectory = Path.Combine(ancestor, "publication");
            Directory.CreateDirectory(publicationDirectory);
            byte[] bytes = RepeatedBytes(1024 * 1024, 0x5a);
            string stage = Path.Combine(publicationDirectory, "large.stage");
            File.WriteAllBytes(stage, bytes);
            string destination = Path.Combine(publicationDirectory, "result.xlsx");
            string movedAncestor = Path.Combine(_directory, "moved-ancestor");

            MethodInfo publishObserved = typeof(ValidatedStageFilePublisher).GetMethod(
                "PublishObserved",
                BindingFlags.Static | BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("Missing protected-context publication seam.");
            ValidatedStagePublicationResult? result = null;
            Exception? publicationError = null;
            using var protectedContextAcquired = new ManualResetEventSlim(false);
            using var releaseProtectedContext = new ManualResetEventSlim(false);
            Action observer = () =>
            {
                protectedContextAcquired.Set();
                releaseProtectedContext.Wait();
            };
            var publicationThread = new Thread(() =>
            {
                try
                {
                    result = (ValidatedStagePublicationResult?)publishObserved.Invoke(
                        null,
                        new object[] { stage, destination, bytes.LongLength, Sha256(bytes), observer });
                }
                catch (Exception exception)
                {
                    publicationError = exception is TargetInvocationException invocation &&
                                       invocation.InnerException != null
                        ? invocation.InnerException
                        : exception;
                }
            }) { IsBackground = true };
            publicationThread.Start();

            bool acquired = protectedContextAcquired.Wait(TimeSpan.FromSeconds(10));
            if (!acquired)
            {
                releaseProtectedContext.Set();
                publicationThread.Join(TimeSpan.FromSeconds(5));
            }
            Assert.True(acquired, "The retained publication context was not acquired.");

            try
            {
                Exception denied = Record.Exception(() => Directory.Move(ancestor, movedAncestor));
                Assert.True(
                    denied is IOException || denied is UnauthorizedAccessException,
                    "The retained ancestor handle unexpectedly allowed a directory-chain rename.");
            }
            finally
            {
                releaseProtectedContext.Set();
            }

            Assert.True(publicationThread.Join(TimeSpan.FromSeconds(30)),
                "The publication thread did not terminate after releasing the retained handles.");
            Assert.Null(publicationError);
            Assert.NotNull(result);
            Assert.True(result!.IsPublished, result.Status + " / " + result.Message);
            Directory.Move(ancestor, movedAncestor);
            Assert.True(File.Exists(Path.Combine(movedAncestor, "publication", "result.xlsx")));
        }

        public void Dispose()
        {
            if (!Directory.Exists(_directory)) return;
            foreach (string path in Directory.GetDirectories(_directory, "*", SearchOption.AllDirectories)
                         .OrderByDescending(path => path.Length))
            {
                try
                {
                    if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
                        Directory.Delete(path, false);
                }
                catch { }
            }
            try { Directory.Delete(_directory, true); }
            catch { }
        }

        private bool IsSupportedTestVolume()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return false;
            string probe = Path.Combine(_directory, "probe.tmp");
            File.WriteAllBytes(probe, new byte[] { 1 });
            ValidatedStagePublicationCapability capability = ValidatedStageFilePublisher.CheckCapability(
                probe,
                Path.Combine(_directory, "probe-result.tmp"));
            File.Delete(probe);
            Assert.True(capability.IsSupported, capability.Code + ": " + capability.Message);
            return true;
        }

        private string CreateStage(string name, byte[] bytes)
        {
            string path = Path.Combine(_directory, name);
            File.WriteAllBytes(path, bytes);
            return path;
        }

        private static byte[] RepeatedBytes(int count, byte value)
        {
            var bytes = new byte[count];
            Array.Fill(bytes, value);
            return bytes;
        }

        private static string Sha256(byte[] bytes)
        {
            using SHA256 sha256 = SHA256.Create();
            return Convert.ToHexString(sha256.ComputeHash(bytes)).ToLowerInvariant();
        }

        private static MethodInfo LayoutMethod(string name)
        {
            return typeof(ValidatedStageFilePublisher).GetMethod(
                       name,
                       BindingFlags.Static | BindingFlags.NonPublic)
                   ?? throw new InvalidOperationException("Missing layout method: " + name);
        }

        private static int InvokeOffset(MethodInfo method, int pointerSize)
        {
            return (int)(method.Invoke(null, new object[] { pointerSize }) ?? -1);
        }

        private static ValidatedStagePublicationResult InvokeObservedPublish(
            string stage,
            string destination,
            long byteLength,
            string sha256,
            Action observer)
        {
            MethodInfo method = typeof(ValidatedStageFilePublisher).GetMethod(
                "PublishObserved",
                BindingFlags.Static | BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("Missing protected-context publication seam.");
            try
            {
                return (ValidatedStagePublicationResult)(method.Invoke(
                    null,
                    new object[] { stage, destination, byteLength, sha256, observer })
                    ?? throw new InvalidOperationException("Observed publication returned no result."));
            }
            catch (TargetInvocationException exception) when (exception.InnerException != null)
            {
                throw exception.InnerException;
            }
        }

        private static string FindUnusedDriveLetter()
        {
            for (char letter = 'Z'; letter >= 'D'; letter--)
            {
                string drive = letter + ":";
                if (!Directory.Exists(drive + "\\")) return drive;
            }
            throw new InvalidOperationException("No unused drive letter is available for the SUBST regression.");
        }

        private static int RunProcess(string fileName, string arguments)
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            }) ?? throw new InvalidOperationException("Could not start " + fileName + ".");
            process.StandardOutput.ReadToEnd();
            process.StandardError.ReadToEnd();
            process.WaitForExit();
            return process.ExitCode;
        }

        private static string GetShortPath(string path)
        {
            var buffer = new StringBuilder(32768);
            uint length = GetShortPathNameW(path, buffer, (uint)buffer.Capacity);
            return length == 0 || length >= buffer.Capacity ? "" : buffer.ToString();
        }

        private static void CreateJunction(string junction, string target)
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/d /c mklink /J \"" + junction + "\" \"" + target + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            }) ?? throw new InvalidOperationException("Could not start mklink.");
            string stdout = process.StandardOutput.ReadToEnd();
            string stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            Assert.True(process.ExitCode == 0, stdout + Environment.NewLine + stderr);
        }

        private static bool TrySetCaseSensitiveInfo(string directory, bool enabled)
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "fsutil.exe",
                Arguments = "file SetCaseSensitiveInfo \"" + directory + "\" " +
                            (enabled ? "enable" : "disable"),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            });
            if (process == null) return false;
            process.StandardOutput.ReadToEnd();
            process.StandardError.ReadToEnd();
            process.WaitForExit();
            return process.ExitCode == 0;
        }

        private static string ReadOccurrenceToken(string path)
        {
            using SafeFileHandle handle = File.OpenHandle(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            Assert.True(GetFileInformationByHandle(handle, out ByHandleFileInformation information));
            ulong index = ((ulong)information.FileIndexHigh << 32) | information.FileIndexLow;
            return information.VolumeSerialNumber.ToString("x8") + ":" + index.ToString("x16");
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

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(
            SafeFileHandle fileHandle,
            out ByHandleFileInformation fileInformation);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CreateHardLinkW(
            string fileName,
            string existingFileName,
            IntPtr securityAttributes);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetShortPathNameW(
            string longPath,
            StringBuilder shortPath,
            uint shortPathCharacters);
    }
}
