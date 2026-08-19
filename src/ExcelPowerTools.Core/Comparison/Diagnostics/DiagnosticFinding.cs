using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace ExcelPowerTools.Core.Comparison.Diagnostics
{
    /// <summary>A verbose, structured diagnostic occurrence suitable for local and portable sinks.</summary>
    public sealed class DiagnosticFinding
    {
        public string EventId { get; set; } = "";
        public string FindingUid { get; set; } = "";
        public string Code { get; set; } = "";
        public DiagnosticSeverity Severity { get; set; }
        public DateTimeOffset TimestampUtc { get; set; }
        public string RunId { get; set; } = "";
        public string ProjectId { get; set; } = "";
        public string Stage { get; set; } = "";
        public string Operation { get; set; } = "";
        public string MessageKey { get; set; } = "";
        public IList<DiagnosticArgument> MessageArguments { get; set; } = new List<DiagnosticArgument>();
        public string FullDetail { get; set; } = "";
        public IList<DiagnosticArtifactContext> Artifacts { get; set; } = new List<DiagnosticArtifactContext>();
        public IList<DiagnosticLocation> Locations { get; set; } = new List<DiagnosticLocation>();
        public DiagnosticInferenceContext Inference { get; set; } = new DiagnosticInferenceContext();
        public DiagnosticSourceLocation Source { get; set; } = new DiagnosticSourceLocation();
        public DiagnosticExceptionInfo Exception { get; set; }
        public DiagnosticEnvironment Environment { get; set; } = new DiagnosticEnvironment();
        public IList<DiagnosticSuggestedAction> SuggestedActions { get; set; } = new List<DiagnosticSuggestedAction>();
    }

    public sealed class DiagnosticArgument
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
    }

    public sealed class DiagnosticArtifactContext
    {
        public string PathClassification { get; set; } = "";
        public string PathOrBaseName { get; set; } = "";
        public string Sha256 { get; set; } = "";
        public long? ByteLength { get; set; }
        public DateTimeOffset? CreatedUtc { get; set; }
        public DateTimeOffset? ModifiedUtc { get; set; }
        public string Format { get; set; } = "";
        public string RevisionId { get; set; } = "";
    }

    public sealed class DiagnosticLocation
    {
        public string RevisionId { get; set; } = "";
        public string WorkbookLabel { get; set; } = "";
        public string LogicalSheetId { get; set; } = "";
        public string SheetName { get; set; } = "";
        public string Address { get; set; } = "";
        public string LogicalLocator { get; set; } = "";
    }

    public sealed class DiagnosticInferenceContext
    {
        public IList<DiagnosticInferenceCandidate> Candidates { get; set; } = new List<DiagnosticInferenceCandidate>();
        public double? AcceptanceThreshold { get; set; }
        public IList<string> Reasons { get; set; } = new List<string>();
    }

    public sealed class DiagnosticInferenceCandidate
    {
        public string CandidateId { get; set; } = "";
        public double Score { get; set; }
    }

    public sealed class DiagnosticSourceLocation
    {
        public string RepositoryRelativeFile { get; set; } = "";
        public string Member { get; set; } = "";
        public int Line { get; set; }

        public static DiagnosticSourceLocation Capture(
            string repositoryRoot = null,
            [CallerFilePath] string file = "",
            [CallerMemberName] string member = "",
            [CallerLineNumber] int line = 0)
        {
            return new DiagnosticSourceLocation
            {
                RepositoryRelativeFile = MakeRelative(repositoryRoot, file),
                Member = member ?? "",
                Line = line
            };
        }

        private static string MakeRelative(string repositoryRoot, string file)
        {
            if (string.IsNullOrEmpty(file) || string.IsNullOrEmpty(repositoryRoot))
                return file ?? "";

            string root = Path.GetFullPath(repositoryRoot)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string fullFile = Path.GetFullPath(file);
            if (!fullFile.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                return fullFile;

            return fullFile.Substring(root.Length).Replace(Path.DirectorySeparatorChar, '/');
        }
    }

    public sealed class DiagnosticExceptionInfo
    {
        public string Type { get; set; } = "";
        public string Message { get; set; } = "";
        public string StackTrace { get; set; } = "";
        public DiagnosticExceptionInfo InnerException { get; set; }

        public static DiagnosticExceptionInfo FromException(Exception exception)
        {
            if (exception == null)
                return null;

            return new DiagnosticExceptionInfo
            {
                Type = exception.GetType().FullName ?? exception.GetType().Name,
                Message = exception.Message ?? "",
                StackTrace = exception.StackTrace ?? "",
                InnerException = FromException(exception.InnerException)
            };
        }
    }

    public sealed class DiagnosticEnvironment
    {
        public string AddInBuild { get; set; } = "";
        public string ParserVersion { get; set; } = "";
        public string SchemaVersion { get; set; } = "";
        public string ExcelArchitecture { get; set; } = "";
        public string WindowsArchitecture { get; set; } = "";
    }

    public sealed class DiagnosticSuggestedAction
    {
        public string ActionId { get; set; } = "";
        public string LabelKey { get; set; } = "";
        public IList<DiagnosticArgument> Arguments { get; set; } = new List<DiagnosticArgument>();
    }
}
