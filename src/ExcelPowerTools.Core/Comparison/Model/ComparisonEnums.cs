namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>The independently reviewable aspect represented by a logical subject.</summary>
    public enum ComparisonAspect
    {
        Unknown = 0,
        Formula = 1,
        Value = 2,
        CachedResult = 3,
        Structure = 4,
        InteriorFill = 5,
        Coverage = 6
    }

    public enum RevisionFormatKind
    {
        Unknown = 0,
        Xlsx = 1,
        Xlsm = 2,
        Xlsb = 3,
        Csv = 4,
        Tsv = 5,
        SeparatedText = 6
    }

    public enum SourcePathClass
    {
        Unknown = 0,
        Local = 1,
        Network = 2,
        CloudSynchronized = 3,
        EmailAttachment = 4,
        PortableArchive = 5
    }

    public enum ProvenanceReliability
    {
        Unknown = 0,
        Weak = 1,
        Moderate = 2,
        Strong = 3
    }

    public enum ObservationExactType
    {
        Missing = 0,
        Blank = 1,
        Formula = 2,
        String = 3,
        Number = 4,
        Boolean = 5,
        Error = 6,
        DateSerial = 7,
        Unsupported = 8,
        IsoDateTime = 9
    }

    public enum CachedResultTrust
    {
        NotPresent = 0,
        Unknown = 1,
        Trusted = 2,
        Stale = 3
    }

    public enum TransitionChangeKind
    {
        None = 0,
        Added = 1,
        Removed = 2,
        Changed = 3,
        Equivalent = 4,
        Reverted = 5,
        Conflict = 6
    }

    public enum ReviewDecisionStatus
    {
        Unreviewed = 0,
        UseMaster = 1,
        UseCandidate = 2,
        Equivalent = 3,
        NeedsInvestigation = 4,
        ReviewedNoStructuralMerge = 5
    }

    public enum DecisionSelectionScope
    {
        SingleSubject = 0,
        MultipleSubjects = 1,
        Group = 2
    }

    public enum MergeRunState
    {
        Planned = 0,
        Prepared = 1,
        Writing = 2,
        Validating = 3,
        Validated = 4,
        OutputPublished = 5,
        AuditCommitted = 6,
        Failed = 7,
        Recovered = 8
    }

    public enum MergeCheckStatus
    {
        NotRun = 0,
        Passed = 1,
        Failed = 2,
        Warning = 3
    }
}
