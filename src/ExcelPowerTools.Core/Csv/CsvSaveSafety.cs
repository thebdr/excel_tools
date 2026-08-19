namespace ExcelPowerTools.Core.Csv
{
    public enum CsvSaveDisposition
    {
        UpToDate,
        SafeToWrite,
        AlreadySynchronized,
        ExternalChange,
        Conflict,
        SourceMissing,
        SourceMissingWithLocalChanges,
        CannotVerify,
    }

    /// <summary>Pure three-way save decision: opened baseline vs current disk vs proposed worksheet bytes.</summary>
    public static class CsvSaveSafety
    {
        public static CsvSaveDisposition Assess(
            CsvSourceFingerprint baseline,
            CsvSourceFingerprint observed,
            CsvSourceFingerprint proposed,
            bool localValuesChanged,
            bool sourceReadable = true)
        {
            if (!sourceReadable) return CsvSaveDisposition.CannotVerify;
            if (observed == null)
                return localValuesChanged
                    ? CsvSaveDisposition.SourceMissingWithLocalChanges
                    : CsvSaveDisposition.SourceMissing;

            if (baseline != null && observed.Equals(baseline))
            {
                if (!localValuesChanged) return CsvSaveDisposition.UpToDate;
                return proposed != null && observed.Equals(proposed)
                    ? CsvSaveDisposition.AlreadySynchronized
                    : CsvSaveDisposition.SafeToWrite;
            }

            if (proposed != null && observed.Equals(proposed))
                return CsvSaveDisposition.AlreadySynchronized;

            return localValuesChanged ? CsvSaveDisposition.Conflict : CsvSaveDisposition.ExternalChange;
        }
    }
}
