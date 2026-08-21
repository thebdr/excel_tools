using System.Collections.Generic;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The plant families the diagnostic mnemonics encode. The numeric values are the legacy
    /// DiagnosticTool codes (Check!A15) and appear in user-facing messages. SAFETY is the
    /// direct-to-SCADA safety PLC (rules §9.2): MSC-shaped mnemonics on a cabinet-only SWP,
    /// with every index aligned 1:1 zero-based.
    /// </summary>
    public enum SwpPlantType { Msc = 1, Ispc = 2, Ups = 3, Safety = 4 }

    /// <summary>
    /// The SWP cells one matched bit-row can receive. Named after the fixed SWP columns: every
    /// diagnostic sheet shares the layout (row 1 headers) B=Export, G=I/O Addr., J=Device,
    /// L=T.S. Ref., Q/R=Description formulas, T/U=descriptions from the IO list, AA=edit stamp.
    /// </summary>
    public enum SwpField
    {
        ExportB,
        AddressG,
        DeviceJ,
        TsRefL,
        Desc1T,
        Desc2U,
        FormulaQ,        // description-1 formula text (mismatch append / wholesale rewrite)
        FormulaR,        // description-2 formula text
        StampAA,         // last-edit timestamp
        FormulaH,        // bit-row FUnit formula repair (=H{blockHeader})
        FormulaI,        // bit-row Location formula repair (=I{blockHeader})
        HeaderLocationI, // block-header Location VALUE repair (the lost "+" prefix)
        HeaderFormulaI,  // strobe-sheet header Location formula (=MID(F15,4,12))
        FUnitH,          // bit-row FUnit VALUE (brutal overwrite: the IO identity replaces the block formula)
        LocationI,       // bit-row Location VALUE (brutal overwrite)
    }

    /// <summary>One spine row's overall state, ranked worst-first for display precedence.</summary>
    public enum SwpRowVerdict
    {
        Invalid,          // the mnemonic did not map to a coordinate
        SheetMissing,     // mapped, but the SWP has no such sheet
        RowMissing,       // sheet exists, no row carries the coordinate triple
        DuplicateTarget,  // two or more spine rows resolve to the same SWP row — nothing queued
        LabelMismatch,    // matched, but FUnit/Location differ from the SWP block (writes still queue)
        Pending,          // matched and at least one field differs — actions queued
        UpToDate,         // matched, every field already holds the intended value
    }

    public enum SwpLogSeverity { Info, Warning, Error }

    /// <summary>
    /// What a FU/Location mismatch does to the WRITES (the warning fires either way):
    /// append the row's compacted identity to the descriptions (the device blanks, and
    /// template-only Q/R formulas get the append extension), or overwrite with the IO values
    /// as-is. Exclusive by design — the three legacy switches collapsed into this.
    /// </summary>
    public enum SwpMismatchMode { AppendToDescription = 0, Overwrite = 1 }

    /// <summary>One build/apply finding, rendered on the _Log sheet and streamed to the dashboard.</summary>
    public sealed class SwpLogEntry
    {
        public SwpLogSeverity Severity = SwpLogSeverity.Info;
        public string CardSheet = "";   // IO-list sheet the row came from ("" when not row-bound)
        public int SourceRow;           // IO-list row number (0 when not row-bound)
        public string Mnemonic = "";
        public string Message = "";

        public static SwpLogEntry For(SwpLogSeverity severity, string message,
            string cardSheet = "", int sourceRow = 0, string mnemonic = "")
        {
            return new SwpLogEntry
            {
                Severity = severity,
                Message = message,
                CardSheet = cardSheet ?? "",
                SourceRow = sourceRow,
                Mnemonic = mnemonic ?? "",
            };
        }
    }

    /// <summary>
    /// Where a diagnostic mnemonic lives inside the SWP: the worksheet plus the row-identity triple
    /// the legacy tool scans for — column A tag type ("DB_CABINETS_DIAG.Alarm1"), column C table
    /// entry (device index) and column D bit number. The legacy "Coordinate" string is
    /// "SheetName|TagType&gt;entry&gt;bit".
    /// </summary>
    public sealed class SwpCoordinate
    {
        public string SheetName = "";
        public string TagType = "";
        public int TableEntry;
        public int Bit;

        /// <summary>Identity within one sheet — what the row scan compares.</summary>
        public string RowKey => TagType + ">" + TableEntry + ">" + Bit;

        /// <summary>The legacy full coordinate, unique across the workbook.</summary>
        public string FullKey => SheetName + "|" + RowKey;

        public override string ToString() => FullKey;
    }

    /// <summary>
    /// One IO-list source row as the filter consumes it — raw cell texts, no interpretation.
    /// The net48 builder fills these from the card sheets; the UPS generators synthesize them.
    /// </summary>
    public sealed class IoRowData
    {
        public string CardSheet = "";
        public int SourceRow;
        public string MnemonicCell = "";  // may hold several mnemonics joined with "/"
        public string AddressCell = "";   // the "Bit" column: "I.80.3.0" ("Q." = output), "" on module rows
        public string IdCell = "";        // module rows carry the bus node here
        public string SlotCell = "";      // AS-I slave rows: the gateway line ("1M" / "2M"); empty = line 1
        public string ModuleNode = "";    // the bus module the row hangs under (a gateway's slaves follow its row)
        public string TypeCell = "";      // IO-list Type ("PA" on module rows)
        public string Desc1A = "";        // description language 1, permanent part
        public string Desc1B = "";        // description language 1, variable part
        public string Desc2A = "";
        public string Desc2B = "";
        public string FUnit = "";
        public string Location = "";
        public string Device = "";
        public string TsRef = "";
        public string ProfinetName = "";
        public string ProfinetIp = "";
    }

    /// <summary>
    /// One accepted diagnostic signal: a single mnemonic (multi-mnemonic cells are split), its
    /// source row data, and the decoded address parts. <c>GeneratedBy</c> is "" for IO-list rows,
    /// "AP" for generated photocell alarms, "ASD" for strobe-device alarms (legacy markers).
    /// </summary>
    public sealed class DiagnosticEntry
    {
        public string CardSheet = "";
        public int SourceRow;
        public string RawMnemonic = "";   // as written in the cell part, before rewrites
        public string Mnemonic = "";      // after BUS_ strip and the EMERGENCY/EMERI rewrites
        public bool IsBusModule;          // no bit address; node from the ID column
        public string Node = "";
        public string Offset = "";
        public string Bit = "";
        public string TypeCell = "";
        public string Desc1A = "";
        public string Desc1B = "";
        public string Desc2A = "";
        public string Desc2B = "";
        public string FUnit = "";
        public string Location = "";
        public string Device = "";
        public string TsRef = "";
        public string ProfinetName = "";
        public string ProfinetIp = "";

        /// <summary>AS-I slave identity "line.slave" ("1M.24A") — "" unless the ID is a slave address.</summary>
        public string AsiNode = "";

        /// <summary>The AS-I gateway's Profinet node the slave hangs off ("90"), taken from the
        /// IO list's row order (a slave row sits under its gateway row); "" when unknown.</summary>
        public string AsiMasterNode = "";
        public string GeneratedBy = "";

        /// <summary>
        /// The I/O address as the SWP receives it (legacy List column C: no "I." prefix). An
        /// AS-I slave is its composed identity ("1M.24A" — line.slave, no trailing ".0.0"); a
        /// Profinet module keeps the legacy "236.0.0". A generated row without a resolved
        /// address yields "" — never a fragment.
        /// </summary>
        public string SwpAddress() =>
            AsiNode.Length > 0 ? AsiNode :
            IsBusModule ? Node + ".0.0" :
            Node.Length == 0 ? "" : Node + "." + Offset + "." + Bit;
    }

    /// <summary>
    /// One SWP bit-row as read from the file bytes — cached formula values are the source (every
    /// formula cell Excel saved carries its last computed value). The identity triple lives in
    /// TagType/TableEntry/Bit; the block fields come from the enclosing device-block header row
    /// ("Table entry:" in column B) and feed the FUnit/Location mismatch check.
    /// </summary>
    public sealed class SwpRowSnapshot
    {
        public int RowNumber;
        public string TagType = "";      // column A ("DB_CABINETS_DIAG.Alarm1")
        public int TableEntry;           // column C
        public int Bit;                  // column D
        public string ExportB = "";
        public string PriorityE = "";    // column E — must be numeric ≠ 0 on exported rows
        public bool PriorityIsText;      // stored as text ("General vs Text category" — rules §3.1)
        public string AddressG = "";
        public string DeviceJ = "";
        public string TsRefL = "";
        public string Desc1T = "";
        public string Desc2U = "";
        public string FormulaQ = "";     // formula text, "" when the cell is a literal or absent
        public string FormulaR = "";
        public string CachedQ = "";      // the displayed description-1 text
        public string CachedR = "";
        public string StampAA = "";
        public bool FormulaQShared;      // t="shared" — the patcher refuses formula edits on these
        public bool FormulaRShared;
        public bool ExportBIsFormula;    // B still pulls its template default from CONFIG_PLANT

        // Structural-integrity capture (the formula cleaner): the FUnit/Location cells of a
        // healthy row are (shared) formulas pointing at their block header — hand edits break that.
        public string FormulaH = "";
        public string FormulaI = "";
        public bool FormulaHShared;
        public bool FormulaIShared;
        public bool HasCellH;            // the cell exists at all (a pasted-blank deletes it)
        public bool HasCellI;
        public string CachedH = "";
        public string CachedI = "";

        /// <summary>Bit n set = column n+1 (A..U) holds a formula; the audit compares masks.</summary>
        public uint FormulaMask;
        public uint SharedMask;

        /// <summary>Per column (1-based), the formula text of EXTERNAL-reference formulas only.</summary>
        public Dictionary<int, string> ExternalFormulas;
        public string BlockFUnit = "";   // block header column H
        public string BlockLocation = ""; // block header column I
        public string BlockLabel = "";   // block header column F (electrical label)
        public int BlockHeaderRow;

        public string RowKey => TagType + ">" + TableEntry + ">" + Bit;
    }

    /// <summary>One device-block header row: the identity cells the repairs read and write.</summary>
    public sealed class SwpBlockHeader
    {
        public int RowNumber;
        public string LabelF = "";           // electrical label (cached)
        public string FUnitH = "";
        public string LocationI = "";        // cached value — must start with "+"
        public string LocationIFormula = ""; // "" when the cell is a literal

        /// <summary>The pseudo-coordinate header edits carry ("header&gt;15").</summary>
        public string Key => "header>" + RowNumber;
    }

    /// <summary>One SWP worksheet indexed for the join: coordinate row-key → snapshot.</summary>
    public sealed class SwpSheetIndex
    {
        public string SheetName = "";
        public string DdeTable = "";     // B4 sheet parameter
        public string Description = ""; // B6 sheet parameter
        public bool IsDiagnosticSheet;   // "# Sheet parameters" marker in A2
        public Dictionary<string, SwpRowSnapshot> RowsByKey =
            new Dictionary<string, SwpRowSnapshot>();
        public List<SwpRowSnapshot> BitRows = new List<SwpRowSnapshot>(); // file order, for the preview
        public List<SwpBlockHeader> BlockHeaders = new List<SwpBlockHeader>();
        public Dictionary<int, SwpBlockHeader> BlockHeaderByRow = new Dictionary<int, SwpBlockHeader>();
        public List<string> DuplicateKeys = new List<string>();           // a broken SWP, logged
        public int BlockCount;
    }

    /// <summary>The SWP workbook as the join consumes it: requested sheets only.</summary>
    public sealed class SwpWorkbookIndex
    {
        public Dictionary<string, SwpSheetIndex> Sheets =
            new Dictionary<string, SwpSheetIndex>();
        public List<string> MissingSheets = new List<string>();
        public List<string> AllSheetNames = new List<string>(); // for wrong-parameter hints
    }

    /// <summary>
    /// A mnemonic resolved against the SWP addressing scheme. <c>Error</c> is "" on success;
    /// otherwise it carries the legacy-style reason ("#Invalid Mnemonic MSC-2") and
    /// <c>Coordinate</c> is null. The raw decoded parts feed the busfield coherence check and the
    /// belt-alarm device rule, which compare the digits as written (no −1 adjustment).
    /// </summary>
    public sealed class SwpMappedMnemonic
    {
        public SwpCoordinate Coordinate;
        public string Error = "";
        public bool IsWarning;
        public string ObjectType = "";    // CABINET / CHUTE / … / BELT / ZONE
        public int RawDeviceNumber = -1;
        public int ReferenceIndex = -1;   // ISPC line (0 = group) / UPS system; -1 for MSC

        public bool Ok => Error.Length == 0;
        public bool IsBeltAlarm => ObjectType == "BELT";
    }
}
