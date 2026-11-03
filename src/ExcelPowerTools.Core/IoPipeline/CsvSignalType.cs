using System.Collections.Generic;

namespace ExcelPowerTools.Core.IoPipeline
{
    /// <summary>
    /// The plant CSV's Type column, as the controller defines it (the `SIGNAL_TYPE_*` defines,
    /// received 2026-11-03) — ONE table that the type picker, the join and the IO_Defines proofing
    /// all read, so the taxonomy can never drift between them again (it had: 9/10 were labelled
    /// VAI/VAO here while the controller calls them AIS/AOS and keeps VAI/VAO at 13/14).
    ///
    /// Two load-bearing facts from the defines: inputs are ODD, outputs are the input + 1 (12 is
    /// skipped to keep that rule), so direction is parity; and the virtual family — the records whose
    /// node/offset/bit are meaningless because no hardware is behind them — is V*, plus AIS/AOS,
    /// which every export on file writes at 0.0.0 (CTRLZONE_DXDANZONE at 0.0.0 on FVT, the
    /// CHUTE_MOTORRUN timers on Jadlog): software-computed analogs, same handling.
    ///
    /// The IO_Defines workbooks spell some acronyms differently from the field tooling; both
    /// spellings resolve (VI/VDI → 5, VO/VDO → 6).
    /// </summary>
    public static class CsvSignalType
    {
        public sealed class Entry
        {
            public int Number;
            public string Acronym;      // the controller's spelling
            public string Description;
            public bool Virtual;        // no hardware behind it: address fields are not a join key
            public bool IsInput => Number % 2 == 1;
        }

        public static readonly IReadOnlyList<Entry> All = new[]
        {
            new Entry { Number = 1,  Acronym = "DI",  Description = "digital input" },
            new Entry { Number = 2,  Acronym = "DO",  Description = "digital output" },
            new Entry { Number = 3,  Acronym = "AI",  Description = "analog input" },
            new Entry { Number = 4,  Acronym = "AO",  Description = "analog output" },
            new Entry { Number = 5,  Acronym = "VI",  Description = "virtual digital input",  Virtual = true },
            new Entry { Number = 6,  Acronym = "VO",  Description = "virtual digital output", Virtual = true },
            new Entry { Number = 7,  Acronym = "DEI", Description = "data exchange in" },
            new Entry { Number = 8,  Acronym = "DEO", Description = "data exchange out" },
            new Entry { Number = 9,  Acronym = "AIS", Description = "analog input, software",  Virtual = true },
            new Entry { Number = 10, Acronym = "AOS", Description = "analog output, software", Virtual = true },
            new Entry { Number = 11, Acronym = "BUS", Description = "remote node state" },
            new Entry { Number = 13, Acronym = "VAI", Description = "virtual analog input",  Virtual = true },
            new Entry { Number = 14, Acronym = "VAO", Description = "virtual analog output", Virtual = true },
        };

        private static readonly Dictionary<string, int> ByAcronym = BuildAcronyms();

        private static Dictionary<string, int> BuildAcronyms()
        {
            var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
            foreach (Entry e in All) map[e.Acronym] = e.Number;
            map["VDI"] = 5;   // the field tooling's spelling of VI
            map["VDO"] = 6;
            return map;
        }

        /// <summary>The entry for a Type cell ("05" normalizes); null for anything the controller does not define.</summary>
        public static Entry Find(string typeValue)
        {
            string t = IoListLogic.Norm(typeValue);
            foreach (Entry e in All)
                if (e.Number.ToString(System.Globalization.CultureInfo.InvariantCulture) == t) return e;
            return null;
        }

        /// <summary>The number behind an acronym as the IO_Defines workbooks or the tooling spell it; 0 when unknown.</summary>
        public static int NumberOf(string acronym) =>
            acronym != null && ByAcronym.TryGetValue(acronym.Trim(), out int n) ? n : 0;

        /// <summary>"1 — DI (digital input)" — the picker line.</summary>
        public static string Label(Entry e) => e.Number + " — " + e.Acronym + " (" + e.Description + ")";

        /// <summary>The picker's choices for one direction, in number order.</summary>
        public static List<Entry> ForDirection(bool isInput)
        {
            var list = new List<Entry>();
            foreach (Entry e in All)
                if (e.IsInput == isInput) list.Add(e);
            return list;
        }
    }
}
