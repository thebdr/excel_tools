using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// Everything a mnemonic needs to reach its SWP destination, as DATA rather than code.
    ///
    /// The legacy tool MUTATED mnemonics so they would ride an existing branch: EMERGENCY became
    /// CABINETAL purely because both prefixes are nine characters long, so the renamed string put
    /// its index and bit on the offsets the CABINET branch already read. That trick cost the real
    /// identity — plant CSVs know the signal by its true name — and it could only ever express
    /// shapes that happened to line up with an existing one.
    ///
    /// A rule instead DECLARES the destination: a regex over the mnemonic, the sheet and tag it
    /// lands on, and where the index and bit come from. Nothing is renamed, a shape no branch
    /// anticipated is one more row, and "recognized but goes nowhere" is expressible (a blank
    /// destination) instead of being a hardcoded exclusion.
    ///
    /// The <c>_Offsets</c> table still applies: a rule names its OBJECT key, and the per-object
    /// mnemonic→entry shift is looked up there exactly as before.
    ///
    /// What a row DOES. Two kinds, because the bridge answers two different questions about
    /// mnemonics and both belong in one table:
    ///
    /// <c>Map</c> — "given this name, where does it go". A regex over the mnemonic plus its
    /// destination.
    ///
    /// <c>Generate</c> — "given this CSV, what alarms EXIST". The photocell producer reads
    /// Sensor.csv, works out which belt each photocell sits on and in what order along it, emits
    /// one belt alarm per photocell, and then seeks a <c>PHT_*</c> row in the IO list to borrow the
    /// address and the electrical identity. No regex over a mnemonic can express that, and it
    /// should not try — but the row still belongs here, so the table stays the ONE place that
    /// answers "what happens to mnemonics". The bespoke part stays in code and the row names it;
    /// what the row carries is the part worth turning: whether it runs at all, and its constants.
    /// </summary>
    public enum SwpRuleKind { Map = 0, Generate = 1 }

    public sealed class SwpMnemonicRule
    {
        public SwpPlantType PlantType;
        public int Order;
        public bool Enabled = true;

        public SwpRuleKind Kind = SwpRuleKind.Map;

        /// <summary>For a Generate row, which producer the row drives ("Photocells", "Strobes").
        /// The name binds to code; an unknown one is inert with a warning, never fatal.</summary>
        public string Generator = "";

        /// <summary>For a Generate row, the mnemonic it EMITS, as a template
        /// ("ALLB_{sys}_{belt:000}_{bit:00}"). The emitted name then flows through the Map rules
        /// like any other, so a generated alarm reaches its destination the same way.</summary>
        public string Emit = "";

        /// <summary>Generator constants as SEMICOLON-separated key=value pairs
        /// ("firstBit=50;mark=AP;desc=Sensor Fault #{sensor}"). Semicolons, not spaces: a value
        /// like a description template contains spaces and must survive intact.</summary>
        public string Params = "";

        /// <summary>The <c>_Offsets</c> key whose entry shift applies ("" = no shift).</summary>
        public string ObjectKey = "";

        public string Match = "";        // regex over the whole mnemonic
        public string Sheet = "";        // template; "" = this rule maps nowhere
        public string Tag = "";          // template; "" = "DB_" + the resolved sheet
        public string Word = "auto";     // auto | Alarm | Warning | State | Alarm2 | …
        public string Index = "";        // {group} | integer | map(Source,{group})
        public string Bit = "";          // {group} | integer
        public string Note = "";

        /// <summary>Why this rule is inert ("" when it compiled).</summary>
        public string Error = "";

        internal Regex Compiled;

        /// <summary>A rule that recognizes a mnemonic on purpose and sends it nowhere — the
        /// AND/OR group members after the first, the EMERGENCY_ON/RST commands.</summary>
        public bool MapsNowhere => Sheet.Trim().Length == 0;

        public bool IsUsable => Error.Length == 0 &&
            (Kind == SwpRuleKind.Generate || Compiled != null);

        /// <summary>Generator constants as a lookup; missing keys fall to the caller's default,
        /// so a row that leaves a cell alone keeps the shipped behaviour.</summary>
        public string Param(string key, string fallback = "")
        {
            foreach (string pair in (Params ?? "").Split(';'))
            {
                int eq = pair.IndexOf('=');
                if (eq <= 0) continue;
                if (pair.Substring(0, eq).Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
                    return pair.Substring(eq + 1).Trim();
            }
            return fallback;
        }

        public int Param(string key, int fallback) =>
            int.TryParse(Param(key), NumberStyles.Integer, CultureInfo.InvariantCulture, out int v)
                ? v : fallback;

        /// <summary>
        /// The TEMPLATE SET this row belongs to ("" = the plant's own map).
        ///
        /// A template is just more rows of this same shape, gathered under a name. One thing
        /// spawns many — a strobe device type expands into its fourteen alarms — and expressing
        /// that as a separate mechanism (its own file format, its own sheet, its own parser) is
        /// how the strobe templates ended up unable to do anything a mnemonic row can do. Sharing
        /// the row means a template row gets the whole vocabulary for free, including
        /// <c>generate:Name</c>: a template row can call custom logic exactly like a top-level one.
        /// </summary>
        public string Set = "";

        /// <summary>The template set this row SPAWNS when it fires ("" = none).</summary>
        public string Spawns => Param("template");
    }

    /// <summary>The workspace values a template may reference, plus the CSV index sources.</summary>
    public sealed class SwpMapContext
    {
        public int FirstLineId = 1;
        public int GroupIndex = 1;
        public int? DeviceIndexOverride;
        public SwpIndexRules Rules;

        /// <summary>Named index maps: "Destination_Parameters" → (normalized key → SWP ordinal).
        /// A rule asking for a source that is absent fails with a legible error rather than
        /// silently mapping to the raw index.</summary>
        public IReadOnlyDictionary<string, IReadOnlyDictionary<string, int>> Sources;
    }

    /// <summary>
    /// The ordered rule table, one set per plant type. First enabled match wins, so a specific
    /// shape is placed above the general one it would otherwise be swallowed by.
    ///
    /// Serialized as TAB-separated lines — never '|', which every alternation regex contains —
    /// so the table round-trips through a worksheet, a text file and version control unchanged.
    /// </summary>
    public sealed class SwpMnemonicMap
    {
        public const string Header =
            "Set\tPlant\tOrder\tOn\tKind\tObject\tMatch\tEmit\tSheet\tTag\tWord\tIndex\tBit\tParams\tNote";

        private static readonly char[] LineEnd = { '\r', '\n' };

        /// <summary>The plant's own map first, then each template set; by Order within either.
        /// A stable order is what lets an untouched table re-save with no version-control diff.</summary>
        private static int Canonical(SwpMnemonicRule a, SwpMnemonicRule b)
        {
            int set = string.Compare(a.Set, b.Set, StringComparison.OrdinalIgnoreCase);
            if (set != 0) return set;
            if (a.PlantType != b.PlantType) return ((int)a.PlantType).CompareTo((int)b.PlantType);
            return a.Order.CompareTo(b.Order);
        }

        private static readonly Regex Placeholder = new Regex(
            @"\{(?<expr>[^{}:]+)(?::(?<format>[^{}]+))?\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex MapCall = new Regex(
            @"^\s*map\s*\(\s*(?<source>[A-Za-z0-9_]+)\s*,\s*(?<arg>[^)]+)\)\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public List<SwpMnemonicRule> Rules = new List<SwpMnemonicRule>();

        // ---- persistence --------------------------------------------------------------------

        /// <summary>
        /// Reads the table. A line that does not parse becomes an INERT rule carrying its reason —
        /// never an exception and never a silent drop: a typo must cost one shape, not the build,
        /// and it must be visible in the sheet that produced it.
        /// </summary>
        public static SwpMnemonicMap Parse(IEnumerable<string> lines, List<SwpLogEntry> log)
        {
            var map = new SwpMnemonicMap();
            if (lines == null) return map;

            int lineNumber = 0;
            foreach (string raw in lines)
            {
                lineNumber++;
                // TrimEnd() with no argument would eat TRAILING EMPTY COLUMNS — a row whose
                // Params and Note are blank must still arrive with its full column count.
                string line = (raw ?? "").TrimEnd(LineEnd);
                if (line.Length == 0) continue;
                if (line.TrimStart().StartsWith("#", StringComparison.Ordinal)) continue;
                // Only the header line itself — a template set legitimately NAMED "Set" must not
                // be swallowed by a prefix test.
                if (line.Equals(Header, StringComparison.OrdinalIgnoreCase)) continue;

                string[] f = line.Split('\t');
                if (f.Length < 7)
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "mnemonic map line " + lineNumber + " has " + f.Length +
                        " column(s), at least 7 are needed (Set, Plant, Order, On, Kind, Object, Match)"));
                    continue;
                }

                var rule = new SwpMnemonicRule
                {
                    Set = At(f, 0),
                    ObjectKey = At(f, 5),
                    Match = At(f, 6),
                    Emit = At(f, 7),
                    Sheet = At(f, 8),
                    Tag = At(f, 9),
                    Word = At(f, 10).Length == 0 ? "auto" : At(f, 10),
                    Index = At(f, 11),
                    Bit = At(f, 12),
                    Params = At(f, 13),
                    Note = At(f, 14),
                };

                // "map" (or blank) / "generate:Name".
                string kind = At(f, 4);
                if (kind.StartsWith("generate", StringComparison.OrdinalIgnoreCase))
                {
                    rule.Kind = SwpRuleKind.Generate;
                    int colon = kind.IndexOf(':');
                    rule.Generator = colon >= 0 ? kind.Substring(colon + 1).Trim() : "";
                }

                if (!TryParsePlant(At(f, 1), out SwpPlantType plant))
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "mnemonic map line " + lineNumber + ": \"" + At(f, 1) +
                        "\" is not a plant type (Msc, Ispc, Ups, Safety)"));
                    continue;
                }
                rule.PlantType = plant;
                rule.Order = int.TryParse(At(f, 2), NumberStyles.Integer, CultureInfo.InvariantCulture,
                    out int order) ? order : lineNumber;
                string on = At(f, 3);
                rule.Enabled = on.Length == 0 || on.Equals("x", StringComparison.OrdinalIgnoreCase) ||
                    on.Equals("true", StringComparison.OrdinalIgnoreCase) || on == "1";

                Compile(rule);
                if (rule.Error.Length > 0)
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "mnemonic map line " + lineNumber + " is inert: " + rule.Error));
                map.Rules.Add(rule);
            }

            // Grouped by plant, then by Order: Map only ever considers one plant's rules, and a
            // deterministic order is what lets the table round-trip through text unchanged.
            map.Rules.Sort(Canonical);
            return map;
        }

        /// <summary>Emits in the canonical (plant, order) order Parse reads back, so the table is
        /// a stable file: re-saving an untouched map produces no version-control diff.</summary>
        public List<string> ToLines()
        {
            var ordered = new List<SwpMnemonicRule>(Rules);
            ordered.Sort(Canonical);

            var lines = new List<string> { Header };
            foreach (SwpMnemonicRule r in ordered)
                lines.Add(string.Join("\t", new[]
                {
                    r.Set,
                    r.PlantType.ToString(),
                    r.Order.ToString(CultureInfo.InvariantCulture),
                    r.Enabled ? "x" : "",
                    r.Kind == SwpRuleKind.Generate ? "generate:" + r.Generator : "map",
                    r.ObjectKey, r.Match, r.Emit, r.Sheet, r.Tag, r.Word, r.Index, r.Bit,
                    r.Params, r.Note,
                }));
            return lines;
        }

        private static void Compile(SwpMnemonicRule rule)
        {
            if (rule.Kind == SwpRuleKind.Generate)
            {
                // A generator's Match is the IO-list NAME it seeks, a template, not a regex.
                if (rule.Generator.Length == 0)
                    rule.Error = "a generate row must name its producer (generate:Photocells)";
                return;
            }
            if (rule.Match.Trim().Length == 0) { rule.Error = "the Match column is empty"; return; }
            try
            {
                rule.Compiled = new Regex(rule.Match.Trim(),
                    RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(250));
            }
            catch (ArgumentException ex) { rule.Error = "the Match regex does not parse — " + ex.Message; }
        }

        private static string At(string[] f, int i) => i < f.Length ? (f[i] ?? "").Trim() : "";

        private static bool TryParsePlant(string s, out SwpPlantType plant)
        {
            plant = SwpPlantType.Msc;
            if ((s ?? "").Length == 0) return false;
            foreach (SwpPlantType candidate in (SwpPlantType[])Enum.GetValues(typeof(SwpPlantType)))
                if (candidate.ToString().Equals(s, StringComparison.OrdinalIgnoreCase))
                { plant = candidate; return true; }
            return false;
        }

        // ---- mapping ------------------------------------------------------------------------

        /// <summary>
        /// The first enabled rule of this plant whose regex matches decides. A rule that maps
        /// nowhere returns a null coordinate with no error — recognized, deliberately not written.
        /// No match at all is the legacy "#Invalid Mnemonic" failure.
        /// </summary>
        public SwpMappedMnemonic Map(string mnemonic, SwpPlantType plantType, SwpMapContext context)
        {
            string m = (mnemonic ?? "").Trim();
            if (m.Length == 0) return Fail("#Empty Mnemonic");

            SwpMnemonicRule rule = RuleFor(m, plantType, out Match hit);
            if (rule == null) return Fail("#Invalid Mnemonic — no rule in the mnemonic map matches it");

            return rule.MapsNowhere
                ? new SwpMappedMnemonic { ObjectType = rule.ObjectKey }
                : Resolve(rule, hit, plantType, context ?? new SwpMapContext());
        }

        /// <summary>
        /// The deciding rule, without resolving anything. This is what makes the table the FILTER
        /// as well as the mapper: "is this a diagnostic at all" is now the same question as "does
        /// a rule claim it", asked before the workspace settings a full mapping would need exist.
        /// Null = no rule claims it.
        /// </summary>
        public SwpMnemonicRule RuleFor(string mnemonic, SwpPlantType plantType) =>
            RuleFor(mnemonic, plantType, out _);

        /// <summary>
        /// The enabled producers for this plant, in table order. A build asks the table which
        /// generators to run rather than hardcoding the list, so unticking a row genuinely
        /// switches one off and the sheet is the whole truth about what creates mnemonics.
        /// </summary>
        public List<SwpMnemonicRule> Generators(SwpPlantType plantType)
        {
            var found = new List<SwpMnemonicRule>();
            foreach (SwpMnemonicRule rule in Rules)
                if (rule.Kind == SwpRuleKind.Generate && rule.Set.Length == 0 &&
                    rule.PlantType == plantType && rule.Enabled && rule.IsUsable)
                    found.Add(rule);
            return found;
        }

        /// <summary>
        /// The rows of one template set, in order. A template is expanded by whatever spawns it —
        /// a strobe device type, a mnemonic naming it in <c>template=</c> — and each row is a full
        /// rule, so a template row can carry its own destination OR call custom logic through
        /// <c>generate:Name</c> exactly as a top-level row does.
        /// </summary>
        public List<SwpMnemonicRule> Template(string set)
        {
            var found = new List<SwpMnemonicRule>();
            if (string.IsNullOrEmpty(set)) return found;
            foreach (SwpMnemonicRule rule in Rules)
                if (rule.Enabled && rule.Set.Equals(set, StringComparison.OrdinalIgnoreCase))
                    found.Add(rule);
            return found;
        }

        /// <summary>Every template set present, for the sheet and for validation.</summary>
        public List<string> TemplateNames()
        {
            var names = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SwpMnemonicRule rule in Rules)
                if (rule.Set.Length > 0 && seen.Add(rule.Set)) names.Add(rule.Set);
            return names;
        }

        /// <summary>The enabled producer of that name, or null — what a call site asks before
        /// running its own bespoke logic.</summary>
        public SwpMnemonicRule Generator(SwpPlantType plantType, string name)
        {
            foreach (SwpMnemonicRule rule in Generators(plantType))
                if (rule.Generator.Equals(name, StringComparison.OrdinalIgnoreCase)) return rule;
            return null;
        }

        public SwpMnemonicRule RuleFor(string mnemonic, SwpPlantType plantType, out Match hit)
        {
            hit = null;
            string m = (mnemonic ?? "").Trim();
            if (m.Length == 0) return null;

            foreach (SwpMnemonicRule rule in Rules)
            {
                // Generate rows declare a PRODUCER; they never claim an incoming mnemonic. Template
                // rows belong to a SET and are expanded by whatever spawns them, so they must not
                // compete with the plant's own map either.
                if (rule.Kind != SwpRuleKind.Map || rule.Set.Length > 0) continue;
                if (rule.PlantType != plantType || !rule.Enabled || !rule.IsUsable) continue;
                Match candidate;
                try { candidate = rule.Compiled.Match(m); }
                catch (RegexMatchTimeoutException) { continue; }
                if (!candidate.Success) continue;
                hit = candidate;
                return rule;
            }
            return null;
        }

        private SwpMappedMnemonic Resolve(SwpMnemonicRule rule, Match hit,
            SwpPlantType plantType, SwpMapContext context)
        {
            if (!TryNumber(rule.Bit, hit, context, out int bit, out string bitError))
                return Fail("#Invalid Mnemonic — Bit: " + bitError);

            if (!TryNumber(rule.Index, hit, context, out int indexValue, out string indexError))
                return Fail("#Invalid Mnemonic — Index: " + indexError);

            // Two different numbers, and conflating them loses information the coherence legs
            // report with: the DEVICE is what the mnemonic says (chute 7), while the ENTRY is
            // where that lands in the SWP table after any map() and the _Offsets shift.
            int rawDevice = indexValue;
            Group device = hit.Groups["i"];
            if (device.Success && int.TryParse(device.Value, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int fromMnemonic))
                rawDevice = fromMnemonic;

            SwpIndexRules rules = context.Rules ?? SwpIndexRules.Official();
            int deviceIdx = context.DeviceIndexOverride
                ?? indexValue + (rule.ObjectKey.Length > 0
                    ? rules.EntryOffset(plantType, rule.ObjectKey)
                    : 0);

            if (!TryTemplate(rule.Sheet, hit, context, out string sheet, out string sheetError))
                return Fail("#Invalid Mnemonic — Sheet: " + sheetError);

            string tag;
            if (rule.Tag.Length == 0) tag = "DB_" + sheet;
            else if (!TryTemplate(rule.Tag, hit, context, out tag, out string tagError))
                return Fail("#Invalid Mnemonic — Tag: " + tagError);

            if (!TryWord(rule, hit, bit, out string word, out bool isWarning, out string wordError))
                return Fail("#Invalid Mnemonic — Word: " + wordError);

            return new SwpMappedMnemonic
            {
                Coordinate = new SwpCoordinate
                {
                    SheetName = sheet,
                    TagType = tag + "." + word,
                    TableEntry = deviceIdx,
                    Bit = bit,
                },
                IsWarning = isWarning,
                ObjectType = ObjectTypeOf(rule),
                RawDeviceNumber = rawDevice,
                ReferenceIndex = ReferenceIndexOf(hit, context),
            };
        }

        /// <summary>
        /// The DWord. "auto" keeps the SWP's own invariant — AL/WR chooses Alarm or Warning and
        /// the number is ceil((bit+1)/32), because a device's bits are split across 32-bit words
        /// while D keeps counting — and needs an "aw" group to read AL/WR from. A literal family
        /// ("Alarm", "Warning") numbers itself the same way; a full literal ("Alarm2", "State")
        /// is taken as written, which is how a shape with no AL/WR at all reaches State.
        /// </summary>
        private static bool TryWord(SwpMnemonicRule rule, Match hit, int bit,
            out string word, out bool isWarning, out string error)
        {
            word = ""; isWarning = false; error = "";
            string spec = (rule.Word ?? "").Trim();
            if (spec.Length == 0) spec = "auto";

            if (spec.Equals("auto", StringComparison.OrdinalIgnoreCase))
            {
                Group aw = hit.Groups["aw"];
                if (!aw.Success)
                { error = "\"auto\" needs an (?<aw>AL|WR) group in the Match regex"; return false; }
                isWarning = aw.Value.Equals("WR", StringComparison.OrdinalIgnoreCase) ||
                            aw.Value.StartsWith("W", StringComparison.OrdinalIgnoreCase);
                word = (isWarning ? "Warning" : "Alarm") + WordNumber(bit);
                return true;
            }
            if (spec.Equals("Alarm", StringComparison.OrdinalIgnoreCase))
            { word = "Alarm" + WordNumber(bit); return true; }
            if (spec.Equals("Warning", StringComparison.OrdinalIgnoreCase))
            { isWarning = true; word = "Warning" + WordNumber(bit); return true; }

            word = spec;
            isWarning = spec.StartsWith("Warning", StringComparison.OrdinalIgnoreCase);
            return true;
        }

        private static int WordNumber(int bit) => (int)Math.Ceiling((bit + 1) / 32.0);

        private static string ObjectTypeOf(SwpMnemonicRule rule) =>
            rule.ObjectKey == SwpIndexRules.GroupCabinetKey ? "CABINET" : rule.ObjectKey;

        /// <summary>The ISPC line / UPS system the row belongs to, for the alarm-strings leg;
        /// −1 when the shape carries no such index (MSC).</summary>
        private static int ReferenceIndexOf(Match hit, SwpMapContext context)
        {
            foreach (string name in new[] { "line", "sys" })
            {
                Group g = hit.Groups[name];
                if (g.Success && int.TryParse(g.Value, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out int value))
                    return value;
            }
            return -1;
        }

        // ---- the little expression language --------------------------------------------------

        private static bool TryNumber(string expression, Match hit, SwpMapContext context,
            out int value, out string error)
        {
            value = 0; error = "";
            string e = (expression ?? "").Trim();
            if (e.Length == 0) { error = "the expression is empty"; return false; }

            Match call = MapCall.Match(e);
            if (call.Success)
            {
                string source = call.Groups["source"].Value;
                if (!TryTemplate(call.Groups["arg"].Value.Trim(), hit, context,
                        out string key, out error)) return false;
                // A rule that asks for a source states that the mnemonic's own index is NOT the
                // SWP entry. Falling back to the raw index when the CSV is missing (which the
                // legacy mapper did for chutes) writes to whatever row that number happens to hit
                // — silently, and wrongly. An absent or incoherent source is an error.
                if (context.Sources == null ||
                    !context.Sources.TryGetValue(source, out IReadOnlyDictionary<string, int> table) ||
                    table == null || table.Count == 0)
                { error = source + " is not loaded — this shape takes its index from it"; return false; }

                string normalized = IoListLogic.Norm(key);
                if (!table.TryGetValue(normalized, out value))
                { error = key + " is not in " + source; return false; }
                return true;
            }

            if (!TryTemplate(e, hit, context, out string text, out error)) return false;
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return true;
            error = "\"" + text + "\" is not a whole number";
            return false;
        }

        /// <summary>
        /// Substitutes {name}, {name:00} and {a+b-1} — regex groups and the workspace's
        /// FirstLineId / GroupIndex, with integer addition and subtraction. Deliberately small:
        /// a rule table is configuration, and an expression language that can surprise its author
        /// is worse than one that says "not supported".
        /// </summary>
        private static bool TryTemplate(string template, Match hit, SwpMapContext context,
            out string result, out string error)
        {
            error = "";
            string t = template ?? "";
            var problem = new StringBuilder();

            result = Placeholder.Replace(t, match =>
            {
                string expression = match.Groups["expr"].Value.Trim();
                string format = match.Groups["format"].Success
                    ? match.Groups["format"].Value.Trim() : "";

                if (!TryArithmetic(expression, hit, context, out int number, out string why))
                {
                    // A non-numeric group is legitimate in a NAME template (an object word).
                    Group direct = hit.Groups[expression];
                    if (format.Length == 0 && direct.Success) return direct.Value;
                    if (problem.Length == 0) problem.Append(why);
                    return match.Value;
                }
                return format.Length > 0
                    ? number.ToString(format, CultureInfo.InvariantCulture)
                    : number.ToString(CultureInfo.InvariantCulture);
            });

            if (problem.Length > 0) { error = problem.ToString(); return false; }
            return true;
        }

        private static bool TryArithmetic(string expression, Match hit, SwpMapContext context,
            out int value, out string error)
        {
            value = 0; error = "";
            var terms = new List<(int Sign, string Token)>();
            int sign = 1, start = 0;
            for (int i = 0; i <= expression.Length; i++)
            {
                if (i == expression.Length || expression[i] == '+' || expression[i] == '-')
                {
                    string token = expression.Substring(start, i - start).Trim();
                    if (token.Length > 0) terms.Add((sign, token));
                    if (i < expression.Length) { sign = expression[i] == '-' ? -1 : 1; start = i + 1; }
                }
            }
            if (terms.Count == 0) { error = "\"" + expression + "\" is empty"; return false; }

            foreach ((int termSign, string token) in terms)
            {
                if (!TryTerm(token, hit, context, out int term))
                { error = "\"" + token + "\" is not a group, a setting or a number"; return false; }
                value += termSign * term;
            }
            return true;
        }

        private static bool TryTerm(string token, Match hit, SwpMapContext context, out int value)
        {
            if (int.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
                return true;
            if (token.Equals("firstLineId", StringComparison.OrdinalIgnoreCase))
            { value = context.FirstLineId; return true; }
            if (token.Equals("groupIndex", StringComparison.OrdinalIgnoreCase))
            { value = context.GroupIndex; return true; }
            Group g = hit.Groups[token];
            return g.Success && int.TryParse(g.Value, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out value);
        }

        private static SwpMappedMnemonic Fail(string error) =>
            new SwpMappedMnemonic { Error = error };
    }
}
