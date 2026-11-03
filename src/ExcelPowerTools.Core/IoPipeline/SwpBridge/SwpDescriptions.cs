using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>One dictionary row: an IO-list wording replaced by its SWP wording.</summary>
    public sealed class SwpDictionaryEntry
    {
        public string Find = "";
        public string Replace = "";
    }

    /// <summary>
    /// Per-mnemonic steering, persisted in the sidecar and edited on the spine's override columns.
    /// Mirrors the legacy List columns P (skip electrical-label check), Q (SWP DB index,
    /// null = AUTO) and U (force electrical IDs onto belt alarms), plus description overrides.
    /// </summary>
    public sealed class SwpRowOverride
    {
        public string Mnemonic = "";
        public bool SkipLabelCheck;
        public int? SwpDbIndex;
        public bool ForceElectricalIds;

        /// <summary>Per-row FLD treatment: this row writes the IO identity into its own H/I cells
        /// (Overwrite) or appends it to the descriptions (AppendToDescription), whatever the
        /// workspace default says. null = follow the default. Either way the FLD rule holds —
        /// the identity is never lost, only the place it lands changes.</summary>
        public SwpMismatchMode? FldMode;

        /// <summary>Per-row answer to "never overwrite CONFIG_PLANT descriptions": true = this row
        /// keeps the CONFIG template text (the bridge contributes only the FLD label), false = the
        /// bridge writes its composed IO text. null = follow the workspace default, which also
        /// applies its automatic scope (generated + "P*"-typed rows). An explicit tick wins over
        /// that scope — a gesture that does nothing would be a bug.</summary>
        public bool? TrustConfigDescriptions;

        public string Desc1Override = "";
        public string Desc2Override = "";

        public bool IsEmpty =>
            !SkipLabelCheck && !SwpDbIndex.HasValue && !ForceElectricalIds && !FldMode.HasValue &&
            !TrustConfigDescriptions.HasValue &&
            Desc1Override.Length == 0 && Desc2Override.Length == 0;
    }

    /// <summary>
    /// The SWP description pipeline — the C# port of the legacy composition. DESC1 =
    /// dictionary-shortened language-1 text; a Profinet device row (name column filled — AS-I
    /// masters included, the legacy "MODULE AS-I → blank" rule erased correct texts) instead
    /// gets the bus-failure phrase plus the module identity:
    /// "FIELD BUS MODULE FAILURE [{profinet name} - {ip}]". With a glossary loaded the GLOSSARY
    /// is the authority for that phrase (the module's IO description must be an official entry,
    /// else the dedicated FIELD BUS MODULE FAILURE row serves both languages) and for the
    /// T.S. Ref card of every official phrase; the _Languages constants remain only for
    /// glossary-less workspaces. DESC2 uses the workspace's chosen second language and falls
    /// back to DESC1 when it comes out blank — flagged on the result, because a single-language
    /// SCADA then shows language-1 text. Overrides win over everything.
    /// </summary>
    public static class SwpDescriptions
    {
        /// <summary>The language-1 fallback when no language table exists yet.</summary>
        public const string DefaultBusFailureText = "FIELD BUS MODULE FAILURE";

        public sealed class Result
        {
            public string Desc1 = "";
            public string Desc2 = "";

            /// <summary>The IO list carries a language-2 text AND the glossary disagrees with it.</summary>
            public bool GlossaryDiffers;
            public string GlossaryText = "";
            public string IoLang2Text = "";

            /// <summary>The T.S. Ref for column L — the glossary's official card when the
            /// description is an official phrase carrying one, else the IO list's.</summary>
            public string TsRef = "";

            /// <summary>The phrase this row keys the glossary with — its Description permanent
            /// part (or the chosen bus phrase). Feeds the _Glossary sheet's relevance.</summary>
            public string GlossaryKey = "";
            public bool TsRefOverridden;      // the IO list held a DIFFERENT card — glossary won
            public string TsRefFromIo = "";

            /// <summary>A bus module whose IO-list description is not an official glossary
            /// phrase — replaced with the glossary's dedicated bus-failure row.</summary>
            public bool BusDescMismatch;
            public string BusIoText = "";
            public string BusChosenEn = "";

            /// <summary>Description 2 came out with no visible character and repeats Description 1:
            /// language-1 text in the second-language cell (the legacy fallback).</summary>
            public bool Desc2FromLanguage1;

            /// <summary>Description 2 is a glossary translation of the permanent part followed by the
            /// LANGUAGE-1 variable part, because the IO list wrote no language-2 one (a part the
            /// dictionary empties still counts as written — nothing rides then).</summary>
            public bool Desc2VariablePartFromLanguage1;

            /// <summary>The language-1 text that rode into Description 2 — after both dictionaries,
            /// which is what the cell shows; "" unless <see cref="Desc2VariablePartFromLanguage1"/>.</summary>
            public string Desc2RidingText = "";
        }

        /// <summary>Everything the composition needs beyond the entry itself.</summary>
        public sealed class Context
        {
            public IReadOnlyList<SwpDictionaryEntry> Dictionary1;
            public IReadOnlyList<SwpDictionaryEntry> Dictionary2;
            public string BusFailureLang1 = DefaultBusFailureText;
            public string BusFailureLang2 = "";
            public IReadOnlyList<SwpGlossaryPair> GlossaryPairs;   // null = no glossary loaded
            public bool GlossaryTranslateAlways;                    // false = fill empty lang-2 only

            private Dictionary<string, SwpGlossaryPair> _byEn;
            private HashSet<string> _allEn;

            /// <summary>The enabled pair whose EN phrase exactly equals the text (trimmed,
            /// case-insensitive); null without a glossary or without a match.</summary>
            internal SwpGlossaryPair FindExact(string text)
            {
                if (GlossaryPairs == null || string.IsNullOrEmpty(text)) return null;
                if (_byEn == null)
                {
                    _byEn = new Dictionary<string, SwpGlossaryPair>(StringComparer.OrdinalIgnoreCase);
                    foreach (SwpGlossaryPair pair in GlossaryPairs)
                        if (pair != null && pair.Enabled && pair.En.Length > 0 && !_byEn.ContainsKey(pair.En.Trim()))
                            _byEn[pair.En.Trim()] = pair;
                }
                return _byEn.TryGetValue(text.Trim(), out SwpGlossaryPair found) ? found : null;
            }

            /// <summary>Whether ANY pair — disabled included — carries this EN phrase. Keeps
            /// the _Glossary relevance honest: a pair disabled there must stay listed (else it
            /// could never be re-enabled), so the row's glossary key follows the phrase even
            /// while the pair is off.</summary>
            internal bool AnyPairHas(string text)
            {
                if (GlossaryPairs == null || string.IsNullOrEmpty(text)) return false;
                if (_allEn == null)
                {
                    _allEn = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (SwpGlossaryPair pair in GlossaryPairs)
                        if (pair != null && pair.En.Length > 0) _allEn.Add(pair.En.Trim());
                }
                return _allEn.Contains(text.Trim());
            }
        }

        public static Result Compose(DiagnosticEntry entry,
            IReadOnlyList<SwpDictionaryEntry> dictionary1,
            IReadOnlyList<SwpDictionaryEntry> dictionary2,
            string busFailureLang1,
            string busFailureLang2,
            SwpRowOverride overrides = null)
        {
            return Compose(entry, new Context
            {
                Dictionary1 = dictionary1,
                Dictionary2 = dictionary2,
                BusFailureLang1 = busFailureLang1,
                BusFailureLang2 = busFailureLang2,
            }, overrides);
        }

        public static Result Compose(DiagnosticEntry entry, Context context, SwpRowOverride overrides = null)
        {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            context = context ?? new Context();
            IReadOnlyList<SwpDictionaryEntry> dictionary1 = context.Dictionary1;
            IReadOnlyList<SwpDictionaryEntry> dictionary2 = context.Dictionary2;
            string busFailureLang1 = context.BusFailureLang1;
            string busFailureLang2 = context.BusFailureLang2;

            // The legacy List concatenated the permanent and variable description parts with a
            // single space — verbatim, trailing space included when the variable part is empty.
            string lang1 = entry.Desc1A + " " + entry.Desc1B;
            string lang2 = entry.Desc2A + " " + entry.Desc2B;

            // A row with a Profinet name is a fieldbus module — AS-I masters/gateways included:
            // the legacy blanked "MODULE AS-I" rows outright, which proposed ERASING correct
            // FIELD BUS MODULE FAILURE texts in the field (FVZ node 115). AS-I SLAVES have no
            // Profinet name and never enter this branch.
            bool isProfinet = entry.ProfinetName.Trim().Length > 0;
            string moduleTag = isProfinet
                ? " [" + entry.ProfinetName.Trim() + " - " + entry.ProfinetIp.Trim() + "]"
                : "";

            string constant1 = string.IsNullOrEmpty(busFailureLang1) ? DefaultBusFailureText : busFailureLang1;

            // With a glossary loaded, IT is the authority for bus-module texts: the module's
            // IO-list description must be one of the official phrases (used as written there,
            // official casing included); anything else falls to the glossary's dedicated
            // "FIELD BUS MODULE FAILURE" row — never to the _Languages constants, which remain
            // only for glossary-less workspaces.
            var result = new Result();

            // Glossary matching is whole-phrase only (substring splicing produced wrong
            // sentences), in TWO tiers: the Description PERMANENT part first (the variable
            // part rides along), then the trimmed FULL description — field files carry
            // glossary-compliant texts split across the two parts, and a full match consumes
            // both (nothing rides along).
            SwpGlossaryPair busPair = null;
            if (isProfinet && context.GlossaryPairs != null)
            {
                busPair = context.FindExact(entry.Desc1A.Trim()) ?? context.FindExact(lang1.Trim());
                if (busPair == null)
                {
                    busPair = context.FindExact(DefaultBusFailureText);
                    if (busPair != null && lang1.Trim().Length > 0)
                    {
                        result.BusDescMismatch = true;
                        result.BusIoText = lang1.Trim();
                        result.BusChosenEn = busPair.En;
                    }
                }
            }

            string permanentKey = isProfinet ? "" : MultiReplace(entry.Desc1A, dictionary1).Trim();
            string fullKey = isProfinet ? "" : MultiReplace(lang1, dictionary1).Trim();
            SwpGlossaryPair signalPair = permanentKey.Length > 0 ? context.FindExact(permanentKey) : null;
            bool matchedFullDescription = false;
            if (signalPair == null && fullKey.Length > 0)
            {
                signalPair = context.FindExact(fullKey);
                matchedFullDescription = signalPair != null;
            }
            // The relevance key follows whichever tier hits ANY pair — disabled ones included,
            // or a pair switched off on _Glossary would vanish from the sheet for good.
            string signalKey = !isProfinet && !context.AnyPairHas(permanentKey) && context.AnyPairHas(fullKey)
                ? fullKey
                : permanentKey;

            string desc1;
            if (overrides != null && overrides.Desc1Override.Length > 0)
                desc1 = overrides.Desc1Override;
            else if (isProfinet)
                desc1 = (busPair != null ? busPair.En : constant1) + moduleTag;
            else
                desc1 = MultiReplace(lang1, dictionary1);

            result.Desc1 = desc1;

            string desc2;
            if (overrides != null && overrides.Desc2Override.Length > 0)
            {
                desc2 = overrides.Desc2Override;   // decided by the user — no glossary opinions
            }
            else if (isProfinet)
            {
                desc2 = busPair != null
                    ? (busPair.Translation.Trim().Length > 0 ? busPair.Translation + moduleTag : "")
                    : string.IsNullOrEmpty(busFailureLang2) ? "" : busFailureLang2 + moduleTag;
            }
            else
            {
                string ioLang2 = MultiReplace(lang2, dictionary2);
                bool hasIoLang2 = ioLang2.Replace(" ", "").Length > 0;

                // A permanent-part match keeps the variable part alongside (the electrical
                // identity); a full-description match consumed everything — the translation stands
                // alone. The variable part rides in LANGUAGE 2 when the IO list has one: that is its
                // own translation of the part, and a single-language SCADA shows nothing else
                // ("DISJONCTEUR DÉCLENCHÉ 24V IM 155 ET MODULES I/O", never "… AND I/O MODULE" —
                // FVZ, 2026-09-10). Only an IO list with no language-2 variable part lets language 1's
                // ride, and the result says so.
                //
                // "Has one" means the IO list WROTE one: a part the dictionary then empties
                // ("CH1/CH2" under a rule deleting the channel markers — FVT, 2026-09-22) still
                // counts as written, so nothing rides — never language 1's raw text, which is what
                // put "CH1/CH2" back into a Description 2 the rules had cleaned. And when language
                // 1's part does ride, it passes dictionary 1 (it is language-1 text) and then
                // dictionary 2: text landing in Description 2 always passes dictionary 2 last.
                string variable2 = MultiReplace(entry.Desc2B, dictionary2);
                bool ioHasVariable2 = entry.Desc2B.Trim().Length > 0;
                string ridingVariable = ioHasVariable2
                    ? variable2
                    : MultiReplace(MultiReplace(entry.Desc1B, dictionary1), dictionary2);
                string glossaryText = null;
                if (signalPair != null && signalPair.Translation.Trim().Length > 0)
                    glossaryText = matchedFullDescription
                        ? signalPair.Translation
                        : signalPair.Translation + " " + ridingVariable;

                bool glossaryWins = glossaryText != null && (context.GlossaryTranslateAlways || !hasIoLang2);
                desc2 = glossaryWins ? glossaryText : ioLang2;
                result.Desc2VariablePartFromLanguage1 = glossaryWins && !matchedFullDescription &&
                    !ioHasVariable2 && ridingVariable.Trim().Length > 0;
                if (result.Desc2VariablePartFromLanguage1) result.Desc2RidingText = ridingVariable.Trim();

                // The IO list and the standard glossary disagreeing is worth a look either way.
                if (hasIoLang2 && glossaryText != null &&
                    !string.Equals(glossaryText.Trim(), ioLang2.Trim(), StringComparison.Ordinal))
                {
                    result.GlossaryDiffers = true;
                    result.GlossaryText = glossaryText;
                    result.IoLang2Text = ioLang2;
                }
            }

            // The language-2 formula fell back to DESC1 whenever its own result held no visible
            // character (the second language is frequently unfilled in the IO list).
            // Language-1 text landing in Description 2 passes dictionary 2 last — except a bus
            // module's phrase + [profinet name - ip], which no dictionary ever touches: the tag is
            // the module's identity, not a description to shorten.
            if (desc2.Replace(" ", "").Length == 0)
            {
                desc2 = isProfinet ? desc1 : MultiReplace(desc1, dictionary2);
                result.Desc2FromLanguage1 = desc2.Replace(" ", "").Length > 0;
            }

            result.Desc2 = desc2;

            // T.S. Ref: the glossary owns the card of every official phrase (its "Numero
            // scheda" column); the IO list's value stands only where the description is not an
            // official phrase — electrical alarms keep their IO-list cards (rules §12.2).
            result.TsRef = entry.TsRef;
            result.GlossaryKey = busPair != null ? busPair.En : signalKey;
            SwpGlossaryPair tsPair = busPair ?? signalPair;
            if (tsPair != null && tsPair.TsRef.Trim().Length > 0 &&
                !string.Equals(tsPair.TsRef.Trim(), entry.TsRef.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                result.TsRef = tsPair.TsRef.Trim();
                if (entry.TsRef.Trim().Length > 0)
                {
                    result.TsRefOverridden = true;
                    result.TsRefFromIo = entry.TsRef;
                }
            }
            return result;
        }

        /// <summary>
        /// fnMultiReplace parity: sequential replacement in dictionary order, empty Find rows
        /// skipped. Order matters — an earlier replacement's output feeds the later ones.
        /// </summary>
        public static string MultiReplace(string text, IReadOnlyList<SwpDictionaryEntry> dictionary)
        {
            string result = text ?? "";
            if (dictionary == null) return result;
            foreach (SwpDictionaryEntry entry in dictionary)
            {
                if (entry == null || string.IsNullOrEmpty(entry.Find)) continue;
                result = result.Replace(entry.Find, entry.Replace ?? "");
            }
            return result;
        }

    }
}
