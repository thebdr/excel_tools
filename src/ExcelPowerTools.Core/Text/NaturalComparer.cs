using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.Text
{
    /// <summary>
    /// Numeric-aware, case-insensitive string comparison: DI2 &lt; DI10, Slot_2 &lt; Slot_10. Digit runs
    /// compare by value (leading zeros ignored, with a deterministic tie-break on run length), everything
    /// else ordinally case-insensitive. Empty/null sorts first.
    /// </summary>
    public sealed class NaturalComparer : IComparer<string>
    {
        public static readonly NaturalComparer Instance = new NaturalComparer();

        public int Compare(string a, string b)
        {
            a = a ?? "";
            b = b ?? "";

            int i = 0, j = 0;
            while (i < a.Length && j < b.Length)
            {
                bool da = char.IsDigit(a[i]);
                bool db = char.IsDigit(b[j]);

                if (da && db)
                {
                    int si = i, sj = j;
                    while (i < a.Length && char.IsDigit(a[i])) i++;
                    while (j < b.Length && char.IsDigit(b[j])) j++;

                    // strip leading zeros (keep at least one digit)
                    int ai = si; while (ai < i - 1 && a[ai] == '0') ai++;
                    int bj = sj; while (bj < j - 1 && b[bj] == '0') bj++;

                    int la = i - ai, lb = j - bj;
                    if (la != lb) return la - lb;                    // more significant digits = bigger
                    for (int k = 0; k < la; k++)
                    {
                        int d = a[ai + k] - b[bj + k];
                        if (d != 0) return d;
                    }

                    // numerically equal (07 vs 7): longer raw run (more padding) sorts first, deterministically
                    if (i - si != j - sj) return (j - sj) - (i - si);
                }
                else if (da != db)
                {
                    return da ? -1 : 1;                              // numbers sort before text
                }
                else
                {
                    char ua = char.ToUpperInvariant(a[i]);
                    char ub = char.ToUpperInvariant(b[j]);
                    if (ua != ub) return ua - ub;
                    i++; j++;
                }
            }

            return (a.Length - i) - (b.Length - j);                  // shorter prefix-equal string first
        }
    }
}
