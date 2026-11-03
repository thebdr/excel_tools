# Fresh comparison build checkpoint

Date: 10 September 2026. Status: **implementation started; feasibility and alignment proof remain in progress.**

The [approved plan](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md) incorporates all 21 adversarial findings. This checkpoint describes newly built code, not the abandoned comparison feature. No new ribbon entry, review workbook, resumable project, or merge workflow is claimed delivered.

## Built in this increment

All new implementation is isolated in `ExcelPowerTools.Core/SpreadsheetReview`. It does not call or derive from the old `Comparison` or `RevisionBridge` implementation. Existing dirty changes and old launch paths were preserved.

- Typed saved values retain their original encoding while numeric-equivalent encodings compare equally. Formulas, saved results, unknown formula text, and missing evidence remain separate.
- Sparse sheet observations retain blank-cell style lookup without materializing formatting-only rectangles. Last useful rows are measured from in-scope content, including formulas and failed observations needing review.
- An explicit complete column map is required by the alignment core. Outside columns cannot supply identity evidence. Inserted/deleted official columns can be represented independently on either side.
- A first ordered alignment candidate uses unique exact constant-content signatures and a library sequence diff. Formula strings do not become row identity keys. Repeated/ambiguous rows remain in raw review regions, with bounded partial-match suggestions, never silent first-match selection.
- Conflicting offsets are checked on overlapping consecutive row evidence, not unrelated regions. Explicit fixed-position comparison is available as a core mode. Local interval repair UI is still pending.
- Exact character differences preserve case, whitespace, line endings, Unicode text elements, and both original strings. Character strike is compared through text correspondence, not raw rich-text run boundaries.
- History comparison retains deleted content, formatting on cells that ever contained data, middle-only differences, presence states, and incomplete-channel diagnostics.
- A row identity ledger rejects contradictory chains atomically. Transitive connectivity does not become direct merge evidence.
- The saved-package reader reads XLSX/XLSM without starting Excel, executing formulas/macros, following external links, or saving inputs. It captures bytes while denying concurrent writers and records SHA-256, path, size, and modification time. The user may need to close an input already held for writing by Excel.
- Reader/engine diagnostics have unique issue IDs, artifact/region context, corrective messages, and source member/file/line information. Persistent application logging is still part of the upcoming workflow.

DiffPlex 1.9.0 is used only for exact sequence/text operations; this is not a final selection of a complete spreadsheet-alignment strategy. The restored package targets netstandard2.0/net45 and declares Apache-2.0. Its binary is unmodified. The license/notice is embedded in the add-in and copied to its loose and packed-output folders. See [the package](https://www.nuget.org/packages/DiffPlex/1.9.0) and [the exact upstream license](https://raw.githubusercontent.com/mmanela/diffplex/3cb6415c331a0af2f2686154cb77169ef29b63ce/License.txt).

## Verified build and tests

- Pre-change baseline: 1,542 passing tests, one environment-gated real-file test skipped.
- New focused corpus: 56 passing tests after the confirmed-Z update, one separately enabled real-file trial. Includes 50 deterministic insert/delete variants inside a metamorphic test and an explicit Z-inclusive/AA-excluded boundary regression.
- Debug add-in: both 32-bit and 64-bit packed XLLs built, with DiffPlex included. Build reported zero warnings and errors.
- Prior full suite, before the Z-specific test/configuration update: 1,597 passed, zero failed, two environment-gated real-file tests skipped, 35 seconds. The new real-file test was then run separately and passed on all five inputs. The focused corpus was rebuilt and rerun after the Z update; the full suite was not rerun and the add-in was not rebuilt for this test-only scope change.
- Initial real saved-file trial ran successfully on all five user-supplied files in the supplied review order. SHA-256 was checked before capture, against the captured snapshot, and after reading. All five originals remained unchanged. The first confirmed-A:Z rerun stopped because another process held R2.3 open. After the user closed it, the complete five-file A:Z rerun passed in 5.33 seconds with all five fingerprints still unchanged.
- No live Excel UI, native rendering, release installation, 200 MB workbook, or 20-revision target-scale validation was performed in this increment.

The spreadsheet skill's read-only preservation guidance was applied to the real artifacts. Validation used the new C# product reader itself so saved values/results and unsupported evidence could be tested without recalculating or exporting the inputs. No auxiliary workbook was authored.

## Real input discovery

The user supplied the sorter revisions under their `_IOlist` work folder, in this order. The test takes the root through `EXCEL_REVIEW_TRIAL_ROOT`; no personal absolute source path is hard-coded in the test.

| Revision file | Bytes | Sheets | Main sheet stored cells | Main sheet retained observations | Last main-sheet content |
| --- | ---: | ---: | ---: | ---: | --- |
| `old/8FVZ_IO0001_R1.0 - sorter.xlsx` | 3,080,673 | 3 | 1,134,964 | 21,335 | AA1620 |
| `old/8FVZ_IO0001_R1.1 - sorter.xlsx` | 3,141,887 | 4 | 1,152,115 | 22,653 | AA1637 |
| `old/8FVZ_IO0001_R1.2 - sorter.xlsx` | 357,232 | 6 | 40,869 | 23,069 | AC1638 |
| `8FVZ_IO0001_R1.3 - sorter.xlsx` | 407,245 | 6 | 50,069 | 31,422 | AC1998 |
| `_latest/8FVZ_IO0001_R2.3 - sorter.xlsx` | 415,981 | 6 | 50,419 | 32,706 | AC1998 |

“Retained observations” includes typed empty strings and formula evidence; it is not a count of substantive differences. The content envelope, not Excel's saved dimension/formatting tail, bounds row alignment. This independently reproduces the old-revision formatting-tail problem without triggering the abandoned dense-grid envelope.

Recorded source fingerprints:

```text
R1.0 1642d3ba5f5b03c9c6db767a8fc33a3db56116adad6d6ca82636b643b6360c30
R1.1 041715c4287e31f599bc9c7418283ce050ef1b15461aa8576e4a193432d0da2f
R1.2 e35761d687ae23f10070d524aca79838a7dab72006dab83795ea2423359daef5
R1.3 cbb40d632011d023eafd7d4231e1ef72472aebd87dff04246e61a3ce55b764b1
R2.3 6a4905608933b707f729bd5ca6f35fd70a3850668628837659732fbf49c0fae9
```

The R1.1 to R1.2 transition has an unpaired `Sheet1` on the before side and unpaired `ASI Addressing`, `Chutes`, and `_MakePlcTags` on the after side. They were not forced into inferred rename matches. `CabinetID_SWP` first appears in R1.1.

### Alignment trial interpretation

The initial trial provisionally compared every content column at its physical position. The user subsequently confirmed **A:Z inclusive** for `R1.S1.C1` across these five sorter revisions; AA onward must not influence matching, row extent, or differences. Other sheets' official caps and all actual column correspondences remain to be confirmed. The initial counts below are retained as historical discovery evidence, not the current official scope. All counts are diagnostic candidate coverage, **not validated logical row matches or complete difference totals**.

| Main-sheet pair | Provisional columns | Unique exact-constant anchors | Unresolved/one-sided regions |
| --- | --- | ---: | ---: |
| R1.0 to R1.1 | A:AA | 288 | 199 |
| R1.1 to R1.2 | A:AC | 616 | 91 |
| R1.2 to R1.3 | A:AC | 1,042 | 222 |
| R1.3 to R2.3 | A:AC | 598 | 231 |

Every scoped row is retained in an anchor or raw region. Partial suggestions are not included as accepted matches. R1.1 to R1.2 reaches the initial 250,000 candidate-pair search budget; the remaining regions remain explicit and are not labelled unchanged. The final five-file test command completed in 7.41 seconds, but these 0.36–3.14 MB inputs do not prove the target-scale envelope.

### Confirmed A:Z rerun

The compiled real-file trial now applies the confirmed cap only to the sorter main sheet, not as a product-wide default. The boundary regression confirms that data in Z participates in inserted-row alignment while contradictory AA data and distant AA-only tails do not affect it.

| Main-sheet pair | Official scope | Unique exact-constant anchors | Unresolved/one-sided regions |
| --- | --- | ---: | ---: |
| R1.0 to R1.1 | A:Z | 266 | 177 |
| R1.1 to R1.2 | A:Z | 616 | 91 |
| R1.2 to R1.3 | A:Z | 1,043 | 221 |
| R1.3 to R2.3 | A:Z | 598 | 231 |

All five captured fingerprints match those recorded above. Their last in-scope content rows remain 1,620 / 1,637 / 1,638 / 1,998 / 1,998. The R1.1 to R1.2 suggestion search still reaches its explicit candidate budget. The initial attempt stopped before capturing `_latest/8FVZ_IO0001_R2.3 - sorter.xlsx`: `System.IO.IOException`, file in use by another process, at `SavedWorkbookReaderTests.Fingerprint` while opening the source for hashing. The user closed that workbook and the complete rerun passed without weakening the no-concurrent-writer capture rule. No source file was saved or changed by either trial.

Task workflow confirmed by the user: when a file lock blocks this work, use a task-attached listener to recheck every three minutes, remain quiet while the lock is unchanged, and resume the pending work once safe capture succeeds. The native listener `retry-locked-comparison-files` was configured for that interval and paused after this successful validation. Reactivate it for a future lock blocker with the then-current artifact and pending action; do not run a completed trial indefinitely or force-close the user's applications. This is an agent-workflow instruction, not a claim that the add-in has gained a background retry feature.

The initial all-content trial found two important implementation flaws and supplied regressions:

1. Comparing each column's dominant offset over the whole sheet falsely diagnosed partial shifts when columns had unique values in different regions. Detection now requires conflicting witnesses on the same consecutive rows and only invalidates affected intervals.
2. Dropping shared-formula followers as whole-cell failures also discarded usable constant-row evidence. The reader now retains known formula presence, saved descriptor, and cache separately from unavailable expanded text. Independent constants can anchor the row; formula comparison remains explicitly incomplete.

## Known gaps and next work

These are required work, not optional polish:

1. The A:Z main-sheet read-only trial is complete. Confirm the other sheets' official caps and column pairing on the real set. Add a readable correspondence trial and obtain judgements for a small ambiguous sample. The current unique-exact-anchor coverage is insufficient for the intended everyday workflow.
2. Evaluate bounded contextual partial matching and repeated-block alternatives against held-out expected results, not just smaller difference counts. Formula expansion must not participate in a circular proof of row identity.
3. Expand shared formulas faithfully. Complete date-type/display context, theme-tint rounding, gradient/automatic colors, rich-text inheritance edge cases, array/spill semantics, and the reader capability inventory. Currently unsupported evidence is diagnosed; it is not declared unchanged. The style reader uses direct/named-style records and does not evaluate conditional formatting; these semantics need further independent Excel fixtures. See [Microsoft's style-record contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-3.0.1).
4. Trial bounds currently include 256 MiB captured source bytes, 2 GiB declared expanded package bytes, 512 MiB per package part, ten million stored cells per sheet, and 20,000 scoped alignment rows. They are explicit prototype resource guards, not final supported-product limits. Isolated-worker cancellation, hostile-package accounting, and worst-case memory remain feasibility gates.
5. Prove the self-contained review container, passive Excel navigation, exact long text, and checkpoint behavior before locking the report architecture.
6. Build the actual read-only ribbon workflow with setup, pairing/skip, column cap, progress, guide/tooltips, context/character review, and save/reopen. Existing comparison commands still launch the old implementation.
7. Continue the plan's multi-revision, XLSB/separated-text, project archive, and capability-gated merge milestones. No merge API or input-write path was introduced in this increment.

No milestone is marked complete by this checkpoint.
