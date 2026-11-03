# Spreadsheet comparison: alignment, Excel evidence, and reusable foundations

> Research snapshot dated 9 September 2026. The subsequent interview decisions and proposed delivery milestones are consolidated in [the fresh plan dated 10 September 2026](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md). That plan supersedes this report's provisional policies and open questions, particularly for the column boundary, empty-cell relevance, conditional formatting, uncertainty, revision order, formula relocation, and partial-row shifts. The research below remains evidence and analysis, not a benchmark or an implementation-completion record.

## Executive assessment

The central problem is not detecting unequal cells. It is establishing which cells still represent the same information after rows and columns have shifted. Existing products already address this with sequence alignment, selected identifiers, partial-row matching, and structural heuristics. A new comparison feature should evaluate these established approaches before committing to a custom engine.

Three conclusions matter most:

1. **A plausible row match is not a claim that the rows are equal.** Alignment may tolerate edits; the subsequent comparison must still report every in-scope difference.
2. **Excel has some revision information, but no universally usable history or permanent row identity can be assumed for arbitrary emailed copies.** Optional revision records can supplement evidence, not replace snapshot comparison.
3. **Manual fill and strikethrough are substantive data in this workflow.** They must remain distinguishable from comparison highlights, conditional formatting, and Excel's automatic stale-value strikethrough.

The most useful commercial references are Ablebits and xlCompare for partial matching, DiffEngineX for ordered alignment and reports, and Microsoft Spreadsheet Compare as a baseline. The most relevant open-source candidates are daff for table alignment and DiffPlex for reusable text/sequence-diff components. Neither is demonstrated here to satisfy the whole product requirement. The findings below distinguish documented behavior, source-code observations, and provisional recommendations.

## 1. Product requirements carried into this research

The target is a desktop Excel add-in for reviewing several independently emitted revisions of broadly similar workbooks. The first objective is identifying changes; selective merging into a new copy comes later.

| Area | Requirement |
| --- | --- |
| Revisions | Multiple copies of the same logical workbook; names and apparent version numbers are unreliable; order is suggested and user-correctable. |
| Sheet selection | Same-name sheets become automatic candidates, with a checkbox to skip. Unmatched sheets can be manually paired, including renamed sheets. No automatic rename decision. |
| Rows | All rows are in scope. Insertions, additions, and deletions must not turn the rest of the worksheet into false changes. A reliable key cannot be required. |
| Columns | Appended and inserted columns are possible. A last-column boundary excludes unofficial content to the right. Boundary behavior after insertion remains to be settled. |
| Compared information | Formulas, values, cell interior/fill, and strikethrough. Exact character-level differences are required. |
| Review | Group similar changes, expand individual details and revision history, show neighboring rows, preserve original visual evidence, and provide source links and full filesystem paths. |
| Interaction | Context menus and multi-selection; tooltips throughout new UI and a separate guide. |
| Persistence | Resumable comparison project, a self-contained Excel review, and a complete export archive of inputs, report, snapshots, decisions, and logs. |
| Safety | Inputs stay untouched. Selected older content is merged only into a newly named output based on the selected latest master, with an audit trail. |
| Diagnostics | Progress and learned duration estimates; detailed logs with stable issue identifiers, artifact context, source location, and corrective actions. |
| Formats | XLSX, XLSM, XLSB, CSV, TSV, and other separated-value text files. |

These requirements define the evaluation target. The existing implementation is not an architectural starting point. Proposed algorithms, thresholds, and UI decisions below are not yet an approved implementation plan.

## 2. Existing comparison products

### Microsoft Spreadsheet Compare / Inquire

Microsoft documents a two-workbook comparison of values, formulas, and formatting, with a side-by-side display, detailed results, and Excel export. Its current support documentation limits availability to qualifying Windows enterprise/equivalent Office editions. Inquire provides a launch point; it is not a separate row-matching algorithm. [1: Microsoft overview](https://support.microsoft.com/en-us/excel/overview-of-spreadsheet-compare), [2: Microsoft workflow](https://support.microsoft.com/en-us/excel/compare-two-versions-of-a-workbook-by-using-spreadsheet-compare)

A Microsoft Office 2013-era white paper explicitly describes automatic realignment for inserted/deleted rows and columns. That is evidence that structural alignment is part of the product family, not evidence of its current internal algorithm, thresholds, or success rate on any particular workbook. The algorithm was not disclosed in the reviewed documentation. [3: Microsoft Spreadsheet Controls white paper](https://download.microsoft.com/download/d/e/8/de8429fa-1159-4adb-8ace-5b07a0ebbeca/spreadsheetcontrols_whitepaper.pdf)

**Evaluation role:** a useful independent baseline. Do not confuse a two-file difference report with a portable, resumable multi-revision project or proof of who edited a cell.

### Commercial references most relevant to row matching

| Product | Documented approach | Relevance and limitations |
| --- | --- | --- |
| **Ablebits Compare Two Sheets** | Offers all-column, key-column, and fixed cell-by-cell modes. Partial-row policies include first match, best match, and full match only; fully matching rows have priority. Formula and formatting comparisons are optional. | Closest documented UX to the proposed downward search with partial matches. Row order matters. The guide excludes merged-cell sheets and warns about volatile functions. Its broad formatting claim does not separately establish rich-text strikethrough fidelity. [4: Ablebits guide](https://www.ablebits.com/docs/excel-compare-worksheets/) |
| **xlCompare** | General worksheets use order-sensitive best matching; table mode uses keys; outline mode considers row groups. Its rules include matching by an equal prefix or allowing a limited number of differing cells. It also supports manual pairing and background/font comparisons. | Particularly useful for understanding configurable matching policy. Its documented defaults are product choices, not universal thresholds. No public source code or independently verified accuracy was established. [5: xlCompare help, especially Compare As and Rules](https://xlcompare.com/help.html) |
| **DiffEngineX** | Describes itself as applying classic diff to spreadsheets. It aligns using selected columns, inserting gaps into copies; it does not reorder rows. Alignment is followed by cell comparison. | Clear reference for order-preserving alignment, contextual reports, navigation links, and character-level highlighting. Selected columns need useful distinguishing information; data may need sorting for its workflow. Sorting should not be imported into this product's default workflow. [6: DiffEngineX short guide](https://diffenginex.com/diffenginex-guide.html) |
| **Synkronizer** | Distinguishes ordinary-sheet, database/keyed, and one-to-one modes. Supports worksheet selection and ranges, formulas, formats, reports, and transferring differences. Database mode exposes duplicate-key and duplicate-record results. | Good reference for visible comparison modes and duplicate handling. Its general-sheet internal alignment algorithm is not sufficiently disclosed for reimplementation. [7: Synkronizer selection](https://www.synkronizer.com/usermanual/synkronizer-11/select.htm), [8: database mode](https://www.synkronizer.com/usermanual/synkronizer-11/select_worksheets_database.htm), [9: FAQ](https://www.synkronizer.com/compare-excel-files-faq) |
| **PerfectXL Compare** | Smart mode accounts for shifts caused by insertions/deletions; Basic mode compares fixed coordinates. Details can be expanded and exported to Excel. | Useful reference for a simple mode switch and progressive detail. The walkthrough does not disclose its alignment algorithm or establish manual fill/partial-strike fidelity. [10: PerfectXL walkthrough](https://www.perfectxl.com/resources/walkthroughs/perfectxl-compare/) |

DiffEngineX also documents an important distinction: aligning known similarities can leave two unmatched rows opposite one another, whereas its Alignment Plus option gives unmatched rows separate gaps. This illustrates the policy choice between “modified row” and “deletion plus insertion.” Neither presentation alone proves what editing operation actually occurred. [11: DiffEngineX alignment explanation](https://www.florencesoft.com/compare-excel-spreadsheets-news/45-excel-row-alignment-explanation.html)

**Assessment:** established products validate the need for more than coordinate comparison, but marketing terms such as “smart” and “best match” are not sufficient specifications. Broad “formatting supported” claims also need targeted tests for fill patterns, themes, and partial strikethrough. None of the reviewed documentation establishes the entire requested multi-revision workflow end to end.

## 3. Open-source options and what can actually be reused

| Project | Published license / technology | Useful component | Important boundary |
| --- | --- | --- | --- |
| **daff** | MIT; Haxe core with several language distributions | Table alignment, row/column differences, contextual tabular reports, patch/merge concepts; identifiers can be supplied but are not mandatory. | Table-oriented, not a ready-made Excel formatting/history engine. Native workbook semantics and the desktop integration need separate handling. [12: daff repository](https://github.com/paulfitz/daff), [13: license](https://github.com/paulfitz/daff/blob/master/LICENSE.md) |
| **na-ka-na/ExcelCompare** | MIT; Java using Apache POI and ODF tooling | Command-line comparison of XLS, XLSX, XLSM, and ODS contents; configurable exclusions. | Explicitly ignores formatting and macros. Its documented feature set does not establish the required partial-row alignment or XLSB support. [14: ExcelCompare repository](https://github.com/na-ka-na/ExcelCompare) |
| **rafal-dot/xlsxDiff** | AGPL-3.0; Python | An Excel-output review concept with cell text/formula and row/column changes, visually similar to tracked changes. | Explicitly ignores formatting; XLSX-focused. It is interesting as a report reference, not a demonstrated complete engine for this scope. Any code adoption needs license review. [15: xlsxDiff repository](https://github.com/rafal-dot/xlsxDiff) |
| **DiffPlex** | Apache-2.0; C# | Existing text-diff components, including character, line, and custom chunking. | Does not understand worksheets, column correspondence, formulas, fill, or row identity by itself. It is a component candidate rather than a spreadsheet comparator. [16: DiffPlex repository](https://github.com/mmanela/diffplex), [17: core API implementation](https://raw.githubusercontent.com/mmanela/diffplex/master/DiffPlex/Differ.cs) |

### What daff's source reveals

The inspected `CompareTable.hx` aligns columns before rows. Without supplied identifiers, it selects up to five relatively distinctive common columns and tries combinations as indexes. Unique cross-table matches establish links; neighboring already matched rows help resolve some duplicates. Its column matching searches near the top, and its general row path explicitly links row zero as a header. [18: daff alignment source](https://github.com/paulfitz/daff/blob/master/coopy/CompareTable.hx)

**Inference:** this is a genuine reuse candidate for broadly tabular inputs, and demonstrates useful content-derived matching without mandatory user keys. Its header assumptions and index heuristics need adversarial testing on title blocks, repeated sections, and broadly edited rows. It is not evidence that arbitrary fuzzy rows will always align correctly. The inspected source is a moving branch, not a pinned dependency selection.

### Reuse recommendation

Evaluate both a table-aligner route and an established sequence-diff route. Avoid committing to a language bridge or port before checking correctness, explainability, licensing, packaging, and cancellation behavior. Reuse a mature character-diff component rather than implementing character edit scripts anew. Keep Excel reading, exact comparison channels, alignment decisions, review persistence, and merge safety independent of the chosen matching component.

Performance, packaging, and accuracy on the target workbooks remain unverified. Published behavior and public source alone cannot establish them; this assessment is not a benchmark or a security audit.

## 4. What Excel stores about identity and changes

### Coordinates are not enduring row identity

The ordinary SpreadsheetML row has a row index; the cell table describes the current worksheet. A row index is a position, not a guaranteed identity retained across independent revisions. That conclusion is a design inference from the ordinary row model, not a claim that every optional Excel extension contains no identifiers. [19: Microsoft Row schema/API](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row?view=openxml-3.0.1)

Excel's worksheet `CodeName` can remain unchanged when the visible sheet name changes. It can therefore be useful supplementary evidence during manual sheet pairing. Microsoft also permits changing it at design time. It must not override the requirement for explicit pairing of differently named sheets, nor be treated as globally unique across unrelated workbooks. [20: Worksheet.CodeName](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.codename)

### Real revision records exist, but are conditional

SpreadsheetML defines shared-workbook revision logs containing cell changes, formatting changes, row/column insertions or deletions, moves, and sheet-name changes. These are genuine history structures, not merely filenames or style hints. Their existence in the specification does not mean they are present, complete, or retained in every ordinary workbook. [21: Microsoft Revisions model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.revisions?view=openxml-3.0.1)

Modern Excel files can also contain revision-related GUIDs and runtime metadata. The `uid` extension identifies its referencing object; it is not a promise that every row has one. Microsoft explicitly describes `revisionPtr` as runtime metadata to be ignored by other implementations, with fields such as `documentId` and `revIDLastSave` also documented as ignorable. These are not a supported general-purpose reconstruction API. [22: revision UID](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/05198417-f8f7-4ee4-9678-59599f962c90), [23: revisionPtr](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5082e3cf-b020-4599-b656-3b4207f0470d), [24: CT_RevisionPtr](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7fc8862a-63f2-4263-8334-b9e72b480c8e)

**Recommendation:** optionally inventory well-understood historical evidence and record its availability/provenance. The core comparison must work when it is absent. Never interpret absence of a revision record as absence of a change.

### Show Changes and Version History are different services

Show Changes can display recent edits with who/where/when and prior values, for up to 60 days. Its help page states that formatting changes are not shown, some earlier values can be missing, and saving a copy, uploading/replacing a file, or editing through unsupported versions can clear its visible history. The documentation also notes that clearing the pane does not necessarily remove every historical trace. [25: Show Changes](https://support.microsoft.com/en-us/office/show-changes-that-were-made-in-a-workbook-978ceea7-bbf6-4337-bca7-22e7cc9892e8), [26: Show Changes limitations](https://support.microsoft.com/en-us/excel/get-help-with-show-changes-in-excel)

Microsoft 365 Version History depends on storage in OneDrive or SharePoint. A standalone emailed workbook is not equivalent to the associated cloud version collection. [27: Version History](https://support.microsoft.com/en-us/office/collab-files/view-previous-versions-of-office-files)

**Conclusion for emitted copies:** compare preserved snapshots. History, authorship labels, filenames, filesystem timestamps, and manually applied colors can supply context, but cannot alone establish a complete edit chronology or identify the person responsible for each cell change.

## 5. Established algorithm families

### Exact sequence alignment: LCS / Myers

Treat the rows as an ordered sequence of comparable tokens. Longest-common-subsequence and shortest-edit-script methods find an alignment through matching elements, with gaps for additions and removals. Myers' 1986 algorithm is designed to be efficient when sequences are similar, with time expressed as O(ND), where N is total input length and D is edit-script size. [28: Myers paper](https://neil.fraser.name/writing/diff/myers.pdf)

**Implication:** excellent unchanged-row anchors are available without a business key. However, using the complete raw row as an exact token makes an edited row look like a removal plus an addition. A second matching stage or a scored sequence model is needed for partial matches. A shortest script is one economical explanation, not proof of historical identity.

### Patience / histogram-style anchors

These methods favor distinctive common elements rather than letting repeated boilerplate dominate the alignment. JGit's documented histogram implementation extends patience-style matching to low-frequency elements and falls back when candidate chains become excessive. [29: JGit HistogramDiff documentation](https://archive.eclipse.org/jgit/site/5.8.1.202007141445-r/org.eclipse.jgit/apidocs/org/eclipse/jgit/diff/HistogramDiff.html)

**Implication:** rare unchanged rows or sections can divide a worksheet into smaller, more meaningful gaps. Blank rows, repeated headings, zero-filled templates, and repeated formulas should not be strong anchors merely because they are numerous. The source cited is a historical implementation reference, not a recommendation to adopt that old package version.

### Weighted sequence alignment

Sequence alignment also has an established scored formulation: pairing similar elements competes with introducing gaps. The Needleman-Wunsch family and related work provide foundations; the underlying elements need not be biological sequences. [30: Needleman and Wunsch](https://pubmed.ncbi.nlm.nih.gov/5420325/), [31: Smith, Waterman, and Fitch](https://pubmed.ncbi.nlm.nih.gov/7334527/)

**Proposed application:** within bounded gaps between reliable anchors, score row-pair similarity and compare coherent sequences of matches, insertions, and deletions. This is more robust than committing to the first locally plausible match. Full unconstrained matrices can be costly; restricted windows must surface “search exhausted” instead of silently turning an unfound match into a proven deletion.

### Spreadsheet-specific structural inference: SheetDiff

The 2010 research prototype SheetDiff compares cells, identifies promising row/column structural changes, and evaluates whether insertions/deletions reduce the resulting difference set. It also presents changes in one worksheet context. Its paper explicitly discusses ambiguity between a heavily modified row and a structural change. [32: Chambers, Erwig, and Luckey, SheetDiff](https://web.engr.oregonstate.edu/~erwig/papers/SheetDiff_VLHCC10.pdf)

**Assessment:** valuable evidence that two-dimensional structure and grouping deserve explicit treatment. Its old product comparisons and experiment-selected threshold must not be treated as current market rankings or universal parameter choices. The academic prototype should not be confused with unrelated newer websites using the same name.

### Comparative assessment

| Approach | Natural strength | Main risk for these artifacts |
| --- | --- | --- |
| Fixed coordinates | Exact differences when structure is unchanged | One insertion produces a cascade of misleading changes. |
| First plausible downward match | Simple and easy to explain | One weak early match can consume the wrong row and misalign everything after it. |
| Exact ordered diff | Finds unchanged structure without keys | Edited rows need a second interpretation stage. |
| Distinctive anchors plus scored gaps | Uses order and local context while tolerating edits | Thresholds, duplicate handling, and fallback limits need real-artifact validation. |
| Content-derived table indexes | Can identify rows without user-provided keys | Header assumptions and unstable/duplicated content can mislead. |
| Structural change inference | Explicitly considers row and column operations | Minimizing differences can prefer a mathematically compact but historically wrong explanation. |

## 6. A candidate alignment policy for evaluation

The following is an analytical recommendation, not a selected implementation. It combines established methods while keeping uncertain identity separate from exact differences.

### Start with strong evidence; refine the uncertain gaps

1. **Confirm the comparison scope.** Establish sheet pairs and each revision's column boundary before computing similarity. Excluded columns must not secretly influence row matching.
2. **Establish likely column correspondence.** Use trustworthy headings when present, plus column content and order. A weakly detected header remains data. Offer manual correction where mappings are ambiguous.
3. **Find distinctive unchanged anchors.** Use values and formulas from the mapped in-scope columns; verify equality rather than trusting a fingerprint alone. Do not require a key column.
4. **Match edited rows between those anchors.** Consider partial similarity, neighboring rows, and order jointly. Consider gaps on both sides so that both additions and removals work.
5. **Leave competing explanations visible.** A row with two similarly plausible peers is not a confirmed identity. Offer alternatives or an unresolved region.
6. **Compare confirmed counterparts exactly.** Report formulas, values, fills, and strikethrough independently, including character-level detail where applicable.

Column and row mapping can depend on each other. Where headings are absent or changed, a bounded refinement pass can use strong row evidence to improve column correspondence and vice versa. It must terminate deterministically and flag unresolved structure; “repeat until it looks right” is not an acceptable policy.

### Why the first partial match can be wrong

Consider this invented, deliberately small example. Row numbers are coordinates, not identifiers:

| Old row | Old description | New row | New description |
| --- | --- | --- | --- |
| 20 | Section A | 20 | Section A |
| 21 | Motor ready / input / 0 | 21 | Motor ready spare / input / 0 |
| 22 | Motor running / input / 0 | 22 | Motor ready / input / 1 |
| 23 | Section B | 23 | Motor running / input / 0 |
| — | — | 24 | Section B |

A first-match rule may pair old row 21 with the new spare row 21. A context-aware interpretation can instead suggest that the spare row was added, old row 21 changed its value and moved to 22, and old row 22 moved to 23 unchanged. The unchanged section boundaries help constrain the alternatives. Yet this remains an inference unless more evidence distinguishes the rows.

Similarity must measure informative agreement, not just a percentage of equal cells. If 190 of 200 columns are blank, two unrelated rows must not become a “95% confident match.” Repeated zeros, common labels, and copied formulas also carry less distinguishing information than a rare description or a distinctive combination of cells. Score calibration belongs in validation; a similarity score should not be labeled a probability without evidence.

Fill and strikethrough should normally be compared after content alignment, not required to match as identity anchors. Editors intentionally change them, so treating them as strong identifiers would penalize precisely the rows the tool is meant to find. Original formatting may be displayed as supporting context without giving it decisive matching authority.

### Irreducible ambiguity

If two old rows are identical in every available signal and only one remains, snapshots cannot establish which original row was removed. The report can truthfully say that one occurrence disappeared, preserve both possible histories, and request a decision if an operation depends on a specific occurrence. Picking the first row silently produces apparent precision without evidence.

Likewise, row reordering is not the same as row insertion. Ordered matching is appropriate for mainly stable layouts, but arbitrary sorting or moved blocks require a separate movement policy. Do not reorder input data to make the algorithm's job easier; that destroys useful positional context and can change workbook behavior.

## 7. Excel-specific comparison semantics

### Values, formulas, and saved calculated results

SpreadsheetML stores formula text separately from its cached last-calculated value. Therefore “value changed” can mean an edited constant, a different saved formula result, or a result observed only after recalculation. These are different kinds of evidence. [33: Microsoft, Working with formulas](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas)

**Recommended comparison channels:** preserve the exact typed constant, logical formula text, and available saved calculated result separately. A formula change remains visible even if the result is the same; a result-only change remains visible but distinguishable. Do not recalculate the original as part of obtaining historical evidence. Missing caches must be labeled unavailable, not interpreted as empty or zero.

Inserted rows or columns can legitimately change addresses inside formulas. For example, an expression at one revision's row 20 might refer to that row, then move to row 21 with adjusted references. A useful report should retain the exact before/after expressions and may additionally classify a verified mapping-induced change as “reference adjusted with structure.” Simple textual replacement of row numbers is insufficient: absolute, relative, cross-sheet, named, table, and range references differ.

Shared and array formulas also require logical interpretation rather than a raw XML-text comparison. Microsoft's formula model distinguishes these representations. A compressed shared-formula representation changing to an expanded representation is not necessarily a change in the formula a person sees. [34: Microsoft CellFormula model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1)

### Interior/fill

A fill color can be represented by an explicit color, an indexed palette entry, or a theme color with a tint; patterned fills have foreground and background components. Cell styles refer to workbook-local font and fill records. Comparing raw style IDs across workbooks can therefore produce false differences or miss real ones. [35: Microsoft ForegroundColor model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1), [36: Microsoft CellFormat model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-3.0.1)

**Recommendation:** resolve the saved fill attributes, retain the representation as evidence, and distinguish no fill from an explicit fill. Decide separately whether visually identical colors with different theme/index encodings are substantive changes or quiet metadata details.

Direct formatting and displayed formatting also differ. Microsoft's `DisplayFormat` example shows a cell whose direct interior is white while conditional formatting makes its displayed interior red. [37: Microsoft Range.DisplayFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.displayformat)

For a workflow in which people manually color edited cells, saved manual formatting is the most natural initial evidence channel. Conditional-format rules and their evaluated display need an explicit policy; neither should be silently substituted for the other. A rule-driven color change is not evidence that a person deliberately recolored a cell.

### Strikethrough

SpreadsheetML supports strikethrough in both a cell font and rich-text run properties. Comparing only a whole-cell Boolean can miss a strike applied to part of the text. A review should be able to show exactly which text span changed its strike state. [38: Microsoft Strike model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.strike?view=openxml-3.0.1)

Recent Microsoft 365 Excel versions also use strikethrough to indicate stale formula results in manual/partial calculation scenarios. Microsoft states that the indication disappears after recalculation. This must not be confused with intentional saved text formatting. The support page establishes the UI behavior, not a complete file-level detection contract. [39: Microsoft Stale Value Formatting](https://support.microsoft.com/en-US/Excel/stale-value-formatting)

**Recommendation:** capture explicit font/run strikethrough as the substantive comparison channel. Label any separately observed stale-value display state. Validate the distinction in the actual Excel versions used; do not infer it from a screenshot alone.

### File formats

XLSX and XLSM are XML-based workbook formats, while XLSB uses BIFF12 binary records. Excel also documents that text formats lose workbook formatting. Thus CSV/TSV can participate in content comparison, but cannot supply original fill or strikethrough evidence. XLSB support must be verified independently rather than inferred from XLSX support. [40: Microsoft supported file formats](https://support.microsoft.com/en-us/excel/file-formats-that-are-supported-in-excel)

**Acceptance implications:** preserve text-file parsing settings and exact strings, including leading zeros, spaces, delimiters inside quotes, and line breaks. Do not reinterpret text starting with `=` as an executable Excel formula merely because it is copied into the review workbook. Reading macro-enabled files must not require running their macros or refreshing external data.

## 8. Scope boundaries and performance

### Official columns

The last-column control needs to show the actual column letter and, where available, its heading in every revision. A fixed letter and a logical official boundary are not always equivalent: if the official region ends at AD and a column is inserted before AD, the former last official column moves to AE.

A provisional UX is a shared suggested boundary with visible per-revision corrections. If the boundary follows matched columns, any adjustment should be shown for confirmation. A fixed physical cap remains useful when that is exactly what the reviewer intends. This is an open decision, not permission to include unofficial columns automatically.

Skipped sheets and excluded columns should remain listed in the project scope report. “Not compared” is not “unchanged.” A genuinely deleted sheet and a newly added sheet must be representable independently, without forcing a rename pairing.

### All rows and formatting-only tails

All rows with in-scope content must be examined, including hidden/filtered rows. Internal blank rows can carry structural context and should not simply be removed before alignment. There is no requirement to materialize Excel's entire grid in memory.

However, the combination of “all rows,” meaningful fill changes, and known excessive formatting beyond the useful workspace creates a real policy question. A content-based last row can omit a deliberately colored blank cell outside that boundary; a formatting-based last row can include hundreds of thousands of irrelevant cells. No algorithm can infer editorial intent from formatting alone.

**Recommendation for discussion:** show the last data/formula row and separately report detected formatting-only extensions, with an explicit decision about including them. Never silently claim that excluded tails were compared. Preserve in-scope formatting on blank cells inside the useful region.

### Scale

The stated planning envelopes are:

| Dimension | Typical | Worst target |
| --- | ---: | ---: |
| Revisions | 5 | 20 |
| Workbook size | 10 MB | 200 MB |
| Sheets per workbook | 5 | 30 |
| Rows per sheet | 1,000 | 5,000 |
| Columns per sheet | 30 | 200 |

Multiplying the worst geometry gives 600 million potential cell positions across the project; this is an upper geometric envelope, not an observed number of populated cells. One unrestricted 5,000-by-5,000 row-matching matrix already has 25 million candidate row pairs, before comparing their columns. That motivates sparse reading, distinctive anchors, bounded refinement, and processing a limited number of worksheet/revision pairs at once.

Do not promise a runtime from file size alone. Saved artifacts can contain unrelated objects or extensive formatting. Estimates should use previously observed extraction, alignment, and report-generation times with visible uncertainty; cached timings should not contain workbook data. Cancellation must preserve inputs and leave a clearly incomplete output rather than an apparently finished review.

## 9. Validation before engine selection

A fair evaluation needs known intended alignments, not just a count of reported changes. It should include anonymized representative artifacts and deliberately constructed adversarial cases. Real inputs remain read-only, and any comparator that normally edits open workbooks should receive disposable, explicitly identified copies.

| Case | Required evidence |
| --- | --- |
| Unchanged workbook | No substantive differences despite package/style-table reorganization. |
| Single and multi-row insertions | Existing rows retain their counterparts; no cascading false edits. |
| Deletion followed by insertion | Both directions work in the same region. |
| Insertion next to a partially edited row | Correct pairing or an explicit ambiguity, not a silent first-match guess. |
| Duplicate rows / repeated blocks | Ambiguity retained when evidence cannot identify the occurrence. |
| Mostly blank or zero-filled rows | Low-information equality does not create false confidence. |
| Column inserted inside the official area | Column mapping and the boundary remain visible and correct. |
| Unofficial columns appended | Their changes do not affect differences or row alignment. |
| Renamed versus deleted/new sheets | Explicit pairing and genuine non-pairing are both possible. |
| Title blocks, repeated headings, weak headers | No mandatory first-row header and no silent data loss. |
| Value, formula, and cache-only edits | Separate exact before/after evidence for each channel. |
| Formula address shifts | Exact expressions retained; structural classification only when verified. |
| Fill-only changes | Direct colors, theme/tint, no-fill, and patterns handled by the agreed policy. |
| Whole-cell and partial-text strike | Exact strike changes preserved; no confusion with stale-value display. |
| Conditional formatting | Direct versus displayed appearance follows the selected policy. |
| Hidden rows and formatting tails | Inclusion/exclusion reported accurately. |
| XLSB and separated text | Required evidence survives each reader; unsupported evidence is not invented. |
| Large, sparse, and repetitive worksheets | Bounded resources, cancellation, progress, and complete coverage accounting. |
| Reopen / changed source / additional revision | Saved decisions are tied to the correct artifacts and mapping version. |

Measure four distinct outcomes: row/column matching correctness, exact-difference completeness, ambiguity correctly exposed, and resource use. A tool producing fewer differences is not automatically better; it may have hidden changes or incorrectly matched rows.

**Proposed comparison experiment:** evaluate a daff-based alignment, an established ordered-diff foundation with a partial-match refinement, and at least one commercial baseline. Choose the engine only after inspecting the difficult cases. Keep the experiment bounded to extraction/alignment and a readable evidence report; do not build the complete review and merging UI around an unvalidated mapping policy.

For future merges, unresolved identity must block only the affected operations. A provisional comparison is useful for inspection; it is not sufficient authority to overwrite a destination cell. Rechecking source fingerprints and destination mappings remains necessary even after a person approves a match.

## 10. Decisions for the next interview

1. **Row movement:** how often are rows sorted, relocated individually, or moved as blocks, rather than just inserted/deleted? This determines whether order-preserving alignment can be the default.
2. **Column boundary:** should the cap follow the last official column through insertions, with confirmation, or remain a physical letter for each revision?
3. **Manual versus displayed formatting:** is the substantive requirement saved manual fill/strike, with conditional/stale display treated separately?
4. **Formatting outside data:** how should colored or struck-through blank cells beyond the last data row be handled?
5. **Ambiguous matching:** what is the preferred review interaction for choosing among plausible row peers or leaving a region unresolved?
6. **Formula shifts:** should a formula that changes only because references follow inserted rows be grouped quietly while preserving exact character detail?

The provisional recommendation is order-preserving, content-based alignment with visible uncertainty; explicit column boundaries; exact comparison after matching; and saved manual fill/strike as separate evidence channels. Row movement and formatting-tail policies must be confirmed before that recommendation becomes a plan.

## Sources and numbered notes

Online sources were checked on 9 September 2026. Undated product documentation is identified as such. Historical publications describe established methods, not current product performance. The Microsoft white paper's alignment claim was available in indexed text; its download endpoint did not render directly. Commercial algorithms remain opaque where stated, and no vendor's advertised accuracy is treated as independently validated.

1. Microsoft. [Overview of Spreadsheet Compare](https://support.microsoft.com/en-us/excel/overview-of-spreadsheet-compare). Undated current support page; product scope and availability.
2. Microsoft. [Compare two versions of a workbook by using Spreadsheet Compare](https://support.microsoft.com/en-us/excel/compare-two-versions-of-a-workbook-by-using-spreadsheet-compare). Undated current support page; Windows editions, review, export, and source-color display.
3. Microsoft. [Spreadsheet Controls white paper](https://download.microsoft.com/download/d/e/8/de8429fa-1159-4adb-8ace-5b07a0ebbeca/spreadsheetcontrols_whitepaper.pdf). Office 2013-era publication; historical explicit claim of structural realignment.
4. Ablebits. [How to Compare Two Sheets in Excel](https://www.ablebits.com/docs/excel-compare-worksheets/). Undated guide; partial matching modes and documented limitations.
5. Spreadsheet Tools. [xlCompare Help Library](https://xlcompare.com/help.html). Page updated 24 May 2026; comparison strategies, Rules, sheet pairing, and formatting settings.
6. DiffEngineX LLC. [The Short Guide to DiffEngineX](https://diffenginex.com/diffenginex-guide.html). Undated guide; diff-style alignment, order restriction, reports, and character-level highlighting.
7. XL Consulting. [Synkronizer 11: Select](https://www.synkronizer.com/usermanual/synkronizer-11/select.htm). Undated manual; selection, comparison modes, projects, and ranges.
8. XL Consulting. [Synkronizer 11: Databases](https://www.synkronizer.com/usermanual/synkronizer-11/select_worksheets_database.htm). Undated manual; keys and duplicate classifications.
9. XL Consulting. [Synkronizer FAQ](https://www.synkronizer.com/compare-excel-files-faq). Undated current page; comparison categories and report/update capabilities.
10. Infotron / PerfectXL. [PerfectXL Compare walkthrough](https://www.perfectxl.com/resources/walkthroughs/perfectxl-compare/). Undated walkthrough; Smart/Basic modes and export.
11. Florencesoft / DiffEngineX. [Row Alignment Explanation](https://www.florencesoft.com/compare-excel-spreadsheets-news/45-excel-row-alignment-explanation.html). 17 December 2012; alignment versus explicit unmatched-row gaps.
12. Paul Fitzpatrick and contributors. [daff repository](https://github.com/paulfitz/daff). Current retrieved README; table comparison and distribution scope.
13. Paul Fitzpatrick and contributors. [daff license](https://github.com/paulfitz/daff/blob/master/LICENSE.md). MIT license in retrieved repository.
14. na-ka-na and contributors. [ExcelCompare repository](https://github.com/na-ka-na/ExcelCompare). Current retrieved README; formats, Java dependencies, MIT license, and formatting exclusion.
15. rafal-dot. [xlsxDiff repository](https://github.com/rafal-dot/xlsxDiff). Current retrieved README; report purpose, formatting exclusion, and AGPL-3.0 license.
16. Michael Manela and contributors. [DiffPlex repository](https://github.com/mmanela/diffplex). Current retrieved README; C# text-diff interfaces and Apache-2.0 license.
17. DiffPlex contributors. [Differ.cs](https://raw.githubusercontent.com/mmanela/diffplex/master/DiffPlex/Differ.cs). Retrieved moving-branch source; character/line/custom diff API.
18. daff contributors. [CompareTable.hx](https://github.com/paulfitz/daff/blob/master/coopy/CompareTable.hx). Retrieved moving-branch source; inferred-index and header assumptions.
19. Microsoft. [Row Class, SpreadsheetML](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row?view=openxml-3.0.1). Open XML API/schema reference; current row index and attributes.
20. Microsoft. [Worksheet.CodeName property](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.codename). Updated 12 September 2021; distinction from visible sheet name.
21. Microsoft. [Revisions Class, SpreadsheetML](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.revisions?view=openxml-3.0.1). Open XML API/schema reference; shared-workbook history structures.
22. Microsoft. [MS-XLSX: uid, 2014/revision namespace](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/05198417-f8f7-4ee4-9678-59599f962c90). Updated 14 February 2019; object-associated revision UID.
23. Microsoft. [MS-XLSX: revisionPtr](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5082e3cf-b020-4599-b656-3b4207f0470d). Updated 8 January 2026; runtime-only metadata restrictions.
24. Microsoft. [MS-XLSX: CT_RevisionPtr](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/7fc8862a-63f2-4263-8334-b9e72b480c8e). Updated 16 June 2023; revision pointer fields.
25. Microsoft. [Show changes that were made in a workbook](https://support.microsoft.com/en-us/office/show-changes-that-were-made-in-a-workbook-978ceea7-bbf6-4337-bca7-22e7cc9892e8). Undated current support page; 60-day recent-change view.
26. Microsoft. [Get help with Show Changes in Excel](https://support.microsoft.com/en-us/excel/get-help-with-show-changes-in-excel). Undated current support page; missing formatting, history clearing, and incomplete previous values.
27. Microsoft. [View previous versions of Office files](https://support.microsoft.com/en-us/office/collab-files/view-previous-versions-of-office-files). Undated current support page; OneDrive/SharePoint requirement.
28. Eugene W. Myers. [An O(ND) Difference Algorithm and Its Variations](https://neil.fraser.name/writing/diff/myers.pdf). Algorithmica, 1986; original paper, hosted copy. Exact sequence alignment and complexity.
29. Eclipse JGit. [HistogramDiff API documentation](https://archive.eclipse.org/jgit/site/5.8.1.202007141445-r/org.eclipse.jgit/apidocs/org/eclipse/jgit/diff/HistogramDiff.html). Version 5.8.1, 2020; historical primary implementation documentation for low-frequency anchors and fallback.
30. Saul B. Needleman and Christian D. Wunsch. [A general method applicable to the search for similarities in the amino acid sequence of two proteins](https://pubmed.ncbi.nlm.nih.gov/5420325/). Journal of Molecular Biology 48(3), 443-453, March 1970; original sequence-alignment publication record, DOI 10.1016/0022-2836(70)90057-4.
31. Temple F. Smith, Michael S. Waterman, and Walter M. Fitch. [Comparative biosequence metrics](https://pubmed.ncbi.nlm.nih.gov/7334527/). Journal of Molecular Evolution 18(1), 38-46, 1981; original research abstract on scored sequence-alignment equivalence and generality.
32. Chris Chambers, Martin Erwig, and Markus Luckey. [SheetDiff: A Tool for Identifying Changes in Spreadsheets](https://web.engr.oregonstate.edu/~erwig/papers/SheetDiff_VLHCC10.pdf). VL/HCC 2010; author-hosted paper, especially section 4 on structural inference.
33. Microsoft. [Working with formulas](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas). Open XML documentation; formula text and cached calculated values.
34. Microsoft. [CellFormula Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1). Open XML API/schema reference; shared, normal, array, and data-table formula types.
35. Microsoft. [ForegroundColor Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1). Open XML API/schema reference; pattern colors, indexing, theme, and tint.
36. Microsoft. [CellFormat Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-3.0.1). Open XML API/schema reference; workbook-local fill/font/style record references.
37. Microsoft. [Range.DisplayFormat property](https://learn.microsoft.com/en-us/office/vba/api/excel.range.displayformat). Updated 13 September 2021; direct versus conditional display formatting.
38. Microsoft. [Strike Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.strike?view=openxml-3.0.1). Open XML API/schema reference; font-level and rich-text-run strikethrough.
39. Microsoft. [Stale Value Formatting](https://support.microsoft.com/en-US/Excel/stale-value-formatting). Undated Microsoft 365 support page; automatic strikethrough for stale calculation results.
40. Microsoft. [File formats that are supported in Excel](https://support.microsoft.com/en-us/excel/file-formats-that-are-supported-in-excel). Undated current support page; XML/binary formats and loss of formatting in text formats.
