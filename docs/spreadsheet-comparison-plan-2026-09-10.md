# Spreadsheet comparison: fresh product and delivery plan

Date: 10 September 2026.

Status: **approved staged plan; implementation authorized on 10 September 2026.** The adversarial-review contracts below are adopted. Feasibility/alignment code and real-file trials are in progress; no product milestone is claimed complete. See [the build checkpoint](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-build-2026-09-10.md) for verified work and remaining gaps.

This plan consolidates the original user requirements and the fresh interview. It does not use the abandoned comparison implementation, its architecture, or its completion claims as a starting point. No existing implementation is deleted by this document. Unrelated add-in functionality and user-owned work must remain intact.

“Confirmed” means a user requirement or an accepted interview decision. “Proposed” identifies a design or delivery choice still subject to review and evidence. The research background is in [the 9 September research report](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-research-2026-09-09.md); its provisional questions are superseded by the decisions below.

The user authorized incorporating the [adversarial review](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-adversarial-review-2026-09-10.md) and starting implementation. Its safety and interaction contracts are integrated here; its alternative layouts, algorithms, and reusable document-family hints remain candidates to evaluate, not untested dependency selections.

## 1. Purpose and success condition

Help a reviewer understand differences between revisions of standardized spreadsheets, choose the most plausible information, and selectively recover it into a new revision.

Finding and explaining changes is the primary goal. Merging is the second goal. A usable delivery is an accessible Excel workflow with trustworthy results on the actual artifacts, not merely an engine, a setup dialog, or passing synthetic tests.

The files may have been emailed, renamed, overwritten by another person's similarly named copy, edited online, or downloaded and edited independently. Two different files can both be labelled “1.1”; a later-labelled file can be missing useful information. The user can establish the desired revision order confidently, but that order is not necessarily chronological or a parent-child relationship.

Consequently:

- The user-approved order controls review presentation. It does not prove ancestry, edit time, authorship, or intent.
- Every supplied artifact remains identifiable independently of its filename or revision label.
- Differences show what is present, absent, or different across snapshots. They do not claim that an absence was intentional, that a value was “reverted,” or that a particular person caused a change without evidence.
- There is no automatic “newest wins” resolution. The selected master provides the output's starting state; the reviewer chooses changes to it.
- Sources are never overwritten. The result is a newly named revision with an audit of the applied selections.

## 2. Confirmed comparison contract

### Revisions and files

- Accept XLSX, XLSM, XLSB, CSV, TSV, and other separated-value TXT files.
- Suggest an order from available names and metadata, but make it easy to correct. Inconsistent labels such as `r1`, `r02`, `v2.1`, `rev3`, or a suffix `31` cannot decide order conclusively.
- Keep distinct copies with the same filename or version label separate. Show full filesystem paths, revision labels, available timestamps, and the user-approved order.
- Let the user select the latest reviewed/master artifact explicitly. File timestamps and numerical sorting do not select it irreversibly.
- Preserve supplied descriptions as context, without treating them or manual cell colors as reliable change logs.
- Add further revisions to a saved project and produce a new export. Previously reviewed evidence and decisions must remain traceable.

### Sheets, headers, and official columns

- Same-name sheets are automatic comparison candidates. Each has a visible checkbox to skip it.
- Differently named sheets may be paired manually. Similarity or a matching internal identifier can be supporting information, never an automatic rename decision.
- A missing sheet and an independently added sheet can remain unpaired. Do not force the user to invent a counterpart. Skipping differences for an unwanted sheet is available and recorded.
- Provide a last-official-column boundary. Unofficial content to its right is excluded from both comparison and row-matching evidence. For the five supplied sorter revisions, the user confirmed **Z inclusive (A:Z)** on `R1.S1.C1`; this is an artifact-specific scope, not a global default.
- The boundary follows the last official column when insertions move it, for example AD to AE. Show an editable limit for each revision, including the heading when available. Ambiguous boundary mapping needs review rather than silent expansion.
- Boundary discovery and comparison are separate stages. Discovery may propose a moved cap, with visible evidence; only the resolved per-revision scope can supply matching evidence. An unofficial duplicate header cannot silently enlarge the boundary. Changing excluded content after scope resolution must not change matches or differences.
- Compare all rows with in-scope content, including hidden or filtered rows. There is no arbitrary last-row cutoff.
- Use trustworthy headers as evidence, following the useful behavior of the add-in's header identifier. Weak header detection means “treat as data,” not “discard this row.” No mandatory unique row key and no assumption that row one is a header.

### Structural matching

- The common case is whole-row insertion or deletion in otherwise broadly stable layouts. Align corresponding rows so those operations do not create a cascade of false cell changes below them.
- Support inserted and appended columns within the official region as well as the boundary corrections above.
- Default row alignment preserves order. It can use partial content matches and neighboring rows, not only exact whole-row equality.
- Partial-row cell insertions/deletions may happen, but **for this scope they are ordinary differing data**. Do not independently realign side-by-side tables or reconstruct local cell-shift operations. Headers do not expand this scope.
- If informative columns support contradictory row offsets, do not silently accept a whole-row shift. Keep the region unresolved and offer a local fixed-position constraint; this is an ambiguity repair, not a new independent-table alignment mode.
- Arbitrary sorting and moved-block recognition are not required as a dedicated first-release matching mode. Such changes must remain visible as differences or unresolved correspondence, not disappear.
- Ambiguous correspondence does not abort the project. Continue unaffected regions, show possible peers with neighboring rows, and allow manual pairing through the context menu.
- Require resolved or explicitly confirmed correspondence only for merges that depend on it. Never present a guessed history as a confirmed match.
- Pairwise correspondence is partial and one-to-one. One logical history cannot contain two observations from the same revision. Conflicting paths across revisions remain alternatives; adding a bridging revision cannot silently fuse existing identities. Reordering changes presentation, not raw evidence or silently accepted identity. Any changed proposals from deliberate reanalysis require visible revalidation.
- Keep entities absent from the reference/master in the evidence set. Direct, verified target correspondence is required for merging; a chain of plausible partial matches is not enough.

### Which information is compared

| Channel | Required behavior |
| --- | --- |
| Stored values | Preserve exact content and type; show additions, removals, and changes. Do not hide case, whitespace, or numeric-looking text differences. |
| Formulas | Compare the formula expressions, independently of whether their calculated results happen to agree. Retain exact before/after text. |
| Saved calculated results | Show available result-only changes separately and at lower emphasis. Missing saved results mean unavailable, not zero or empty. No recalculation is required to obtain historical evidence. |
| Cell interior/fill | Compare the required manual formatting evidence, separately from content. Preserve the original appearance in review context. |
| Strikethrough | Compare manual strike, including a change affecting only part of a text value. |
| Character detail | Show character-by-character changes for text and formula expressions; make spaces and line breaks understandable when relevant. |

Conditional-formatting results are not a separate requirement. Read manual fill and strike through the chosen reader's normal properties; do not build a conditional-format evaluation engine. If a reader's ordinary properties already include dynamic appearance, that is acceptable, but its behavior must be identified and tested rather than portrayed as proof of a manual edit. Other formatting and workbook-object differences are outside the substantive comparison scope.

Every in-scope difference remains inspectable. Content is visually foremost; formatting, result-only changes, and less significant character differences can be quieter or grouped, not suppressed.

Define canonical typed values independently of storage spelling and display context: numeric 1 differs from text "1", while numeric encodings 1 and 1.0 do not create a character edit. Retain original encodings as provenance. Text typing follows the explicit parsing policy. Record the workbook date system and relevant display context; disclose interpretation changes without adding general number-format comparison or silently normalizing away numeric differences.

Resolve effective per-character strikethrough independently of rich-text run boundaries. Equivalent run repartitioning is not a formatting change. When inserted characters move existing struck text, explain formatting through the character correspondence rather than inventing strike toggles from shifted offsets; retain exact snapshots if correspondence is ambiguous.

### Empty-cell relevance across the complete revision set

Evaluate relevance after correspondence has been established, using every supplied revision rather than just the first and last:

- If a logical cell is empty in every revision, it produces no difference output, even if formatting differs.
- If any revision contains a value or formula there, it remains relevant in revisions where that content is absent. This retains deleted content and any relevant formatting evidence on that cell.
- A formula returning an empty-looking result still counts as content. Zero, false, and error values are not blank.
- A value introduced in a middle revision and missing at both endpoints must still appear in expandable history.
- Formatting-only tails must not expand the useful data region into a huge false grid. Show the detected data boundary and explain the applied relevance rule in the coverage summary.
- Unresolved matching is not permission to discard a region's content. Keep its evidence visible until identity is resolved or explicitly left unresolved.

The evidence model distinguishes present content, present blank, absent row/column, excluded scope, unavailable channel, failed extraction, and unresolved correspondence. Only an observed blank is blank. Determine content relevance across logical positions, then retain or retrieve fill/strike for their matched blank observations, including inherited formatting. A sparse reader cannot discard relevant post-deletion formatting just because that revision no longer has a stored value there.

### Formula changes caused by structure

When verified row, column, sheet, and reference mappings establish that the formula still refers to the same logical inputs, retain the exact expression difference but group it as **“References adjusted by structural changes.”**

This is not a general claim of mathematical equivalence. A changed range might include different inputs even when its endpoints look similar. If reference correspondence or formula semantics cannot be verified, do not classify the change as harmless. Literal string substitution is insufficient.

When recovering an older formula into the new revision, follow verified logical references in the new layout rather than blindly preserving old addresses. Show the adjusted expression before applying it. Uncertain references remain pending for review.

Maintain an explicit formula-verification capability matrix with verified, unresolved, and unsupported outcomes. Coordinate-sensitive expressions, indirect references, names, unavailable dependencies, and changed range membership cannot be certified merely from changed addresses or matching saved results. Normalizing an expression using a guessed mapping is not independent proof of that mapping. Known supported direct-reference cases must work; declaring every formula unsupported is not a usable implementation.

## 3. Proposed user workflow

The hybrid shape is confirmed; the following screen organization and names are proposed.

### A. Start or reopen

Keep the ribbon compact and discoverable: a clearly labelled comparison entry point and a way to reopen a project. Put revision setup, detailed review, and merge preparation in the in-app workspace/window. A user should not need to know an internal feature name or find a hidden development command.

The first screen explains the three outcomes: compare revisions, review evidence, and optionally create a merged revision. Existing unrelated Power Tools commands remain available.

### B. Set up the project

1. Add files and inspect the proposed order, labels, full paths, timestamps, and duplicate-name warnings.
2. Correct the order and select the master. State visibly that order is a review sequence, not a claim about which file was copied from which.
3. Confirm sheet candidates, exclusions, manual pairings, and unpaired sheets.
4. Set the official-column boundary, with per-revision corrections. Review any weak boundary/header suggestions.
5. Confirm text-file parsing when detection is uncertain, including delimiter and encoding; preview the actual parsed content.
6. Start comparison with progress, cancellation, coverage information, and available duration estimates.

Setup uses a logical-sheet overview and exception-first revision matrix. Apply a boundary once per sheet family, show its per-revision mappings, and highlight anomalies rather than asking for hundreds of repeated confirmations. Bulk setup commands state their scope. Cancellation preserves confirmed setup. Document-family hints may later reuse user-confirmed anchors, boundaries, and fixed-position regions, with explicit drift detection and no mandatory keys.

All new windows, controls, and significant actions receive explanatory tooltips. A separate `?` guide explains the workflow and includes examples of ambiguity, skipped scope, reference adjustment, and merge safety. This rule applies across the new feature, not retroactively to every older add-in screen.

Keep context menus primary, with visible common commands, keyboard equivalents, accessible names, predictable focus, and usable controls at 200% display scaling. Hovering cannot be necessary to discover a primary action.

### C. Review differences

The primary view groups related changes. It shows first and last states initially, with expandable observations in every revision. The user can inspect any intermediate state, including temporary changes that disappear from the endpoint view.

When endpoints agree but another revision differs, the collapsed item remains in the initial review list and visibly says so, for example "Endpoints agree; differs in two other copies." Its content is reachable in one action. No endpoint-equality filter may silently classify it as unchanged.

Each group can expand to individual rows and cells. Context menus offer relevant actions, and multi-selection is required. Group selections do not prevent cell-level exceptions. Comparison status, review decisions, and unresolved alignment remain separate concepts.

Decision states are unreviewed, explicitly keep master, choose candidate/channel, deferred, and invalidated. Comparison coverage, human-review completion, and merge eligibility have separate counters. Unselected data stays at the master baseline but is not labelled reviewed. A partially reviewed output is allowed only with its remaining review scope clearly stated; only affected unsafe selections block merging.

Batch actions explicitly target selected cells, visible items, or the entire group. Preview applied, hidden, overridden, and blocked counts. Preserve cell-level overrides unless the user explicitly replaces them, and provide undo. A filter or collapsed group cannot silently expand the action's scope.

Ambiguity repair offers pair, find another/unlisted peer, no corresponding row, leave unresolved, and undo. Preview affected mappings and decisions before committing. Rejected candidates and old evidence remain traceable.

Every action shows a stable artifact identifier and distinguishing alias, independent of review position. Candidate source, first/last in review order, and output master are distinct labels. Reordering does not rename artifact identities, and two copies called "1.1" remain distinguishable without opening setup.

The detail view shows:

- Before/after and other revision candidates, with original coordinates and paths.
- Neighboring rows and the included row context, so a reviewer can recognize the plausible version without opening every source.
- Original fill and strikethrough alongside comparison annotations; report colors must not overwrite the evidence they explain.
- Exact character-level details and clearly labelled formula versus saved-result evidence.
- Source navigation links and internal snapshot links. If a source has moved, explain how to relink it rather than silently showing an empty detail.
- Matching alternatives and reasons when correspondence is uncertain. Do not display an uncalibrated similarity score as a probability.

Color roles live in centrally named constants. Use the requested red-to-green meanings and violet for special states, subject to adjustment during testing. Include text/symbol labels so color is not the only explanation. Preserve the add-in's established interaction affordances where applicable.

Strikethrough is never the report's generic character-deletion marker. Use a distinct, labelled annotation grammar so original fill/strike, added/removed characters, and acceptance states cannot impersonate one another. Whitespace markers are distinguishable from literal text. Copying exact evidence excludes all visualization decorations.

### D. Save and share

Produce a self-contained Excel review that can be read without the add-in and reopened for continued review through the add-in. It must not depend on a forgotten companion sidecar to retain the project decisions and shown evidence.

Proposed distinction:

- **Review workbook:** scope, revision inventory, retained comparison evidence/context, matching decisions, review selections, and logs sufficient to reopen the review. Original input binaries need not all be embedded in this workbook.
- **Complete project archive:** input files, review workbook, retained snapshots, decisions, settings, manifests, and logs. It can be shared as one archive and opened through the add-in. Relative archive references make it portable; original paths remain provenance.
- **Merge prerequisites:** if the review is shared without its original inputs, it remains useful for inspection, but creating an output requires the matching input snapshots to be available or relinked and verified.

Large evidence sets must not be silently cut at worksheet limits. Partition or progressively expose them while retaining exact data and coverage accounting. The storage mechanism remains an implementation choice and must survive an ordinary Excel open/save round trip.

The standalone-reading contract is stronger than embedded storage: without the add-in or original paths, ordinary Excel navigation must reach retained differences, intermediate history, character detail, and contextual snapshots. Add-in lazy rendering may accelerate the working view but is not the only path to required exported evidence. Lossless retained data is separate from visible cell coercion/length limits; deduplicate and partition presentation by text, hyperlinks, styles, and resource limits, not only row counts. Prove this with long TXT fields, a 100,000-difference report, varied formatting, and dense cases before committing to the container architecture.

### E. Create a merged revision

Use the chosen master as the starting point. The reviewer selects information from any supplied revision, previews the resulting operations and formula adjustments, and edits the output filename. Unselected data stays as in the master; this is a baseline, not a claim that the master is correct.

Create a new file, never overwrite an input, and record the result as a new revision. Ship the merge audit with the result/project so it cannot be separated from the selections it explains by accident.

## 4. Proposed matching and comparison design

These are behavioral boundaries, not a reuse of old classes or file layouts. The final alignment strategy and project architecture are not locked. A first exact-sequence/text component and a bounded reader trial are implemented; see the build checkpoint for their scope and evidence.

1. **Immutable input evidence.** Capture stable snapshots of the chosen file bytes, with project-local artifact identities and fingerprints. Preserve original paths and metadata separately. Extract stored information without executing macros, refreshing external connections, or recalculating sources.
2. **Explicit scope and correspondence.** Record sheet pairing, column boundaries/mappings, row mappings, exclusions, manual decisions, and unresolved regions. Each decision has provenance and can be invalidated visibly when its inputs change.
3. **Exact independent comparison channels.** Once counterparts are known, compare values, formulas, results, fill, and strike independently. Similarity used for alignment never excuses a real difference.
4. **Revision-wide evidence.** Keep observations across all revisions. Adjacent comparisons are changes along the chosen review sequence, not asserted ancestry. Resolving a selection against the master may need direct correspondence rather than blindly chaining uncertain links through intermediate files.
5. **Review and decision state.** Keep durable evidence and user choices separate from workbook rendering. Re-rendering, filtering, expanding history, or reopening must not lose selections or change their meaning.
6. **Preflighted merge operations.** Build an explicit plan against the selected master, validate dependencies, preview effects, and write only the approved new artifact.

Capture means saved file bytes at a recorded time, not the user's possibly unsaved Excel view. Never silently save a source; warn about unsaved state when reliably detectable. Acquire a stable copy while another process may be saving, or diagnose/retry instead of accepting a mixed capture. Navigation distinguishes the captured snapshot from the current original and discloses changed bytes before showing a live path as historical evidence.

Archive import and source navigation share the non-execution boundary. Enforce contained paths, collisions/no-overwrite, manifest verification, and compressed/decompressed size and entry-count limits. Read-only Excel opening alone is not a safe macro policy; adapters must handle their complete execution/refresh boundary without weakening the user's security settings. Prefer internal snapshot navigation. Tests must verify no unexpected execution, external access, or writes outside the project area.

### Bounded engine evaluation

Use the existing [research assessment](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-research-2026-09-09.md) to compare a daff/table-aligner route with an established ordered-sequence-diff route plus partial-match refinement. Where available, use a commercial comparator as an independent reference, not as unquestionable ground truth.

The sequence-based candidate should test distinctive unchanged anchors, informative content similarity, coherent gap alignment, and neighboring context. It must not accept the first weak downward match automatically. Blank cells, repeated zeros, repeated headings, and copied formulas are weak identity evidence. Manual fill and strike are comparison channels, not mandatory matching keys.

Column correspondence precedes or participates in bounded row/column refinement. The process must terminate deterministically, expose exhausted search or unresolved mappings, and stay within declared resource limits. Excluded columns cannot serve as hidden matching clues.

Reuse established character-diff functionality if it passes exactness and rendering tests. Select libraries only after checking their pinned version, license, integration cost, correctness, resource behavior, and packaging in the real add-in. Published documentation alone does not establish suitability. The first DiffPlex-based candidate now has synthetic and small real-file trial evidence; comparative engine benchmarks, rendering, and target-scale validation remain pending.

Evaluate optional reusable document-family hints and a small set of coherent interval explanations as alternatives to a single winning alignment. A reference-centred strategy must retain older-only entities outside the reference. These are candidates, not permission to require keys or normalize away differences. Compare a sheet-first workspace with a grouped review queue using the same difficult review task before choosing the main layout. An isolated, bounded worker and portable XLSX checkpoints with a private working journal are architecture candidates; neither can leave an undisclosed required sidecar.

### Reading and evidence fidelity

Test each supported file format independently. Do not infer XLSB support from an XLSX reader. Text inputs have no original fill/strike evidence; report that as unavailable, not as an observed format change. Preserve exact strings, leading zeros, quoted delimiters, and embedded line breaks according to the confirmed parsing settings.

Formula extraction must distinguish logical expressions from storage conventions such as shared formulas. Retain saved results without recomputing them. Compare resolved fill/strike properties rather than workbook-local style IDs. The exact treatment of visually equivalent fill encodings is an implementation proposal to validate, not an excuse for noisy style-table differences.

Optional Excel history may enrich provenance when genuinely present and understood. It is not required for comparison and does not override snapshot evidence or user pairing. Do not depend on unsupported revision metadata to reconstruct a family tree; the limitations and primary sources are documented in the research report.

## 5. Merge safety and remaining merge decisions

### Required safeguards

- Every selection points to an artifact, source evidence, target logical location, selected channel/operation, and the mapping that justifies it.
- Only affected operations are blocked by local ambiguity. A formula depending on an unresolved mapping is also affected, even if its own cell matches confidently.
- Preview formula adjustments using the final proposed layout, including interactions between multiple structural operations. Revalidate when selections change.
- Never substitute a saved calculated result for a formula silently. Never execute copied formulas or text as part of writing a review report; evidence is literal and non-executing.
- Validate the selected input/master snapshots and output destination before writing. A live path changing since capture must be disclosed; it must not silently replace the reviewed snapshot.
- Preserve unselected workbook content and relevant workbook features. A format or operation that cannot be safely preserved is an explicit unsupported/pending action, not a lossy successful export.
- Write to a new destination, validate the produced workbook, and publish completion only after success. Failure or cancellation must leave sources untouched and no apparently complete but partial result.
- Audit each applied operation with source and target coordinates, exact before/after evidence, selected revision, reference adjustment, decision provenance, operation identifier, and outcome. Do not invent who authored historical changes.

- Before enabling any write, classify its capability: ordinary constant, formula, fill, rich-text strike, table cell, merged region, array/spill region, or structural change. Reading safety/dependency metadata outside the comparison cap is distinct from using excluded content as matching evidence. Actual collateral changes require impact preview and explicit approval. Validate the whole output, not just selected cells.
- Give every merge operation a durable identity and prepared, validated, and published states. Output, audit, and project registration must have a recoverable commit protocol with exclusive new-destination creation. Interruptions and retry cannot duplicate application or replace another output. Define the delivery unit for text masters without appending audit rows to their data. Private working journals are allowed, but saved/exported projects never require a forgotten companion to reopen.

### Merge demonstrations, not blockers for read-only comparison

1. **Excluded columns during row restoration.** Inserting a worksheet row can move data outside the compared columns. Proposed default: keep existing right-side notes attached to their rows by inserting a whole row, but do not import excluded older content. Show this with a concrete example before implementing it as the default.
2. **Selecting content versus formatting.** Proposed independent selections for value/formula, fill, and strike, allowing text recovery without restoring an old color. Whole-group choices must show exactly which channels they apply.
3. **Workbook dependencies, before even a one-cell merge.** Demonstrate ordinary edits and restored/deleted rows, inserted columns, added sheets, formulas on skipped sheets, calculated table columns, names, merged cells, array/spill formulas, and external references. A safe read-only difference can be supported before its merge operation is supported, with that limitation visible. Preservation distinguishes unchanged logical content from approved reference adjustments caused by structure.

Partial-row shift reconstruction remains out of scope. Resolving such a region means reviewing ordinary cell differences and explicitly selecting content; it does not authorize an inferred local structural operation.

## 6. Persistence, progress, diagnostics, and presentation

### Project durability

Use explicit project and artifact identities, a versioned project format, and fingerprints of the retained source evidence. Reordering revisions, changing scope, correcting mappings, or adding files can change the meaning of existing selections. Preserve unaffected decisions; mark affected ones for review instead of applying them to a newly guessed location.

One versioned evidence/decision ledger is authoritative. Generated worksheet cells are views, not an alternative source of approved merge data. Define editable reviewer notes separately; edits/sorts/deleted metadata must be detected or safely disregarded with an explanation. Missing or damaged project state produces a diagnosed non-resumable review, not guessed approvals. Ordinary Excel save/reopen must preserve a valid checkpoint.

Copied projects are explicit independent forks; saves use generation checks so an old copy cannot silently erase newer decisions. This does not introduce live coauthoring or automatic decision merging. Reopening renders frozen evidence under its recorded reader/engine/settings versions. Deliberate reanalysis is versioned and invalidates affected decisions visibly. Show saved/unsaved state; a private recovery journal never becomes an undisclosed dependency of a portable checkpoint.

Save meaningful progress and decisions. Reopen/extend/re-export must work without requiring a permanent archive-management service. A review opened without the add-in remains readable; write/apply controls that require the add-in are clearly identified.

A generated review cannot silently discard unsaved user decisions on close. Reconcile the new persistence contract with existing add-in save/close behavior; a view refresh or settings change must not close the user's only Excel workbook or lose their filters and position.

### Scale and feedback

| Planning dimension | Typical | Worst target |
| --- | ---: | ---: |
| Revisions | 5 | 20 |
| File size per workbook | 10 MB | 200 MB |
| Sheets per workbook | 5 | 30 |
| Useful rows per sheet | 1,000 | 5,000 |
| Official columns per sheet | 30 | 200 |

These are validation targets, not promised runtimes. Use sparse extraction and bounded processing rather than materializing every possible cell across every revision. Do not impose an unexplained dense-grid limit based on formatting debris.

Show progress by phase and current artifact/sheet, with cancellation. Learn duration estimates from previous runs and keep the timing records in the user's temporary area without cell data. Label cold-start or uncertain estimates honestly.

Coverage is explicit: compared, excluded, unavailable, unresolved, or failed. A resource limit must identify the artifact and region, measured quantities, applicable limit, and corrective options. Never call an unexamined region unchanged or silently truncate a report.

### Diagnostics

Use unique run/issue/operation identifiers and verbose logs. Include the artifact path and fingerprint, revision and sheet, coordinates or region, stage, relevant settings/limits, mapping identifiers, exception details, source-code location, and suggested corrective actions. Persist enough context to investigate without reproducing the entire session. Full source cell content should stay with the local project evidence, not be added to timing telemetry.

The UI summarizes the issue plainly and makes detailed logs easy to open/copy. Serious safety or extraction problems are conspicuous; routine ambiguities remain reviewable without overwhelming the user. All diagnostic paths must distinguish incomplete output from successful completion.

### Excel presentation and integration

- Fit generated columns to their data within Excel's actual limits, including filter controls and timestamp readability; expand/wrap long evidence appropriately.
- Render populated sheets correctly on first opening. Requiring a sheet switch to see results is a failure.
- Show full revision paths. An empty changes view must explain whether there are no differences, a filter is active, scope was excluded, or comparison failed.
- Selecting a difference must populate its character/context details; placeholder-only or blank detail sheets do not count as delivered functionality.
- Use smooth text rendering, existing safe/silent dialogs, tooltips, the separate guide, and the add-in's relevant navigation conventions.
- Check that the feature is discoverable in both debug and installed configurations without changing unrelated ribbon functionality.
- Update the user-facing add-in reference alongside implemented behavior, not with unsupported future-completion claims.

## 7. Proposed delivery milestones and acceptance gates

Each milestone must leave a usable, verified outcome. Passing core tests does not substitute for seeing it work through the real Excel launch path. No milestone below is currently marked complete.

| Milestone | Deliverable | Exit evidence |
| --- | --- | --- |
| 0. Agreed cases and feasibility | Approved contracts; representative cases with known intended matches; independent format extraction/preservation checks; early standalone-report and state/capture feasibility. | Reader limits, typed evidence, non-execution boundaries, and passive Excel navigation/scale proven before architecture is locked. Real sample judgments come from the user. |
| 1. Alignment proof | Bounded candidate evaluation with readable row/column correspondence, exact differences, and a small multi-revision consistency trial. | Correct insert/delete and partial-match behavior; no silently wrong accepted mappings in the labelled corpus; explicit uncertainty; scope invariance, residual historical rows, and measured automation/review effort. |
| 2. Usable read-only vertical slice | Launch from the real ribbon; two-file setup, sheet pairing/skip, column cap, comparison, Excel review, character details, context, tooltips, guide, logs, cancellation, and save/reopen. | The user can complete the workflow on a real revision pair. No blank first-render or placeholder details; all four substantive channels are inspectable; inputs unchanged. Start with validated XLSX/XLSM handling, without representing that as full-format completion. |
| 3. Complete multi-revision review | All required input formats; corrected order and duplicate labels; expandable history; grouped review, multi-selection, manual pairing; new revisions; self-contained review and portable complete export. | Works on the target multi-file examples, including parallel copies and a middle-only change. Reopen on a moved project/archive preserves decisions and evidence. Resource and interruption behavior measured. |
| 4. Selective merging | Capability-gated cells/groups into a newly named master copy, with formula verification, dependencies before the first write, audit, recoverable publication, and new-revision registration. Add structural actions only after their examples and decisions are settled. | Whole-output fidelity, no unapproved collateral mutations, no duplicate application on retry, no overwrite, complete audit linkage, unchanged sources, and format-specific round-trip proof. |
| 5. Release validation and cleanup | Real Excel end-to-end verification, target-scale runs, recovery tests, accessible guidance, and removal or retirement of obsolete comparison entry points after explicit implementation approval. | One coherent launch/workflow, no stale competing comparison commands presented as the delivered feature, documented support boundaries, and a user-validated real-artifact run. |

Proposed implementation order is comparison accuracy first, then a complete read-only workflow, then revision history/portability, then merge. This does not remove later milestones from the requested feature. If XLSB or a required preservation case needs a materially different approach, escalate during feasibility work rather than hiding the gap until release.

## 8. Validation matrix

Tests need known expected correspondences and exact differences, not just plausible screenshots or low difference counts. Synthetic cases complement the actual artifacts; they do not replace them.

| Case | Required observation |
| --- | --- |
| Unchanged content with reordered internal style records | No invented value/formula/fill/strike changes. |
| One or many inserted/deleted rows | Existing downstream rows retain correct peers; inserted/deleted content remains visible. |
| A new row next to a partially edited row | Correct contextual match or explicit ambiguity, not a silent first-match guess. |
| Identical rows and repeated blocks | No invented certainty about which occurrence survived. |
| Many blanks/zeros/repeated formulas | Low-information equality cannot create a confident false pairing. |
| Partial-row cell shift | Ordinary cell differences; no independent table realignment presented as a fact. |
| Inserted official column | Correct correspondence and a visible shifted/correctable last-column boundary. |
| Changed unofficial columns | Neither differences nor matching are influenced by excluded content. |
| Weak headers, title blocks, repeated headers | Data is retained; strong headers assist without becoming mandatory keys. |
| Same-name, renamed, skipped, deleted/new sheets | Candidate, manual pairing, exclusion, and genuine non-pairing all work. |
| Same version label on divergent copies | Both artifacts remain selectable with distinct provenance. |
| A later file missing an earlier change | Absence and candidate content shown without claiming intent or ancestry. |
| Change only in a middle revision | History and relevance retain it even when endpoints agree. |
| Formatting-only empty tail | No output for never-populated cells; coverage feedback explains the boundary. |
| Content deleted but fill/strike remains | Evidence stays relevant because an earlier revision had content. |
| Formula showing blank; zero; false; error | Not mistaken for an unused empty cell. |
| Formula-only, constant-only, result-only edits | Channels remain distinct; unavailable cache is not an invented result. |
| Verified structural reference shift | Exact formula diff retained in the structural-adjustment group. |
| Changed range membership or unresolved reference | Not incorrectly labelled as a harmless adjustment; dependent merge stays pending. |
| Manual fill and whole/partial strike | Exact required evidence survives extraction, review, and supported merge. |
| Conditional/dynamic display | Follows the declared reader policy; no special evaluator or inferred human edit. |
| XLSB and separated text | Explicit format coverage, faithful parsing, literal/non-executing review evidence. |
| New revision or changed mapping after decisions | Affected choices are revalidated; none silently jump to another row. |
| Source moved or changed after capture | Provenance preserved; relinking verifies identity; changed bytes never substitute invisibly. |
| Save/reopen/export/import | Evidence, context, scope, selections, and logs remain available without an accidental sidecar dependency. |
| Merge with interacting structural selections | Final-layout formula preview and dependency validation agree with the new output. |
| Macro-enabled/binary/complex-feature master | Only supported safe writes are enabled; unselected features survive round trip. |
| Large sparse or highly repetitive inputs | Measured resource bounds, useful progress, cancellation, and no silent truncation. |
| Disk error, cancellation, interrupted save | Sources intact; incomplete result cannot be mistaken for success; actionable issue ID/log. |
| First open and navigation in real Excel | Visible data, usable widths, working links/context menus/multi-selection, populated character details. |

For each validation run, record mapping correctness, exact-difference completeness, how uncertainty was exposed, output fidelity, runtime/resource use, and what was actually tested in Excel. A smaller diff count is not itself evidence of a better match.

The labelled acceptance corpus requires zero silently wrong accepted correspondences, complete in-scope exact differences for correctly mapped cells, explicit coverage states, typed/formatting round-trip fidelity, and no unintended writes. These are test gates, not a claim of perfect accuracy on arbitrary future workbooks. Distinguish uniquely resolvable cases from genuine ambiguity. Measure automation coverage and user correction workload so an all-unresolved engine cannot pass as usable. Keep held-out cases; agreement between engines is not ground truth.

Add adversarial cases for conflicting partial-row offsets, cap discovery with duplicated outside headers, inconsistent three-revision chains, post-deletion blank formatting, numeric encoding/type/date-system differences, equivalent rich-text runs, endpoint-equal/middle-only changes, filtered batch overrides and undo, passive no-add-in details, altered/stripped metadata, unsaved/live source mismatch, archive collisions/limits, and interruption after output creation but before registration.

## 9. Next review points

The user has approved incorporating the review and starting implementation. Remaining judgments should be made with evidence rather than another long sequence of abstract yes/no questions:

1. The user supplied five sorter revisions, R1.0, R1.1, R1.2, R1.3, and R2.3, in review order. They have been read in the new read-only trial and fingerprint-checked unchanged. The official cap on `R1.S1.C1` is confirmed as Z inclusive; other sheets' caps, column correspondence, and expected ambiguous matches remain to be reviewed. A second document family and divergent same-version copies remain useful held-out additions.
2. Review a small set of ambiguous row/column matches during the alignment proof. The user's knowledge supplies the expected answer where snapshots alone cannot.
3. Before structural merging, review concrete examples for whole-row recovery outside the official cap, independently selecting formatting, and formula/dependency effects.

Implementation is authorized against this staged plan. Source files in the new feature and required integration may be changed; unrelated dirty work and original comparison inputs remain untouched. Do not claim a milestone or the feature delivered merely because foundational tests pass. Record actual builds, test evidence, remaining gates, and real Excel validation separately.

## 10. Adversarial finding traceability

All 21 findings from COMP-PLAN-AR-20260910 are adopted as contracts or explicit acceptance gates:

| Findings | Incorporated contract | Required gate |
| --- | --- | --- |
| ALIGN-01, ALIGN-02, ALIGN-03 | Contradictory offsets, two-stage cap discovery, one-to-one/cross-revision identity and residual evidence | Alignment proof |
| ALIGN-04, ALIGN-06 | Distinct observation states, ever-content blank formatting, canonical types/date context | Reader and exact-difference feasibility |
| ALIGN-05, SOLID-04 | Bounded formula verification and dependency/capability checks before every write | Before any merge |
| ALIGN-07, UX-06 | Effective character strike and non-conflicting annotation grammar | Read-only vertical slice |
| UX-01, UX-04 | Visible middle-only differences and persistent source/master identity | Review UI |
| UX-02, UX-05 | Scoped batches, overrides, undo, separate states, full ambiguity repair | Review UI and decision tests |
| UX-07, UX-08 | Exception-first setup, keyboard/discoverability/accessibility | Interaction trial |
| UX-03, SOLID-01 | Passive standalone access, exact retained data, multi-limit scale proof | Before container/layout lock |
| SOLID-02 | Authoritative ledger, checkpoint/fork/stale-save rules, frozen analysis | Resumable delivery |
| SOLID-03 | Durable operation identity, exclusive publication, audit and retry recovery | Before any merge |
| SOLID-05, SOLID-06 | Safe archive/navigation/acquisition and saved-versus-live snapshot provenance | Acquisition feasibility |
