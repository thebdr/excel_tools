# Adversarial review of the fresh spreadsheet-comparison plan

Date: 10 September 2026. Review ID: **COMP-PLAN-AR-20260910**.

Reviewed artifact: [fresh plan](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md), as read during the original review. Navigation links now point to the corresponding sections of the amended plan.

Follow-up on 10 September: the user approved incorporation of all 21 findings and authorized implementation. The plan now includes a finding-to-contract traceability table. This report preserves the original assessment; statements below about no source edits or benchmarks describe the review itself, not the subsequent build. Current evidence is recorded in [the build checkpoint](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-build-2026-09-10.md).

## Verdict

**The direction is viable, but several safety and interaction contracts need to become explicit before a full implementation. A bounded read-only alignment/feasibility trial can proceed once those trial contracts are defined.**

Three independent agents reviewed UI/UX, false positives/false negatives, and reliability/alternatives. The coordinating review checked overlaps, challenged scope-expanding recommendations, and verified consequential Excel facts against Microsoft documentation.

The reviewers raised 21 findings. These were design risks and missing operational definitions, **not observed bugs in a new implementation**. During this review no engines were benchmarked, no workbooks were opened or changed, and no source code was edited. The abandoned comparison implementation was not reviewed or used as the design foundation. The amendments proposed here were subsequently incorporated following user approval.

The most important distinction is between:

- Detecting different observations correctly.
- Inferring which observations belong to the same logical row/cell.
- Helping a person choose content without hiding relevant evidence.
- Safely producing exactly the approved result.

Success in one does not prove success in the others.

## 1. Highest-priority amendments

### A. Make correspondence conservative and internally consistent

Related findings: ALIGN-01, ALIGN-02, ALIGN-03. Plan: [structural matching](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:51), [scope](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:40), [revision-wide evidence](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:187).

The plan already requires partial matches and visible uncertainty. What is missing is an operational boundary on which matches may become accepted identities.

Add these invariants:

1. Contradictory shifts among informative columns cannot silently become a confident whole-row shift. Offer a local fixed-position constraint when the user wants ordinary cell differences; do not build independent table realignment.
2. Separate boundary discovery from comparison. A proposed shifted official cap is visible; only the resolved per-revision scope can supply row-matching evidence. A duplicate heading outside the boundary must not quietly enlarge it.
3. Correspondence is partial and one-to-one within each revision pair. One logical history cannot contain two rows from the same revision.
4. Conflicting cross-revision paths remain competing explanations. A bridging revision cannot silently fuse previously distinct identities.
5. Reordering changes presentation, not the raw observations or silently accepted identities. If an order-dependent reanalysis changes proposed matches, show that and revalidate affected decisions.
6. An entity absent from the master is still retained as recoverable historical evidence. Direct target correspondence is needed before a merge, not presumed from a chain.

This does not require universal perfect matching. It requires an honest boundary between an observation, a hypothesis, and a confirmed correspondence.

### B. Define observation states and exact-value semantics before choosing readers

Related findings: ALIGN-04, ALIGN-06, ALIGN-07. Plan: [comparison channels](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:64), [relevance](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:83), [reader fidelity](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:207).

A present blank cell, absent row, excluded region, unavailable format channel, failed extraction, and unresolved correspondence are different states. None should collapse into a generic empty value.

Apply the confirmed ever-had-content rule across logical positions, then preserve/retrieve formatting for their matched blank observations. Reading only populated cells from each revision independently can lose the formatting that remains after content deletion. Never-populated colored neighbors still produce no output.

Separate canonical typed values from storage encodings and display context:

- Numeric 1 and text "1" remain distinct.
- Numeric storage spellings 1 and 1.0 should not create a fictitious character edit.
- Text typing follows an explicit parser policy, not an adapter's opportunistic guessing.
- Date-system context matters without adding general number-format comparison. Microsoft documents that the same date has serials differing by 1,462 between the 1900 and 1904 systems; its basic Open XML sample can return dates as numeric serials. Do not hide underlying numeric changes or silently alter a recovered date's interpretation. [Date systems](https://support.microsoft.com/en-us/excel/date-systems-in-excel), [reading cell values](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-retrieve-the-values-of-cells-in-a-spreadsheet)

For rich-text strike, compare effective character formatting rather than raw run boundaries or unadjusted character offsets. Splitting one run into several equivalent runs must not create differences.

### C. Turn review actions into a precise, reversible contract

Related findings: UX-01, UX-02, UX-04, UX-05, UX-07, UX-08. Plan: [review workflow](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:131), [setup](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:116).

Keep the confirmed first/last default, but require a visible collapsed indicator when they agree and intermediate copies differ: for example, “Endpoints agree; content differs in two other copies.” These remain review items in the initial list.

Define separate decision states: unreviewed, explicitly keep master, choose candidate/channel, deferred, and invalidated. Keep comparison coverage, review completion, and merge eligibility as separate counters. A completed calculation is not a completed human review.

Every batch action must name its scope: selected cells, visible items, or entire group. Preview hidden items, overrides, blocked operations, and affected counts. Preserve cell overrides unless explicitly replacing them. Provide undo.

Example acceptance case: a group has 100 items, of which 20 are visible. Two visible items have overrides and one is unresolved. A visible-item action previews 17 applicable changes, two preserved overrides, and one pending item; the hidden 80 are untouched.

Ambiguity repair needs more than a candidate list: pair, find an unlisted peer, reject all/no counterpart, leave unresolved, and undo. Show the consequences for neighboring mappings and prior decisions before committing a repair.

Use stable artifact identifiers plus aliases at the action point. Distinguish first/last in review order, candidate source, and output master—even when two copies are both named “1.1.” Right-click remains primary, with discoverable commands and keyboard equivalents rather than hover-only instruction.

For 20 revisions and 30 sheets, use a sheet-family overview and exception-first setup. Eight anomalous boundaries should be obvious among hundreds of otherwise consistent sheet instances. Do not turn this convenience into automatic renamed-sheet pairing.

### D. Keep report annotations separate from source formatting

Related findings: UX-06, ALIGN-07. Plan: [original appearance and annotations](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:158).

**Do not use strikethrough as the report's ordinary character-deletion marker.** It already has substantive meaning in the source.

Similarly, green/red source fills must not look like acceptance/rejection decisions. Preserve original appearance in a clearly labelled evidence view; use a distinct annotation treatment for additions/removals, such as labelled gutters, markers, or borders. A manual strike, a removed character, and an approved selection must be distinguishable at once.

Whitespace markers must not look like literal source characters. Copying exact evidence must not include visualization decorations. Color remains configurable and is never the sole signal.

### E. Prove the standalone workbook promise early

Related findings: UX-03, SOLID-01. Plan: [save/share](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:160), [scale](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:253).

“Stored inside the workbook” is weaker than “inspectable without the add-in.”

Define the standalone-reading contract: a recipient without the add-in or original paths must reach retained differences, intermediate history, character detail, and context using materialized sheets and ordinary navigation. Add-in-driven lazy rendering may improve the live workspace, but cannot be the only path to essential exported evidence.

Excel limits include 32,767 characters per cell, 65,530 hyperlinks per worksheet, and 65,490 unique cell styles. A 100,000-difference report with two ordinary source hyperlinks per item can exceed the link limit well before its row limit. Partitioning only by rows is insufficient. These are documented limits, not proof that the requested report is impossible. [Microsoft Excel limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)

Use lossless retained evidence plus deduplicated/partitioned presentation. Test long TXT fields, many links, varied original formatting, and dense inputs. Measure report generation, opening memory, and ordinary Excel save/reopen before committing to the container design.

The maximum input geometry is a target, not evidence of actual density. Do not promise that sparsity always solves it, or reject the self-contained requirement without testing.

### F. Specify authoritative state, safe capture, and publication/recovery

Related findings: SOLID-02, SOLID-03, SOLID-05, SOLID-06. Plan: [input evidence](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:184), [durability](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:241), [publication](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:217).

Choose one authoritative, versioned evidence/decision ledger. Rendered cells are views, not an alternative source of trusted merge values. Define whether reviewer notes are editable; arbitrary edits, sorts, or deleted metadata cannot silently become approved selections.

Microsoft's Document Inspector can remove hidden worksheets and custom XML. This is not a claim that ordinary saving destroys either; it means missing project state needs a diagnosed non-resumable condition, not guessed reconstruction. [Document Inspector](https://support.microsoft.com/en-us/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks)

A copied saved project becomes an explicit independent fork; saving an older generation over a newer one needs conflict detection. This does not require coauthoring or automatic merging of reviewers' decisions. Reopen frozen evidence under its recorded reader/engine/settings versions; deliberate reanalysis is a separate operation with visible invalidation.

Capture explicitly means saved bytes at a recorded time. Never silently save a user's source. Distinguish “open captured snapshot” from “open current original”; a changed current file must not impersonate historical evidence. Define acquisition while another process is saving and warn about unsaved Excel edits when reliably detectable.

Extend the no-execution boundary to archive import, navigation, and reader fallbacks. Read-only opening alone is not macro safety: Microsoft documents programmatic opening defaults and an Excel 4.0 macro exception to the force-disable setting. Source navigation needs deliberate handling and must not weaken the user's security settings. [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open), [AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)

Archive import needs contained paths, collision/no-overwrite rules, verified entries, and resource limits. Microsoft notes that a basic extraction API does not bound total decompressed size or entry count. Verify the selected implementation rather than assuming a library supplies every safeguard. [ZIP extraction](https://learn.microsoft.com/en-us/dotnet/api/system.io.compression.zipfile.extracttodirectory?view=net-9.0)

For merging, output, audit, and project registration cannot simply be assumed to commit together. Give each operation a durable identity and prepared/validated/published states, plus recovery after each interruption point. Retrying cannot overwrite an existing result or apply twice. Define the delivery unit for text outputs too: appending audit rows to a CSV would change its data contract. A private working journal is acceptable only if a saved/exported review does not depend on a forgotten companion file.

### G. Gate every merge operation—not only structural changes

Related findings: ALIGN-05, SOLID-04. Plan: [formula adjustment](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:96), [deferred dependency decisions](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:231).

The plan already rejects textual reference substitution and general equivalence claims. Add an explicit verification capability matrix: verified, unresolved, or unsupported forms. Unsupported formulas remain inspectable; known ordinary-reference cases must still work.

If A10 moves to A14, changing =ROW(A10) to =ROW(A14) follows the same logical cell but changes the coordinate-derived result. It may be described as a structural expression adjustment, never automatically as harmless or equivalent. An unchanged =INDIRECT("A10") can keep addressing the old coordinate after the original input moves. Unchanged text or cached results cannot certify safe relocation. These examples follow Microsoft's function definitions; no workbook experiment was run. [ROW](https://support.microsoft.com/en-us/excel/functions/row-function), [INDIRECT](https://support.microsoft.com/en-us/excel/functions/indirect-function)

Also forbid circular proof: normalizing formulas with a guessed mapping does not make those normalized formulas independent evidence that the mapping is correct.

Move operation-level dependency checks before the first cell merge. Excel documents calculated-column formula propagation; the actual behavior of any chosen writer still needs testing. A one-cell write may not be isolated merely because the requested selection is small. [Calculated columns](https://support.microsoft.com/en-us/excel/use-calculated-columns-in-an-excel-table)

Distinguish comparison scope, safety inspection, and approved mutation. Reading dependency metadata outside the cap may be necessary for safety; using excluded content as matching evidence remains prohibited. Actual collateral changes require an explicit impact preview and approval. Validate tables, names, merged/array/spill regions, and affected references for every enabled operation and output format.

## 2. Small adversarial cases to retain in the test corpus

These are synthetic thought experiments, not measured failures.

### Partial shift with conflicting column evidence

| Row | Old A | Old B:D | New A | New B:D |
| --- | --- | --- | --- | --- |
| 1 | A | 10 / x / p | A | blank |
| 2 | B | 20 / y / q | B | 10 / x / p |
| 3 | C | 30 / z / r | C | 20 / y / q |
| 4 | blank | blank | blank | 30 / z / r |

Matching old row 1 to new row 2 may maximize three columns' agreement, but violates the intended ordinary-difference interpretation of this partial shift. Do not confidently infer that whole-row identity. After a fixed-position constraint, expose the B:D differences without a local structural-merge proposal.

### A convincing but inconsistent multi-revision chain

R1 has (Motor,10) and (Valve,20). R2 has (Motor,20). R3 has (Valve,20).

Partial matches by description and then value can connect R1's Motor through R2 to R3; a direct exact match also connects R1's Valve to R3. A logical history must not absorb both R1 rows. Keep the conflict visible.

### Content deletion with meaningful later blank formatting

A uniquely matched row's field is "retain" in R1, blank/red in R2, and blank/blue in R3. Preserve deletion and later relevant fill evidence. A neighboring cell empty in every revision still produces no output. Replacing R2 with an extraction failure must not turn the failure into a deletion.

### Endpoints agree but useful data exists elsewhere

Across five supplied copies, a cell is blank in the first and last, but populated in a middle copy. Its collapsed review item must visibly advertise the intermediate difference before any expansion.

### Equivalent strike runs and shifted character offsets

"ABCD" has C and D struck. Repartitioning storage runs is not a change. Inserting X to produce "XABCD", while C and D remain struck, must not invent strike toggles just because offsets moved. Striking C for the first time is a genuine formatting change. If character correspondence is ambiguous, retain exact snapshots rather than inventing an editing history.

### Publication interrupted after output creation

Write a valid new output, then interrupt before project registration. Reopening/retrying must identify the existing operation and result, retain the audit, and avoid duplicate application or overwriting another file.

## 3. Alternatives worth evaluating

These are proposals, not selected implementations or accuracy claims.

| Alternative | Benefit | Cost or failure mode | Fit with confirmed requirements |
| --- | --- | --- | --- |
| Optional reusable document-family hints | Reuses confirmed headings, official boundaries, section anchors, and fixed-position regions for standardized documents. | Template drift can reinforce a bad assumption; hints need provenance and visible validation. | Strong fit if optional, with no mandatory key and no automatic renamed-sheet pairing. |
| A few coherent interval explanations instead of one winning alignment | Shows the disputed region and competing match/gap interpretations; local constraints resolve related ambiguities together. | More evidence/UI complexity; agreement among a limited candidate set is not proof of uniqueness. | Strong fit with local uncertainty and manual eyeballing. |
| Reference-centred matching with residual historical entities | Reduces dependence on long adjacent-match chains and keeps selections related to the target. | Rows absent from the reference need their own retained evidence; the reference must not become assumed truth. | Compatible with first/last presentation and user ordering if those remain views, not identity rules. |
| Isolated comparison worker with bounded, disk-backed working evidence | Separates expensive analysis from Excel's UI and may improve cancellation/crash containment. | More packaging and recovery work; it does not make a huge final XLSX cheap to open. | Compatible if private caches are disposable and the exported review stays self-contained. |
| Complete portable XLSX checkpoints plus a private working journal | Saves durable decisions without rewriting a large report after every click; supports diagnosed recovery/forks. | More lifecycle/versioning work; ordinary Excel round trips still need proof. | Compatible only if no hidden companion is required to reopen a saved checkpoint. |
| Narrow, capability-gated copy-on-write merging | Makes supported operations explicit and focuses preservation testing. | Some operations/formats stay read-only during intermediate milestones. | Compatible with staged delivery, not with permanently dropping required formats or silently converting XLSB. |

### UI alternatives

1. **Sheet-first aligned inspection:** a large, familiar worksheet context with a grouped difference navigator, candidate revision selector, and detail panel. Particularly promising for standardized layouts and judging row matches. It must not try to put twenty full sheets side by side.
2. **Review queue plus evidence dossier:** a grouped list, substantial contextual snapshots, revision history, and a separate selections tray. Clearer for prioritizing decisions, but risks hiding worksheet structure in tiny cards.
3. **Workbook-first evidence with a narrow decision pane:** strongest continuity with the portable artifact, but selection, filtering, protected/generated evidence, and add-in decisions need especially precise boundaries.

**Coordinating recommendation:** test sheet-first inspection against the queue/dossier approach on the same difficult task. Keep the first/last default, right-click actions, multi-selection, and portable workbook in both. Evaluate optional document-family hints alongside the general alignment method. Do not lock a screen or engine from a diagram alone.

A baseline-first replacement for the user's first/last default, mandatory business keys, cloud processing, independent table-shift reconstruction, or a required external project database would change confirmed scope. None is adopted by this review.

## 4. Changes to the delivery gates

### Before locking architecture or building the substantial review UI

- Resolve the scope bootstrap, observation-state, typed-value, and identity invariants.
- Perform a small multi-revision consistency case during the alignment proof, not only a two-file experiment.
- Prove an exported review's no-add-in navigation and realistic size/format limits.
- Define decision/action states, source/master identity, authoritative saved state, safe capture/import, and annotation grammar.
- Compare the two UI approaches using the same task: find a middle-only change, reject a wrong peer, recover a candidate without copying its old color, and undo a filtered batch decision.

These do not require completing the full product before learning from a trial.

### Before the first resumable/exported delivery

- Exercise ordinary Excel save/reopen, copied projects, missing metadata, edited/sorted report cells, stale saves, source moves, and deliberate reanalysis.
- Demonstrate keyboard/context-menu parity, large display scaling, first-render correctness, useful progress/cancellation, and exception-first setup.
- Confirm that all retained required detail is reachable without the add-in.

### Before any merge

- Pass operation-level capability/dependency checks, including apparently simple cells.
- Define exact authority for collateral structural/reference changes.
- Fault-inject through output/audit/project publication and retry.
- Verify the whole output and unselected features, not merely the selected cells; preserve source fingerprints.

### Quantitative validation contract added by the coordinating review

Use both representative real artifacts and constructed cases with independently known intended transformations. Mark cases as uniquely resolvable or genuinely ambiguous; do not penalize honest uncertainty where the evidence cannot identify an occurrence.

For the labelled acceptance corpus, require:

- Zero silently wrong accepted correspondences.
- Complete exact-difference reporting for the agreed scope and correctly mapped cells.
- Explicit accounting for excluded, unavailable, unresolved, and failed regions.
- Correct preservation of typed values and formatting evidence through the enabled readers/export routes.
- No unintended writes or source modifications.

These are test gates, **not a promise of zero errors on every future workbook**.

Also measure automation coverage, false added/deleted/changed groups, unresolved workload, user correction effort, runtime, and memory. An engine that marks everything unresolved is not a usable success. An engine that reports fewer differences by matching the wrong rows is not a success either.

Keep a held-out subset of cases so tuning on the first two artifacts does not become the only evidence. Agreement between two engines or a commercial comparator is a useful signal, not ground truth. Decide thresholds and review-workload targets from the trial, not invented percentages.

## 5. Finding register

P1 means a serious contract or safety gap to resolve before the dependent milestone; P2 means a meaningful UX/exactness issue. Neither label asserts an implemented defect.

| ID | Priority | Finding | Principal plan reference |
| --- | --- | --- | --- |
| ALIGN-01 | P1 | Partial shifts can be mistaken for global row identity shifts. | [L51](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:51) |
| ALIGN-02 | P1 | Moving official caps need a separate discovery/confirmation stage. | [L40](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:40) |
| ALIGN-03 | P1 | Adjacent partial matches need global one-to-one identity consistency. | [L187](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:187) |
| ALIGN-04 | P1 | Empty, absent, unavailable, excluded, and failed observations must differ. | [L83](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:83) |
| ALIGN-05 | P1 | Formula verification requires an explicit support/proof boundary. | [L96](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:96) |
| ALIGN-06 | P1 | Typed values, storage encodings, and date context need one reader contract. | [L64](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:64) |
| ALIGN-07 | P2 | Equivalent rich-text runs and shifted offsets must not invent strike edits. | [L64](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:64) |
| UX-01 | P1 | Endpoint-equal, middle-different content needs a visible collapsed signal. | [L131](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:131) |
| UX-02 | P1 | Batch scope, hidden items, overrides, undo, and review states are underspecified. | [L137](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:137) |
| UX-03 | P1 | Standalone readers need usable detail, not just embedded evidence. | [L160](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:160) |
| UX-04 | P2 | Source identity, review position, and master direction must remain visible. | [L145](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:145) |
| UX-05 | P2 | Ambiguity repair needs reject-all, find-another, no-peer, and undo. | [L51](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:51) |
| UX-06 | P2 | Report deletion/approval decoration must not mimic manual strike/fill. | [L158](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:158) |
| UX-07 | P2 | Large setup needs exception-first review and scoped bulk operations. | [L116](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:116) |
| UX-08 | P2 | Context menus and tooltips need discoverable and keyboard-accessible counterparts. | [L129](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:129) |
| SOLID-01 | P1 | Standalone XLSX fidelity/scale needs an early measured feasibility gate. | [L160](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:160) |
| SOLID-02 | P1 | Authoritative state, copied projects, stale saves, and reanalysis need contracts. | [L241](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:241) |
| SOLID-03 | P1 | Output, audit, and project registration need recoverable publication. | [L217](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:217) |
| SOLID-04 | P1 | Dependency safety applies before even a one-cell merge. | [L231](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:231) |
| SOLID-05 | P1 | Non-execution and resource boundaries must cover archive import/navigation. | [L184](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:184) |
| SOLID-06 | P2 | Saved snapshots and a live/unsaved Excel view must not be confused. | [L217](C:/Source/Repos/excel_tools/docs/spreadsheet-comparison-plan-2026-09-10.md:217) |

## 6. What should stay unchanged

The original plan already correctly separates review order from ancestry; supports same-labelled copies; avoids mandatory keys; keeps weak headers as data; retains exact differences after fuzzy matching; treats unavailable evidence distinctly in principle; avoids source recalculation; preserves manual formatting policy; localizes uncertainty; and writes only new outputs with provenance.

Do not turn this review into a return to the abandoned implementation, a generic version-control product, mandatory template setup, automatic “newest wins,” or a blanket refusal to compare anything ambiguous.

**Recommended next action:** amend the plan with these explicit contracts, then run the small alignment, portable-report, and interaction trials. Library thresholds and final screen layout remain empirical choices. No implementation authorization is inferred from this review.
