# SWP Bridge — write-path robustness: test checklist

*Covers the fixes of 2026-08-22 for the robustness review of 2026-08-21 (7 confirmed defects, the
reopened `#REF!` finding, and the guards wired on the way). Layer 1 is what the automated suite
already proves; layer 2 is what only Excel and the real files can prove. Tick as you go.*

## 0. Before anything

- [ ] C: has a few GB free. A run reporting hundreds of failures **in about 2 seconds** is the disk
      (temp-file tests failing instantly), not a regression — free space and rerun.
- [ ] `dotnet test tests/ExcelPowerTools.Tests` is green and the total is noted. The working tree
      may still carry another tool's uncommitted `Comparison/RevisionBridge` files, which do not
      compile — test from a clean worktree of `HEAD` with the SWP Bridge files copied in, or after
      that work is resolved.
- [ ] Release pack rebuilt with Excel closed (`src/ExcelPowerTools/bin/Release/net48/publish/
      ExcelPowerTools-AddIn64-packed.xll`), and the installed `.xll` replaced.
- [ ] Real files at hand: SWP root `…\FVZ_Intermarche\3 Software\3_Diagnostica\SWP\sorter\`
      (dated version folders with `SWP_01_Sorter_v2.xlsx` + `CONFIG_PLANT_01_Sorter.xlsx`), the IO
      list `8FVZ_IO0001_R1.3 - sorter.xlsx`. Work on **copies**: a scratch copy of the IO list for
      edits, and **version-folder mode ON** so the original SWP is never written.

## 1. Automated coverage (passes in the suite)

| Fix | What it proves | Tests |
|---|---|---|
| C1 — empty IO Device never blanks J | no mismatch: J kept + `_Log` line; append: J kept, label without device; overwrite: I written, J kept | `SwpDiffTests.An_empty_io_device_never_blanks_the_swp_device`, `…In_append_mode_an_empty_io_device_keeps_j…`, `…In_overwrite_mode_an_empty_io_device_keeps_j…` |
| C2 — identity-less generated rows | no H/I/J write, no mismatch, no bare label; IO-list row with no identity warns once | `…An_identity_less_generated_row_never_touches_the_swp_identity`, `…An_io_list_row_without_any_identity_is_reported…`, `…A_partial_identity_compares_only_its_present_members` |
| H1 — empty composed description | T/U kept + `_Log`; the label alone still lands | `…An_empty_composed_description_never_blanks_an_existing_one`, `…An_empty_composed_description_still_carries_the_fld_label` |
| H2 — trust beats the wholesale rewrite | trusted row keeps its formula, ordinary row rewritten, one summary line | `…Trusted_rows_keep_their_config_formula_under_the_wholesale_rewrite` |
| M1 — the FLD group | marking (J blank + T + U + Q/R ext.), none in overwrite mode; toggle together; split group held back; group follows the diff; sidecar round trip; patcher applies the group whole or not at all | `…In_append_mode_the_device_blank_and_its_label_writes_form_one_group`, `…Without_a_device_blank_there_is_no_group`, `SwpQueueTests.An_fld_group_toggles_together…`, `…A_split_group_is_held_back_from_apply`, `…The_group_follows_the_fresh_diff_on_reconcile`, `SwpSidecarJsonTests.Round_trips_everything`, `SwpXlsxPatcherTests.An_fld_group_applies_together_or_not_at_all`, `…A_whole_group_with_nothing_else_to_apply_yields_no_bytes`, `…An_intact_group_applies_whole` |
| M2 — value-only Q/R, any formula shape | literal → `"text"&" "&T{row}` (empty → `T{row}`) + `_Log`; a formula not ending in `)` is extended | `…A_value_only_description_cell_becomes_a_formula_that_shows_the_label`, `…A_formula_of_any_shape_gets_the_extension` |
| Reopened — error-valued identity | `#REF!` in A / `#N/A` in D: row skipped, `_Log` names row + cell | `SwpSheetReaderTests.An_error_valued_identity_cell_is_skipped_with_a_warning` |
| SharedMask wired | a literal over a shared-formula cell in ANY column conflicts | `SwpXlsxPatcherTests.A_shared_formula_cell_refuses_a_literal_in_any_column` |
| One write per cell per batch | the second action on a cell conflicts | `…Two_actions_on_one_cell_let_only_the_first_through` |
| Coherence case | lowercase Plant_Inputs signal id still matches | `SwpCoherenceTests.Plant_inputs_matching_ignores_case…` |

## 2. Field checks (Excel, real files)

Common setup for every check: *open workspace* on the sorter SWP root, plant type MSC, the scratch
IO list, version-folder mode, `_Log` sheet visible. "Refresh" = the dashboard's Refresh.

### F1 · C1 — the IO list has no Device, nothing else differs
- [ ] In the scratch IO list, blank the **Device** cell of a CABINET alarm whose SWP row has a
      value in J. Leave FU / Location as they are. Refresh.
- [ ] Workspace: the row's *Device ▸* cell is empty; `_Queue` holds **no** `Device (J)` action for it.
- [ ] `_Log`: one Warning `the IO list carries no device for this row — the SWP's "…" is kept`.
- [ ] Apply → open the new version's SWP: J still holds the device.

### F2 · C1 in append mode with a mismatch
- [ ] Same row: also change its **Location** in the scratch IO list (a mismatch), Device still blank;
      workspace default = append. Refresh.
- [ ] *Description 1 ▸* ends with the label **without a device** (`=S1+NEWLOC`); still no `Device (J)`
      action; the Verdict is orange (label mismatch); `_Preview` shows J unchanged.

### F3 · C2 — generated rows (an upstream workspace)
- [ ] Open an **UPS** workspace (FUR_Jadlog PRODUCTION tree: `…\DB_Params\Sensor.csv` +
      `StrobeDevice.csv`) in **overwrite** mode. Refresh.
- [ ] Every strobe row (card sheet `(generated)`) queues **no** `FUnit value (H)`, `Location value (I)`
      or `Device (J)` action; `_Preview` on one shows H / I / J unchanged.
- [ ] Switch the workspace to append (reconfigure in place): no strobe row proposes a T / U ending in
      a lone space or a bare label; `_Log` has no "no functional unit…" line for generated rows.
- [ ] A photocell whose `PHT_*` row is missing logs `… no address and no electrical identity (the SWP
      row keeps its own)` and queues no identity write.

### F4 · H1 — the IO list has no description
- [ ] Scratch IO list: blank **both parts of both descriptions** on a row whose SWP T holds text. Refresh.
- [ ] No `Description 1 (T)` / `Description 2 (U)` action; `_Log` Warning `the composed description is
      empty — the SWP's own text is kept`.
- [ ] Add a Location mismatch on the same row: *Description 1 ▸* is the bare label `=S1+…-…`.

### F5 · H2 — trust + wholesale rewrite
- [ ] Setup form: tick **Never overwrite CONFIG_PLANT descriptions** AND **Rewrite the Q/R description
      formulas wholesale**; reconfigure in place.
- [ ] A `P*`-typed module row (or a generated row): *Formulas* column is **not** `rewrite`; an
      ordinary alarm row reads `rewrite`.
- [ ] `_Log`: exactly one line `N row(s) keep their CONFIG_PLANT description formulas …`.
- [ ] `_Preview` of the trusted row: the after-recalc Q text is not empty (CONFIG's template still feeds it).
- [ ] Tooltip of the wholesale checkbox mentions the precedence.

### F6 · M1 — the FLD group on `_Queue`
- [ ] Append mode; a row with a Location mismatch whose SWP J holds a device. Refresh. `_Queue` shows
      `Device (J)` → *(empty)*, `Description 1 (T)`, `Description 2 (U)` (and `Formula (Q)` / `(R)` when
      the row's formula is template-only), every one with **Group = `FLD`**; the row's `I/O addr (G)`
      or `Export (B)` actions, if any, have an empty Group.
- [ ] Double-click Include on the **T** action: **all** FLD rows of that group flip to ☐ together.
      Double-click the **J** action: all flip back to ☑. Type `x` / clear one: same.
- [ ] `swp-bridge.json` (SWP root) carries `"fldGroup": "SHEET|DB_…>n>b"` on those actions.
- [ ] Group excluded → Apply: none of them applied. Group included → Apply → open the SWP in Excel,
      recalc: J blank, T / U carry the label, Q / R **display** it.
- [ ] Conflict path: build, then edit that row's **T** cell in the SWP in Excel (save, close), Apply
      **without** refreshing. `_Queue`: the T action reads Conflict (three-way), the J action reads
      Conflict `held back — another FLD edit of this row conflicted…`, and J in the file still holds
      the device. Refresh → both re-base to Pending and apply together next time.
- [ ] Hand-edit the sidecar: set `"include": false` on one member only, reopen the workspace → the
      Apply confirmation reads `N included action(s) held back: …`; nothing of that group is written.
- [ ] Home sheet: the `_Queue · Include` gesture line mentions the group; the `Group` column has **no**
      blue border (no gesture), Include keeps its border.

### F7 · M2 — a value-only Q cell
- [ ] Scratch SWP copy: on one bit row, replace Q's formula by its **value** (copy, paste values);
      create a Location mismatch for that row in the scratch IO list; append mode. Refresh.
- [ ] `_Queue`: `Formula (Q)` with New value `"<the old text>"&" "&T{row}`; *Formulas* reads `append`;
      `_Log` Warning `… is a plain value, not a formula — rewritten as a formula over the same text …`.
- [ ] `_Preview`: the after-recalc Q text is the old text + the label.
- [ ] Apply → open in Excel: Q displays old text + label; **no repair dialog**.

### F8 · reopened — an error-valued identity cell
- [ ] Scratch SWP copy: make one bit row's **A** formula evaluate to `#REF!` (delete the cell it
      references), save in Excel. Refresh.
- [ ] `_Log` Warning `SWP sheet 'X' row N: the tag type (A) is #REF! — the row cannot be matched …`;
      the IO row targeting it reads **RowMissing** (red, `tag not found`).
- [ ] Break more than ten rows on one sheet: ten listed, then one `… N rows carry an error-valued
      identity` line.

### F9 · shared-formula guard on a column the old guard ignored
- [ ] Scratch SWP copy: put `=Q17` in **T17** and fill it down over a few rows (Excel stores a
      fill-down as one shared formula); save. Target one of those rows with a changed description.
- [ ] Apply: the `Description 1 (T)` action reads Conflict `shared formula — not editable headlessly`;
      the file opens afterwards **without a repair dialog**.

### F10 · package integrity on the real 13–25 MB file
- [ ] After any Apply above: the patched SWP opens in Excel with **no repair dialog**; the `.bak` (or
      the previous version folder) is intact; the Apply summary's applied count equals the number of
      `Applied` rows in `_Queue`; Q / R show the expected text after recalculation.

### F11 · coherence case
- [ ] MSC workspace with `Plant_Inputs.csv`: a mnemonic whose CSV signal id differs only in case
      reads **ok** in the *Plant_Inputs* column.

### F12 · filters survive the override re-render (2026-08-24)
- [ ] Filter the Workspace sheet (drop-down: one Sheet value + a Verdict value), then tick
      `[Mismatch OK]` on a visible row: the diff recomputes and **both filters are still applied**
      (the row set stays narrowed; the tick is normalized to ☑).
- [ ] Repeat with a text filter ("contains") on Mnemonic and with a dashboard row activation
      (filter-to-sheet): both survive an override edit; clearing a filter stays cleared.
- [ ] A number/date-shaped criteria on `SWP row` survives too. (A filter **by colour** is the
      known COM limit: it cannot be read back and is dropped by the re-render.)

### F13 · IO Bridge — ISPC plant CSVs load (2026-08-24)
- [ ] IO Bridge workspace with the ISPC pair
      `…\9_Backup\20260820_1715\FVZ_ISPC1_20260820_1715\PRODUCTION\ISPC\db_csv\Plant_Inputs.csv` +
      `Plant_Outputs.csv`: **builds** (the Outputs file used to be refused — its data rows carry a
      trailing `;#;` marker column the header line lacks).
- [ ] *Get mnemonics from CSVs* on the same Outputs file: the summary reads `N row(s)`, not
      "not a plant CSV".
- [ ] A scratch copy of the Outputs file with the header renamed to junk (`#;A;B;C;…`) still
      builds, and the pulls' note says `header not recognized — columns taken by the legacy positions`.
- [ ] Picking `IODriver\BusField.csv` as the Outputs CSV still refuses, with the message naming
      both what was looked for and the positional fallback.

## 3. Regressions to re-run by eye

- [ ] The three field cases of 2026-08-21 still behave: CABINETAL01914 (the safety door) lands its
      `+SG14` in I (overwrite) or in the label (append); a CHUTE template-only formula gets exactly
      one `&" "&T` extension; a `+EMERGENCY` row with the standard `SUBSTITUTE(…&T…)` formula gets
      **no** extension.
- [ ] A sidecar written before this change (no `fldGroup`) loads without notes; the Group column
      fills on the next refresh.
- [ ] Belt sheets: no `Device (J)` action unless `[Force IDs]` is ticked — with it ticked in append
      mode the blank still rides in an FLD group.
