# IO Bridge — porting plan (from SWP Bridge), 2026-11-03

The IO Bridge is the older of the two bridges; the SWP Bridge grew the workspace discipline the repo
now treats as rules (CLAUDE.md "UI conventions"). This plan brings the IO Bridge up to the same
shape, item by item, in an order where each step lands alone and the workspace stays usable after
every one. Two requests from the same review are appended: mnemonic syntax proofing against the
official `IO_Defines` workbooks, and a single browse for the CSV pair.

Conventions used below: **S / M / L** are relative sizes (S = a session, M = a day, L = several
days). "Core" is `src/ExcelPowerTools.Core/IoPipeline/IoBridge`, "add-in" is
`src/ExcelPowerTools/IoPipeline/IoBridge`, "tests" is `tests/ExcelPowerTools.Tests/IoPipeline/IoBridge`.
Every item ends with the same three obligations: unit tests in Core where the logic lives, a
paragraph in `docs/add-in-reference.md` §IO Bridge, and one real build of a workspace on the FVT
LOWER SORTER files (`8FVT_IO0010 … R1.4.xlsx` + the 2026-11-03 `Plant_Inputs/Outputs.csv`) before
the item is called done — a unit test that the model looks right is not evidence that Excel shows it.

---

## Phase A — observability (no behaviour change to join or apply)

### A1 · `_Log` findings sheet and apply log file — item 3 (S)

**What.** Every build and every apply records findings — rows the join could not pair and why,
names the CSV carries twice, moved rows and the side(s) their alignment was refused on, CSVs skipped
at apply (open in the delimited workspace), ambiguity notes, regex counts (A4). Rendered on a
`_Log` sheet with a jump link per finding (double-click → the row on its card sheet, blue border per
the rule), and written as `IoBridge_Log.txt` next to the IO list on apply.

**Where.** Core: an `IoBridgeLog` entry list on `BridgeWorkspaceModel` (reuse the shape of
`SwpLogEntry`: level, sheet, row, text), filled by `IoBridgeJoin` and `CsvPatchApplier`. Add-in: a
`RenderLog` next to `IoBridgeQueueSheet` modelled on `SwpBridgeAuxSheets.RenderLog`/`LogLines`;
tab colour grey; sits right after `_Queue`.

**Tests.** Join emits one entry per unpaired row and per duplicate name; apply emits per skipped
file. **Docs.** New "Findings" paragraph.

### A2 · AutoFilter survives a re-render — item 5 (S)

**What.** Any refill of a card sheet (override edit in Phase C, rebuild after apply, column-config
apply) captures the AutoFilter criteria first and re-applies them after. Required by CLAUDE.md
("A re-render never discards the user's view state"); the IO Bridge rebuilds sheets fresh today.

**Where.** Lift `CaptureAutoFilter` / `RestoreAutoFilter` out of `SwpBridgeSpineSheet` into a
shared helper under `src/ExcelPowerTools/Interop` (both bridges call it), then call it around the
refill in `IoBridgeBuilder`. Colour/icon filters remain the COM limit, as documented for SWP.

**Tests.** COM-only; verify by hand (filter `//Match` to `moved`, refresh, filter still set).
**Docs.** One sentence in "The view".

### A3 · Tab colours and double-click conventions — item 11 (S)

**What.** Card sheets blue `RGB(180,210,240)` tabs, `_Queue` yellow `RGB(255,235,156)`, `_Log`
grey, `_Preview` (A4/B1) yellow; `_Columns` unchanged. Audit that every double-click gesture has
the blue border and vice versa (the two existing source-opening cells, the `_Queue` Include
toggle, the new `_Log` links). A `_Preview` sheet sits right after the card sheet it was opened
from, never at the strip's end.

**Where.** `IoBridgeBuilder` — a `ColorWorkspaceTabs` mirroring `SwpBridgeBuilder.ColorWorkspaceTabs`.

### A4 · Report what each regex moved — item 7, first half (S)

**What.** Each build logs (A1) and shows on the Home/first sheet header how many IO rows the
exclude regex filtered (whole / partial) and how many CSV records the compare-virtuals regex opted
in. A regex that does not parse is a warning line, never fatal (already inert today; make it
visible).

**Where.** Counters in `IoBridgeJoin.Build` → model → `_Log` + header note.

---

## Phase B — decisions become visible and reviewable

### B1 · `_Preview` sheet — item 1 (M)

**What.** A scratch sheet showing one row's write in its neighbourhood, `current → new` per
changing cell, opened by double-clicking the row's `//Match` cell (one column per action: the two
existing gesture cells keep their actions; the verdict cell takes the preview). Two panes:

- **CSV side** — the record's line as the file holds it, the records at ±3 addresses on the same
  card in file order, and for a *moved* row both locations (where the record is, where the IO row
  says it should be, the occupant of each). Queued SetFields show `old → new` on the record; a
  queued integrate shows the record that will be appended; a delete shows it struck through.
- **IO side** — the IO-list row in source column order with the queued cell changes marked.

Opens with the workspace on the first row that has a pending action, otherwise the first anomaly.

**Where.** Add-in `IoBridgePreview` modelled on `SwpBridgePreview` (Open / Prepare / Rerender);
data from `BridgeRow`, `CsvBinding.AllFields` and `session.Queue`. Core: a small
`IoBridgePreviewModel` builder so the neighbourhood and the `old → new` composition are testable
without Excel.

**Tests.** Neighbourhood selection, moved-row two-location rendering, the three action kinds.

### B2 · Per-row overrides that name the opposite of the default — item 6 (M)

**What.** Bracketed, editable, sidecar-persisted columns on each card sheet, per mnemonic
identity (`SignalRef.IdentityKey`, fallback mnemonic key):

- `[Ignore row]` — the row claims nothing and carries no verdict (today only the exclude regex can
  do this, globally). Replaces nothing; the regex stays the bulk tool.
- `[Mismatch OK]` — the name/address disagreement is deliberate; the verdict is kept but drawn grey
  and excluded from the anomaly counts; no action is auto-proposed for it (B3).
- `[Accept CSV address]` / `[Accept IO address]` — on a *moved* row, the recorded answer to "which
  side is right"; B3 turns it into the alignment action. Stored as the mode chosen, never as a
  flip; a per-row tick on a row whose answer has become the default shows unticked.

Editing one recomputes the sheet in memory (no file re-read) and re-renders under A2.

**Where.** Core: `IoBridgeRowOverride` + sidecar field (schema token bumped, `LoadNotes` on
migration — the sidecar already has `LoadNotes`). Add-in: columns appended after the status
columns in `IoBridgeColumnLayout`; edit capture in `IoBridgeEditCapture` (ticks are not queued
actions — they are decisions).

**Tests.** Override round-trip through JSON; join honours `Ignore`; migration of an old sidecar.

### B3 · Diff-driven queue for the mechanical intents — item 2 (L)

**What.** Today the queue holds what the user typed or picked and re-anchors it later. The SWP
model recomputes the intended writes on every build and reconciles them into the standing actions.
Port the reconciliation **for the intents the join can compute by itself** — the moved-row
alignments (per B2's accepted side), description refreshes, integrate of `missing` rows the user
opted into — while free-typed cell edits keep the present record-and-re-anchor path. Rules as in
`SwpQueue`: include choices and the first-seen old value survive, satisfied intents retire as
*AppliedExternally*, vanished ones drop, the note carries the evidence.

**Where.** Core: an `IoBridgeIntents` pass after `IoBridgeJoin.Build` producing `BridgeAction`s;
`IoBridgeQueue.Reconcile(intents)` beside `Record`/`Reanchor`. Add-in: builder calls it before
rendering highlights.

**Tests.** Reconcile matrix (new / unchanged / satisfied / vanished / include preserved); moved
row flips side after an override edit.

---

## Phase C — sessions and safety

### C1 · Sessions as files — item 4 (M)

**What.** `<io-list>.iobridge.json` stays the default. Add to the setup form the SWP trio:
*Open…* (any `*.json;*.json.bak` — the `;` separator, never `,`), *Save as…* (writes at once,
switches dialog + running workspace + next Setup to it, carries the whole history), *Save as new…*
(carries settings and file references, starts queue and per-row overrides empty — one IO list
driving two CSV sets, upper and lower sorter, must not share decisions). A session saved beside
its IO list records no path and stays self-locating; one saved elsewhere records the IO list it
belongs to.

**Where.** `IoBridgeSidecarStore` (PathFor → a session file concept, `.bak` recovery, `RootOf`
analogue), `IoBridgeSidecar` (+ `IoListPathRelative`/`IoListPath` rules), `IoBridgeSetupForm`
(session row, three buttons, title line names the session file first), `IoBridgeRegistry` (last
session). Lift what is generic out of `SwpBridgeSidecarStore` rather than copying it.

**Tests.** Self-locating vs recorded-root sessions; Save-as-new leaves queue and overrides empty;
`.bak` opens.

### C2 · Apply safeguards — item 9 (M)

**What.** Two additions, both opt-in at setup:

- **Write to a copy** — the IO list is patched headlessly into `<name>_edited_<stamp>.xlsx` (or a
  sibling folder, mirroring the SWP "version folder" mode) through the OpenXML patcher path the SWP
  Bridge already has (`SwpXlsxPatcher` writes cells by sheet/row/column; the IO Bridge needs only
  literal values, no formulas). Removes the "the IO list is open in Excel — save?" prompt for users
  who want it; COM live-patch stays the default.
- **Held-back groups** — the three address fields of a CSV-side alignment (node/offset/bit, plus
  Card for an off-card move) are one group: a conflicted member holds the others back, noted on
  each, like the SWP's `FLD` group. Today a partial address write is possible.

**Where.** Core: `BridgeAction.Group` (string, serialized, "" today → old sidecars read fine);
`CsvPatchApplier` and `IoBridgeApplyEngine` honour it; a `IoListXlsxPatcher` thin wrapper over the
SWP's stream rewriter. **Tests.** Group hold-back; copy mode leaves the original untouched byte for
byte; re-read verification before write.

### C3 · Coherence checks — item 8 (S)

**What.** Build-time audit legs, reported in `_Log` (A1) and as row verdicts: duplicate IO
addresses on one card sheet; duplicate CSV addresses on one card (the 2026-11-03 Inputs file has at
least one such pair — flag, do not guess); the same name on two CSV records (already "ambiguous,
left alone" in the moved pass — now also logged); a CSV record whose Card names no sheet in the
workspace (today silently dropped by `CollectCsvOnly`). Nothing here blocks a build.

**Where.** Core: `IoBridgeCoherence.Run(model)` after the join. **Tests.** One per leg.

---

## Phase D — rules as data

### D1 · One regex exception on a default, most-specific-first — item 7, second half (S)

**What.** Where B2 introduced a per-row answer and the setup holds a workspace default (the
moved-row preferred side; later the integrate default type per direction), allow one regex over
the **IO-list sheet name** naming the sheets that take the OTHER mode. Resolution: per-row tick →
sheet rule → default. The caption names the mode it applies and flips with the default; the tick on
a governed sheet flips away from the rule's mode; each build reports how many rows the pattern
moved; a pattern that does not parse is refused at setup and inert-with-a-warning at build.

**Where.** Reuse `SwpSheetNameRule` as-is (it is already generic); a `IoBridgeSheetRule<TMode>`
resolved in the intents pass (B3).

### D2 · Name-equivalence table — item 10 (M)

**What.** The scattered name tolerances — `BUS_` prefix, filtered generator keys, and the
zero-padding family found on LOWER SORTER (`STATWORD51` ↔ `STATWORD051`, `MISSINGBLADE_PH_021` ↔
`_21`) — become one table, one rule per row (`_NameRules` sheet, registry-persisted, parsed like
`_Columns`): a regex with capture groups on both sides and a canonical form, e.g.
`^(STATWORD|CTRLWORD|GRUPPOMOTORE)0*(\d+)$ → $1$2`. The join compares canonical keys; the verdict
cell still shows both raw spellings; the **M/N** actions write the OTHER side's raw spelling, never
the canonical one. A rule that does not parse is inert-with-a-warning. Zero-padding is **off by
default** — the CSV name is what SCADA knows the signal by, and a real naming error must not read
clean.

**Where.** Core: `IoBridgeNameRules` used by `IoListLogic.MnemonicKey` through an injected
canonicalizer (keep `MnemonicKey` pure for the pulls). **Tests.** Table parse, canonical join,
raw-spelling writes, inert bad rule.

---

## Phase E — the two new requests

### E1 · Mnemonic syntax proofing against `IO_Defines` (M, plus a spike)

**Source.** `C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\1 Manuali e Specifiche\IO_Defines\`,
all three now OpenXML (converted 2026-11-03; the `.xls` originals stay beside them and are ignored):
`IO_Defines_MSC_v411.xlsx`, `IO_Defines_ISPC_v5.4.1.xlsx`, `IO_Defines_UPSTR_v0_6_6_EN.xlsm`
(macro-enabled — read as a plain package, macros never run), plus `MSC_DPDP_DETAIL.xlsx`
(reference only). Verified on all three: every template tab carries the same four-column table,
located **by header content** — `Nome | Tipo | Parametri | Commento` (MSC, Italian) or
`Name | Type | Parameters | Comment` (ISPC/UPSTR, English), the header being row 3 after a title
and, on ISPC/UPSTR, an `INPUTS` / `OUTPUTS` section row that recurs down the sheet. ISPC adds
`Mandatory` and `Logic` columns, UPSTR a `Logic` one (`NO` / `NC` / "NO if virtual, NC if real").
Tabs where no such header resolves (version history, index, the `STRUCT_*` and `I-O …`
data-exchange layouts, `IO Mirror` whose one template is `IOM_name`) are skipped, not errors.

Names are templates with lowercase placeholder runs — MSC `CABINETALxxxyy`,
`CLASSDIAGaabbbbcccccdd`, `CHUTE_BTN_yy_xxxxx`, `SIRENAxx`; ISPC `RESETx_Zy_z`, `PRERUN_Bx_Zy_z`;
UPSTR `RESET_x_yy_zz`, `LAMPxx_y_zz_l`, `SLIDE_LAMP_x_yyy` — with the widths and ranges stated in
*Parametri/Parameters* ("xxx = Indice Cabinet; Da 0 a 149", "yy = Zone ID, starting from 01",
"x = Line number starting from 1", "yyy = belt ID starting from 0"), plus free-suffix families
(MSC "MARCIA… si può aggiungere liberamente un suffisso"), exact names (`AUTOM`, `MANUAL`) and
one word placeholder (`IOM_name` → `\w+`). Three columns feed checks beyond syntax:

- *Type* — the acronym (DI / DO / AI / AO / VI / VO / VAI / VAO / DEI / DEO / BUS, "DI/VI" =
  either, MSC `DI ⏎ DO` on `EMERGENCY` = both exist). It carries the **direction** too, on all
  three files (MSC has no INPUTS/OUTPUTS sections; its *Tipo* is the only direction source), so
  the direction check is one rule: the template's acronym family vs the row's address prefix
  (`I.`/`Q.`) and the CSV file it sits in. The ISPC/UPSTR section rows are a redundant confirmation.
- The **CSV Type number ↔ acronym** mapping is the controller's own `SIGNAL_TYPE_*` defines
  (received 2026-11-03), now ONE Core table — `CsvSignalType` — read by the picker, the join and
  the proofing: `1 DI · 2 DO · 3 AI · 4 AO · 5 VI · 6 VO · 7 DEI · 8 DEO · 9 AIS · 10 AOS · 11 BUS ·
  13 VAI · 14 VAO`; 12 is skipped so that inputs are odd and an output is its input + 1 (direction
  is parity). The repo had 9/10 as VAI/VAO and no 13/14 — fixed with the table (DONE): AIS/AOS
  stay in the virtual family (every export on file writes them at 0.0.0), VAI/VAO join it. The
  IO_Defines files spell 5/6 as `VI`/`VO`, the tooling as `VDI`/`VDO`; both resolve. `-` (the
  `IOM_name` mirror) is not a type and matches nothing.
- *Logic* `NO` / `NC` / "NO if virtual, NC if real" — checked **first against the CSV's own
  Logic / NC-NO Logic field** (the file the controller runs; `0`/`1` ↔ NO/NC per the existing
  linked-pair convention), and only then against the IO list's *Normal condition*. A CSV that
  contradicts the template is a row verdict (`· logic`), the IO list contradicting the CSV is the
  existing C/D linked-pair difference, and a template disagreement with both is a `_Log` warning
  only — a plant may legitimately differ.

**What.** A per-session, opt-in setting *Proof mnemonics against IO_Defines* (plant type MSC /
ISPC / UPS, one file per type, remembered in the registry, blank = off). Every IO-list mnemonic
and every CSV name is matched against the compiled templates of the chosen type; a name no template
accepts reads `· unknown name` (grey, never an error — the file itself says "se tramite questo
documento non si riesce a determinare il mnemonico, significa che è uno SPECIAL"), a name whose
numeric field is out of range reads `· out of range`, and a type mismatch (template says DO, CSV
says 1) is a `_Log` warning. Nothing is written; it is a read-only check surfaced on the row and
in `_Log`, with the matched template and its *Commento* shown in the preview (B1).

**Compilation rule (to confirm in the spike).** Lowercase letter runs → one numeric field each
(same letter repeated = one field): width fixed `\d{n}` when the run is 2+ letters or the
parameter text gives a zero-padded start ("starting from 01"), free `\d+` for a single letter with
"starting from 1" / "Progressivo (1..X)"; uppercase, digits and `_` literal; a row whose
parameters say suffix-free → `^NAME(_?\w*)?$`; a word placeholder (`name`) → `\w+`; ranges parsed
from "Da A a B" / "A..B" / "0000-9999" / "starting from N" into min(/max) per field. The same
letter may mean different things on different rows (`x` = line on ISPC, = system on UPSTR), so
fields are scoped to the row, never global. Rows that fail to compile are listed in `_Log` as
"template not understood" and skipped — the proofing never blocks a build.

**Spike first.** (1) ~~`.xls` handling~~ — resolved by the conversion; the setting points at the
`.xlsx`/`.xlsm` and the loader refuses a `.xls` with a message naming the conversion. (2) Run the
compiler over all tabs of the three files and count templates understood / not per file — the
acceptance bar is that every name on LOWER SORTER card 1 that is a standard signal matches exactly
one MSC template, and that the ISPC/UPSTR files parse to the same ratio. (3) Confirm
`MSC_DPDP_DETAIL.xlsx` is reference, not a template source. (4) Decide whether the
`INPUTS`/`OUTPUTS` section is reliable enough on every tab to be a verdict or only a `_Log` line.

**Where.** Core: `IoDefinesTable` (load from rows, compile, `Match(name) → template | null`,
`Check(name, csvType)`), kept apart from the SWP's `SwpMnemonicMap` (which decides *destination*,
not *syntax*; the two must not be merged — see CLAUDE.md "Mnemonics are never rewritten").
Add-in: setup rows under the IO LIST section ("the regexes deciding which of its mnemonics
count" is where this belongs), row label suffixes, `_Log` lines. **Tests.** Template compilation
for each shape listed above; range checks; free-suffix families; `.xls` handling per the spike's
decision.

### E3 · Chain reactions — companion signals (L, later; design note now)

**Idea** (the user's term, after the PL5 mechanism): a signal is rarely alone. A parent mnemonic
implies **companions** whose names derive from the parent's by syntax and whose addresses derive
from the parent's by **offset**, and the bridge should use the chain both ways — to **create**
the companions when the parent is integrated, and to **verify** they are present and sit where
the chain says when the parent already exists. Evidence on LOWER SORTER card 1:

- `MODULATOR_DEI_000nn` (DEI, `I.n.1.0`) is always accompanied by `MODULATOR_IOL_ERR_000nn` (DI)
  at **offset +32, bit 7** (`I.35.1.0` → `I.35.33.7`), and by `MODULATOR_DEO_000nn` in Outputs at
  `Q.n.0.0`. The 69 "moved" rows moved **as chains** — the parent shifted by +3 and every
  companion with it — which is why a chain check would have flagged the whole block at once.
- `EMERGENCYxxxyy` (DI) must have a DO of the **same name** for the mushroom lamp
  (`IO_Defines_MSC` *Diagnostic*, row 10: "Creare un DO, pari nome").
- `CHUTE_FLAP_O_xxxxx` / `CHUTE_FLAP_C_xxxxx` exist as a DI pair and as a DO pair;
  `CTRLWORDnn` / `STATWORDnn` / `GRUPPOMOTOREnn` / `LTI_ACTFEEDBACK_nn` form one drive quad at
  fixed offsets on the same node.

**Shape.** A `_Chains` sheet (registry-persisted, parsed like `_Columns`, one rule per row,
inert-with-a-warning when it does not parse — the same discipline as `SwpMnemonicMap`): parent
regex with capture groups → a list of companions, each `name template` (built from the captures,
e.g. `MODULATOR_IOL_ERR_$1`) + `side` (I/O) + `type` (CSV type number) + `offset delta` + `bit`
(absolute or `+n`) + `required | optional`. Offsets are relative to the parent, never absolute, so
one rule covers every node. Rules may chain (a companion can itself be a parent) with a depth cap.

**Verification** (first to ship): every parent row on a card sheet gets its companions resolved
against the CSV; a required companion that is absent reads `· chain missing <name>`, one present
at another address than the rule predicts reads `· chain offset` with the actual delta, and a
companion present with no parent is a `_Log` line. Shown on the parent row and in the preview (B1)
as the chain tree. **Creation** (second): integrating a parent (I) offers to queue the required
companions too — one `CreateCsvRow` each, addresses computed from the rule, descriptions composed
as for any create — as one queue group (C2's `Group`), applied together or not at all. Aligning a
moved parent (A/B) offers to carry its companions along by the same delta. **Starting set**: the
IO_Defines rows that state a companion in their comment (`EMERGENCY`, the `CHUTE_FLAP` pairs, the
ISPC `_z` "z = 0 must always exist and virtual" families) seed the default `_Chains`; the modulator
and drive chains are added from the field evidence above once confirmed against the MSC package
documentation. Depends on B1 (preview), B3 (intents) and C2 (groups); nothing in it changes the
join's claiming, which stays address-first then name.

### E2 · One browse for the CSV pair (S)

**What.** On the setup form, a single *Browse…* beside the CSV rows opens a multi-select dialog
filtered `Plant CSV|*.csv` (the `;` syntax is not involved — one pattern) in the last CSV folder;
picking one or two files fills Inputs/Outputs by **content** (`CsvColumnsMap.Locate` with each
kind — the file whose header resolves as Inputs goes to Inputs, likewise Outputs; by name
`Plant_Inputs` / `Plant_Outputs` only as the tiebreak when content is ambiguous). Picking a single
file fills its slot and, when the sibling with the other kind's name exists in the same folder,
offers it in the other box. The two text boxes stay editable; the two per-row *Browse…* buttons
stay for the odd case (CSVs in different folders). Content before names, per the plant-CSV
recognition rule.

**Where.** `IoBridgeSetupForm` — `BrowseCsvPair()`; one small Core helper
`PlantCsvKindDetector.Detect(path) → Inputs | Outputs | Unknown` (wraps `Locate` with both kinds)
so the decision is unit-tested. **Tests.** Detect on the two real export schemas and on a
BusField.csv (Unknown). **Docs.** One sentence in "Setup".

---

## Order and dependencies

```
A1 Log ─┬─ A4 regex counts ─┐
A2 AutoFilter ──────────────┼─ B1 Preview ─ B2 Overrides ─ B3 Intents ─ D1 Sheet rule
A3 Tabs/borders ────────────┘                     │
E2 CSV pair browse (independent, any time)        └─ C1 Sessions ─ C2 Apply safeguards
C3 Coherence (after A1)                                   D2 Name rules (after A1; B1 shows them)
E1 IO_Defines proofing (spike after A1; feature after B1 so the preview can show the template)
E3 Chain reactions (after B1 + B3 + C2; verification first, creation second)
```

Suggested first batch: **A1 + A2 + A3 + E2** (one session each, all independent, all land without
touching join or apply). Then **B1**, which everything user-facing afterwards leans on. **B3** is
the only item that changes what the queue *means* and should ship alone, with the FVT files run
end-to-end (build → queue → apply to copies → rebuild) before it is merged.

## Out of scope (SWP-specific, not ported)

CONFIG_PLANT text modes, description formulas and the `_xlfn.`/`_xlpm.` spelling, glossary and
dictionaries, FLD append/overwrite, `_Offsets`, strobe templates, the formula cleaner, the
production-tree checks, version folders with copy-along (C2 takes only the "write to a copy" idea).
