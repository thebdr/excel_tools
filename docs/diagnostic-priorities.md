# Diagnostic data priorities — the FLD rule

*Statement of principles for the SWP / diagnostic tooling. Recorded 2026-08-21 from the field case
`CABINETAL01914` (a safety door reported under a block naming sixteen doors). This is the reference
every SWP Bridge write path is measured against; `CLAUDE.md` carries the short form.*

## 1. Context

The SWP is the diagnosis data SCADA shows to the people who keep the plant running. When an alarm
fires, the maintainer on site has **exactly two things**: the SCADA message and the
troubleshooting (T.S.) manual. Nothing else. Whatever those two do not say, the maintainer does not
know — and a plant can be giant.

## 2. Definitions

| Term | Meaning |
|---|---|
| **Functional Unit (FU)** | A system identifier inside a building. |
| **Location** | A system identifier inside a Functional Unit. |
| **Device** | A real, labeled hardware piece. The physical label carries **only** the Device. |

Together: the **FLD**. The maintainer reaches the Location by following the coordinates, then needs
the Device to find the one piece of hardware that triggered the alarm.

## 3. The rule

> **The FLD must always be carried.** Either in the row's own cells (FU / Location / Device — the
> SWP's H / I / J) or appended to the description text. Never neither.

Concretely, on a mismatch between the IO list and the SWP block identity:

- **Not "brutal overwrite"** (append mode): a mismatch of *either* member of the trio appends the
  compacted identity (`=S1+SG14-B1`) to the descriptions — **always**, no exceptions by row type.
- **Brutal overwrite**: the trio lands in the row's own cells; the descriptions stay as they are.
- A row that cannot be overwritten (shared formulas) falls back to the append. The FLD is never
  dropped because a mechanism was unavailable.

## 4. Priority order

1. **Pointing to the right place and the right device.** This is the main, most important
   information the diagnostics carry. Everything else comes at lower priority.
2. **Description accuracy — secondary.** A wrong description ("circuit breaker tripped" for what
   is really a Profinet fault) is acceptable **if the device pointed at is the right one**. Once a
   person stands in front of the hardware, the description no longer matters: red LEDs or no LEDs,
   tripped switches — they know what to do once they have found it.
3. **Cosmetics and convenience** — last.

## 5. Why the order is absolute

Pointing the maintainer to a **wrong place, or to no place**, triggers instant frustration and a
chain of consequences:

- the maintenance team cannot solve it on site → a call to the service team;
- the service loop spans **days or weeks**, with a chain of calls and blaming in between;
- the customer's final, experienced reaction: **choose another supplier for the next projects, or
  take legal action.**

The difference between a maintenance team reaching the problem on time and that chain is, in most
cases, whether the FLD was carried.

## 6. Corollaries already applied in the tooling

- **T.S. reference is mandatory** on IO-list alarms, and an existing one is never blanked — the
  manual is the maintainer's second and last source; a wrong section still beats no section.
- **Mismatches always warn**; a deliberate one is silenced per row (`[Mismatch OK]`) — the
  treatment (append or overwrite) still applies, because the FLD must still be carried.
- **Virtual twins are surfaced, never silently accepted**: a virtual CSV record carrying a real
  signal's name looks present but acquires nothing from the field — an alarm that never fires sends
  nobody anywhere.
- **Generator keys are filtered out of comparisons**, never written over the list's conventions.
- **Descriptions derive from the glossary and the IO list**; they are kept accurate, but they are
  the second priority, never traded against the first.
