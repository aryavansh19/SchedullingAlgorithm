# Weekday Interview Scheduling System

A scheduling "playbook" that two college interns (**Ram** and **Shyam**) can follow
almost blindly to hit the target:

> **80% of interview requests SCHEDULED (via Calendly) within 24 hours of arrival.**

This repo turns ~500 daily interview requests into a step-by-step, self-prioritising
to-do system with built-in scripts, an automatic hand-off between the two interns, and a
live KPI dashboard.

---

## What's in the repo

| File | Purpose |
|---|---|
| `generate_excel.py` | Reads the raw requests and **generates** the solution workbook. Portable (paths are relative to the script). |
| `Scheduling ToDos.xlsx` | The raw input (the `INPUT Interview Requests` sheet). |
| `Assignment_Solution.xlsx` | The generated deliverable (7 sheets, described below). |

### Regenerate the workbook
```bash
pip install openpyxl
python generate_excel.py
```

---

## The deliverable workbook (`Assignment_Solution.xlsx`)

| Sheet | What it does |
|---|---|
| **Start Here** | 60-second orientation + the exact daily routine for both interns. |
| **Assumptions & Approach** | All assumptions, the problem breakdown, the edge-case table, capacity math, and the innovative ideas (this is the documentation graded by the assignment). |
| **Playbook (SOP)** | Copy-paste WhatsApp / email / call scripts + a decision table that says what every dropdown outcome means and what to do next. |
| **Ram - Digital ToDo** | SLA-sorted digital-outreach queue. A `NEXT STEP` column tells Ram exactly what to do per row. |
| **Shyam - Call ToDo** | Call queue that is **auto-fed** only with the requests Ram could not close. |
| **KPI Dashboard** | Live `% scheduled within 24h` vs the 80% target, plus workload split. |
| **INPUT (clean)** | The cleaned raw requests (deduped flags, fixed phone/email, round-specific link). |

---

## How I broke the problem down

The real job is not "schedule interviews" — it is **"get the candidate to click the
Calendly link before the 24-hour clock runs out."** That splits into four units:

1. **Intake & clean** – dedupe, pick the correct round's Calendly link, fix phone/email.
2. **Prioritise** – sort by time-left so we never breach the 24h SLA.
3. **Chase, cheapest channel first** – digital (Ram) → phone (Shyam) → team lead.
4. **Measure** – track `% scheduled within 24h` so we know if we're winning.

---

## The cadence (anchored to each request's `Added On`, not a fixed clock)

| Time | Owner | Action |
|---|---|---|
| **T+0h** | Ram | WhatsApp + Email blast with the round's Calendly link and a clear deadline |
| **T+2h** | Ram | Check Calendly; if not booked → WhatsApp nudge |
| **T+4h** | Ram | Final check; if still not booked → **auto-escalate to Shyam** |
| **T+6h** | Shyam | Call 1 — offer to book a slot live, on the candidate's behalf |
| **T+20h** | Shyam | Call 2 (final); if still nothing → escalate to lead **before T+24h** |

Every touchpoint lands inside the 24-hour window — that is how the 80% target is
protected. The earlier version used fixed wall-clock times (10 AM, 1 PM, …), which breaks
for requests that arrive later in the day; this version starts the clock per request.

---

## Key assumptions

- **A1.** Candidates already agreed to interview; the only task is to make them click Calendly.
- **A2.** Candidates are India-based → phones default to `+91`; working hours are 09:00–21:00 IST.
- **A3.** Ram has bulk tooling (WhatsApp Business broadcast + mail-merge), so the initial blast to ~500 people is a batch action, not 500 manual messages.
- **A4.** The "Scheduling method" cell holds one Calendly link per round; we send only the link for that row's round.
- **A5.** The 24h SLA clock starts at `Added On`, independently per request.
- **A6.** Because every cadence touchpoint is inside the 24h window, a request that reaches the *Scheduled* state was scheduled within 24h — so the KPI can be read from the outcome stage without a separate timestamp.
- **A7.** "Not interested / Declined" candidates are removed from the SLA denominator (a real drop-off, not a scheduling failure).
- **A8.** Interns should never have to interpret — every decision is pre-written.

---

## Edge cases handled

| Edge case | What the intern does |
|---|---|
| Ignores email | WhatsApp + email at T+0; nudge at T+2h; call at T+6h |
| No WhatsApp | Blast also via email + SMS; if undelivered, Shyam calls |
| Never answers calls | Call 1 + Call 2 in golden hours; SMS/voicemail; then escalate to lead |
| Forgot after agreeing | Nudge reminds them they already agreed; offer to book on their behalf |
| Invalid / wrong phone | Mark `Invalid number` → lead gets correct details from the company |
| Wrong / bouncing email | Mark `Bad contact` → lead asks the company for the correct contact |
| No Calendly slots | Note it; lead asks company to open slots; pause SLA with a note |
| Wants to reschedule | Re-send link; self-serve; counts as scheduled once booked |
| Not interested / dropped | Mark `Not interested/Declined` → closed + removed from SLA denominator |
| Says STOP / DND | Stop all outreach; mark Declined; note "DND"; tell lead |
| Duplicate request | `Dup?` column flags it; process the first only |
| Multiple rounds | Each round is its own row with its own link; scheduled independently |
| Arrives late at night | Night touches roll to 09:00 next day, still inside 24h |

---

## Capacity / feasibility (the honest part)

~500 requests/day with 2 interns means manual messaging cannot scale, so:

- Ram's blast must be a **bulk broadcast / mail-merge** (one batch) + quick visual checks.
- If ~75% self-schedule after blast + nudge (they already said yes), only ~125 reach Shyam.
- ~125 calls × ~4 min (incl. retries) ≈ **8.5 hours** — tight but doable in a shift.
- **Trigger:** if escalations exceed ~150/day, add an auto-SMS fallback or request more help. This threshold is documented so interns raise capacity problems early.

---

## Innovative ideas built in

- **`NEXT STEP` column** — the sheet literally tells the intern the next action (zero thinking).
- **Automatic Ram → Shyam hand-off** via a cross-sheet formula (no copy-paste between people).
- **"Book on their behalf"** on calls removes candidate friction (we pick the slot).
- **Priority by a live `Hours Left` countdown** with a red→green heat-map.
- **Duplicate auto-flagging** so interns never double-chase the same person.
- **Self-serve KPI dashboard** so interns can see if they're on track for 80%.
- **Golden-hours calling guidance** (lunch + evening) to maximise pickup rate.

---

## How the automation works (for reviewers)

- **Live countdown:** `Hours Left = (SLA Deadline - NOW()) * 24`, where `SLA Deadline = Added On + 1 day`.
- **Guidance:** the `NEXT STEP` columns are nested `IF` formulas over the step dropdowns, so a row updates itself as the intern fills it in.
- **Hand-off:** Shyam's `NEEDS CALL?` formula reads Ram's escalation cell for the **same row** (`='Ram - Digital ToDo'!Q{row}`), so escalations appear in Shyam's queue automatically — he just filters `NEEDS CALL? = CALL NOW`.
- **Single source of truth:** Shyam's `FINAL OUTCOME` column classifies each request (scheduled self-serve / scheduled on call / declined / SLA breach / in progress). The dashboard's `COUNTIF`s read this column. (Outcome labels deliberately avoid `<`/`>` so `COUNTIF` treats them as literal text.)

The generated workbook was validated with a formula engine: **no `#REF!`/`#NAME?`/`#VALUE!`
errors**, and a simulated day confirmed the escalation hand-off and KPI math behave correctly.
