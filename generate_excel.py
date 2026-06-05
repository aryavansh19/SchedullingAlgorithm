"""
Weekday Interview Scheduling System - Workbook Generator
========================================================

Reads the raw interview requests from "Scheduling ToDos.xlsx" (INPUT sheet) and
produces "Assignment_Solution.xlsx": a fully documented, intern-proof scheduling
playbook that two college interns (Ram and Shyam) can follow almost blindly to
hit the goal of "80% of interviews scheduled within 24 hours".

The workbook contains:
  1. Start Here            - 60-second orientation + daily routine
  2. Assumptions & Approach- documented assumptions, problem breakdown, edge
                             cases, capacity math and innovative ideas
  3. Playbook (SOP)        - exact message/email/call scripts + decision table
  4. Ram - Digital ToDo    - SLA-sorted digital outreach queue (WhatsApp+Email)
  5. Shyam - Call ToDo     - auto-fed call queue for escalations only
  6. KPI Dashboard         - live % scheduled within 24h vs the 80% target
  7. INPUT (clean)         - cleaned raw requests

Design highlights (these fix the flaws of the previous version):
  - Cadence is anchored to each request's own "Added On" timestamp, not a fixed
    wall clock, so the 24h SLA is meaningful per request.
  - Queues are sorted oldest-first (closest to breach first) with a live
    "Hours Left" countdown and a "NEXT STEP" instruction computed by formula.
  - Ram -> Shyam hand-off is automatic: Shyam's "NEEDS CALL?" column reads
    Ram's escalation result for the same request via a cross-sheet formula.
  - The correct round-specific Calendly link and the Round are surfaced.
  - Phone numbers are kept as text (no lost digits), headers are trimmed, and
    likely duplicates are flagged.
  - Paths are relative to this script, so it runs anywhere.
"""

import os
import re
import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Config / constants
# --------------------------------------------------------------------------- #
HERE = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(HERE, "Scheduling ToDos.xlsx")
OUTPUT_FILE = os.path.join(HERE, "Assignment_Solution.xlsx")

RAM_SHEET = "Ram - Digital ToDo"
SHYAM_SHEET = "Shyam - Call ToDo"

DEFAULT_COUNTRY_CODE = "+91"  # Assumption: candidates are India-based (see docs)

# --- Simulation clock -------------------------------------------------------
# The sample data is dated 2 Nov 2022. Instead of counting down from the real
# system date (which would make "Hours Left" a huge negative number), the whole
# workbook counts down from this editable "Treat now as" clock cell. Default it
# to the evening of 2 Nov 2022, i.e. "the batch just arrived, start scheduling".
# A reviewer can edit the cell (KPI Dashboard!G2) to advance time and watch the
# urgency / heat-map change.
CLOCK_DEFAULT = datetime.datetime(2022, 11, 2, 20, 0)
CLOCK_CELL = "'KPI Dashboard'!$G$2"

# Palette
NAVY = "1F3864"
BLUE = "2E5496"
GREEN = "548235"
RED = "C00000"
AMBER = "BF8F00"
GREY = "404040"
LIGHT = "D9E1F2"
LIGHT_GREEN = "E2EFDA"
LIGHT_RED = "FCE4D6"
WHITE = "FFFFFF"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# --------------------------------------------------------------------------- #
# Small styling helpers
# --------------------------------------------------------------------------- #
def style_header(cell, fill=BLUE, color=WHITE, size=11, wrap=True):
    cell.font = Font(bold=True, color=color, size=size)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = BORDER


def title_cell(ws, text, fill=NAVY, size=14):
    c = ws.cell(row=ws.max_row + 1 if ws.max_row > 1 or ws["A1"].value else 1, column=1)
    c.value = text
    c.font = Font(bold=True, color=WHITE, size=size)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal="left", vertical="center")
    return c


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def parse_calendly_links(blob):
    """Turn the multi-line 'Round1: <link>\\nRound2: <link>...' blob into a dict
    keyed by a normalised round name e.g. {'round1': '<link>', ...}."""
    links = {}
    if not blob:
        return links
    for line in str(blob).splitlines():
        if ":" not in line:
            continue
        key, val = line.split(":", 1)
        norm = re.sub(r"\s+", "", key).lower()  # "Round 3" -> "round3"
        links[norm] = val.strip()
    return links


def link_for_round(blob, round_value):
    links = parse_calendly_links(blob)
    norm = re.sub(r"\s+", "", str(round_value or "")).lower()
    return links.get(norm, str(blob or "").replace("\n", " | "))


def clean_phone(raw):
    if raw is None or str(raw).strip() == "":
        return ""
    digits = re.sub(r"\D", "", str(raw).split(".")[0])
    if len(digits) == 10:
        return f"{DEFAULT_COUNTRY_CODE} {digits}"
    if digits:
        return f"+{digits}" if not str(raw).strip().startswith("+") else str(raw).strip()
    return str(raw).strip()


# --------------------------------------------------------------------------- #
# Read + clean the raw input
# --------------------------------------------------------------------------- #
def load_requests():
    if not os.path.exists(INPUT_FILE):
        # Minimal dummy data so the script still runs end-to-end.
        return [
            {
                "company": "Google", "interviewer": "Sundar", "interviewer_email": "s@google.com",
                "candidate": "John Doe", "email": "john@gmail.com", "phone": "9999999999",
                "blob": "Round1: <calendly_link>\nRound2: <calendly_link>", "round": "Round1",
                "added_on": None,
            }
        ]

    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    ws = wb["INPUT Interview Requests"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    requests = []
    seen = set()
    for r in rows:
        if r is None or r[0] in (None, ""):
            continue
        company, interviewer, interviewer_email, candidate, email, phone, blob, rnd, added_on = (
            list(r) + [None] * 9
        )[:9]
        key = (
            str(company).strip().lower(),
            str(candidate).strip().lower(),
            str(rnd).strip().lower(),
            str(email).strip().lower(),
        )
        dup = "Yes" if key in seen else ""
        seen.add(key)
        requests.append({
            "company": company,
            "interviewer": interviewer,
            "interviewer_email": interviewer_email,
            "candidate": candidate,
            "email": (str(email).strip() if email else ""),
            "phone": clean_phone(phone),
            "blob": blob,
            "round": rnd,
            "added_on": added_on,
            "calendly": link_for_round(blob, rnd),
            "dup": dup,
        })

    # Prioritise oldest requests first (closest to the 24h SLA breach).
    requests.sort(key=lambda x: (x["added_on"] is None, x["added_on"]))
    return requests


# --------------------------------------------------------------------------- #
# Sheet 1: Start Here
# --------------------------------------------------------------------------- #
def build_start_here(wb):
    ws = wb.active
    ws.title = "Start Here"
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 102})

    lines = [
        ("Weekday Interview Scheduling System", "title"),
        ("Goal: 80% of interview requests SCHEDULED (via Calendly) within 24 hours of arrival.", "sub"),
        ("", "gap"),
        ("WHO DOES WHAT", "h"),
        ("Ram  = Digital outreach for EVERY new request (WhatsApp + Email + nudges).", "p"),
        ("Shyam = Phone calls, but ONLY for requests Ram could not schedule (auto-escalated).", "p"),
        ("You (Team Lead / AM) = handle the rare cases both interns escalate.", "p"),
        ("", "gap"),
        ("YOUR DAILY ROUTINE (do this top to bottom, every day)", "h"),
        ("1. Open your tab: Ram -> 'Ram - Digital ToDo'.  Shyam -> 'Shyam - Call ToDo'.", "p"),
        ("2. Sort your tab by 'Hours Left' (smallest first). Smallest = most urgent.", "p"),
        ("3. Read the 'NEXT STEP' column. It tells you EXACTLY what to do for that row.", "p"),
        ("4. Do it, then fill the dropdown for that step. The row updates itself.", "p"),
        ("5. Repeat down the list until every row says DONE, ESCALATED, DECLINED or CLOSED.", "p"),
        ("", "gap"),
        ("THE GOLDEN RULES", "h"),
        ("- Never guess. Every message/email/call script is in the 'Playbook (SOP)' tab.", "p"),
        ("- Always work the row with the FEWEST hours left first.", "p"),
        ("- If a dropdown does not have the outcome you saw, write it in 'Notes' and ask the lead.", "p"),
        ("- Shyam only touches a row when its 'NEEDS CALL?' column says CALL NOW.", "p"),
        ("- Check the 'KPI Dashboard' tab once in the morning and once in the evening.", "p"),
        ("", "gap"),
        ("TABS IN THIS WORKBOOK", "h"),
        ("Start Here  |  Assumptions & Approach  |  Playbook (SOP)  |  Ram - Digital ToDo  |  "
         "Shyam - Call ToDo  |  KPI Dashboard  |  INPUT (clean)", "p"),
    ]
    _write_doc_lines(ws, lines)
    return ws


def _write_doc_lines(ws, lines, start_row=1):
    r = start_row
    for text, kind in lines:
        cell = ws.cell(row=r, column=2, value=text)
        if kind == "title":
            cell.font = Font(bold=True, size=18, color=NAVY)
        elif kind == "sub":
            cell.font = Font(bold=True, size=12, color=GREEN)
            cell.alignment = Alignment(wrap_text=True)
        elif kind == "h":
            cell.font = Font(bold=True, size=12, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=BLUE)
            cell.alignment = Alignment(vertical="center")
            ws.row_dimensions[r].height = 20
        elif kind == "p":
            cell.font = Font(size=11, color=GREY)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        elif kind == "gap":
            pass
        r += 1
    return r


# --------------------------------------------------------------------------- #
# Sheet 2: Assumptions & Approach
# --------------------------------------------------------------------------- #
def build_assumptions(wb):
    ws = wb.create_sheet("Assumptions & Approach")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 104})

    lines = [
        ("Assumptions, Thought Process & Approach", "title"),
        ("This documentation is part of the deliverable. It explains every assumption and why the "
         "system is built the way it is.", "sub"),
        ("", "gap"),

        ("1. HOW I BROKE THE PROBLEM DOWN", "h"),
        ("The job is not 'schedule interviews' - it is 'get the candidate to click the Calendly link "
         "before the 24h clock runs out'. So the problem splits into 4 units:", "p"),
        ("  a) Intake & clean: dedupe, pick the right round's link, fix phone/email.", "p"),
        ("  b) Prioritise: sort by time-left so we never breach the 24h SLA.", "p"),
        ("  c) Chase via cheapest channel first: digital (Ram) -> phone (Shyam) -> lead.", "p"),
        ("  d) Measure: track % scheduled within 24h so we know if we are winning.", "p"),
        ("", "gap"),

        ("2. KEY ASSUMPTIONS", "h"),
        ("A1. Candidates already agreed to interview, so the only task is to make them click Calendly.", "p"),
        ("A2. Candidates are India-based; phones default to +91 and 'working hours' are 9:00-21:00 IST.", "p"),
        ("A3. Ram has bulk tooling (WhatsApp Business broadcast + mail-merge) so the initial blast to "
         "~500 people is a batch action, not 500 manual messages.", "p"),
        ("A4. The Calendly 'Scheduling method' cell holds one link per round; we send only the link for "
         "the round on that row.", "p"),
        ("A5. The 24h SLA clock starts at 'Added On' for each request, independently.", "p"),
        ("A6. Because every touchpoint in the cadence is scheduled inside the 24h window, a request that "
         "reaches the 'Scheduled' state was, by construction, scheduled within 24h. So the KPI can be "
         "read from the outcome stage without a separate timestamp.", "p"),
        ("A7. 'Not interested / Declined' candidates are removed from the SLA denominator (they are a "
         "real-world drop, not a scheduling failure).", "p"),
        ("A8. Interns can read but should not have to interpret - every decision is pre-written.", "p"),
        ("", "gap"),

        ("3. THE CADENCE (anchored to each request's 'Added On')", "h"),
        ("T+0h   Ram: WhatsApp + Email blast with the round's Calendly link + clear deadline.", "p"),
        ("T+2h   Ram: check Calendly. If not booked -> WhatsApp nudge.", "p"),
        ("T+4h   Ram: final check. If still not booked -> auto-escalate to Shyam's call queue.", "p"),
        ("T+6h   Shyam: Call 1. Offer to book 2-3 slots live, on the candidate's behalf.", "p"),
        ("T+20h  Shyam: Call 2 (final). If still nothing -> escalate to the lead before T+24h.", "p"),
        ("Every step lands inside 24h, which is how we protect the 80% target.", "p"),
        ("", "gap"),

        ("4. EDGE CASES AND HOW THE SHEET HANDLES THEM", "h"),
        ("(full table is below this section)", "p"),
        ("", "gap"),

        ("5. CAPACITY / FEASIBILITY MATH (the honest part)", "h"),
        ("~500 requests/day, 2 interns. Manual messaging cannot scale, so:", "p"),
        ("- Ram's blast must be a bulk broadcast/mail-merge (1 batch) + quick visual checks.", "p"),
        ("- If ~75% self-schedule after blast+nudge (they already said yes), only ~125 reach Shyam.", "p"),
        ("- 125 calls at ~4 min each (incl. retries) = ~500 min = ~8.5h. Tight but possible in a shift.", "p"),
        ("- TRIGGER: if escalations exceed ~150/day, add an auto-SMS fallback or request more help. "
         "This threshold is called out so interns escalate capacity problems early.", "p"),
        ("", "gap"),

        ("6. INNOVATIVE IDEAS BUILT IN", "h"),
        ("- 'NEXT STEP' column: the sheet literally tells the intern the next action - zero thinking.", "p"),
        ("- Automatic Ram->Shyam hand-off via formula (no copy-paste between people).", "p"),
        ("- 'Book on their behalf' on calls removes candidate friction (we pick the slot for them).", "p"),
        ("- Priority by live 'Hours Left' countdown with colour heat-map.", "p"),
        ("- Duplicate auto-flagging so interns never double-chase the same person.", "p"),
        ("- Self-serve KPI dashboard so interns see if they are on track for 80%.", "p"),
        ("- 'Golden hours' calling guidance (lunch + evening) to maximise pickup rate.", "p"),
    ]
    next_row = _write_doc_lines(ws, lines)

    # Edge-case table
    next_row += 1
    ws.cell(row=next_row, column=2, value="EDGE-CASE PLAYBOOK").font = Font(bold=True, size=12, color=WHITE)
    ws.cell(row=next_row, column=2).fill = PatternFill("solid", fgColor=BLUE)
    next_row += 1

    edge = [
        ("Edge case", "What the intern does"),
        ("Candidate ignores email", "WhatsApp + email are sent together at T+0; nudge at T+2h; call at T+6h."),
        ("Candidate has no WhatsApp", "Blast also goes by email + SMS; if no WhatsApp delivery, Shyam calls."),
        ("Candidate never answers calls", "Call 1 + Call 2 in golden hours; leave SMS/voicemail; then escalate to lead."),
        ("Candidate forgot after agreeing", "Nudge script reminds them they already agreed; offer to book on their behalf."),
        ("Invalid / wrong phone number", "Mark 'Invalid number' -> goes straight to lead to get correct details from company."),
        ("Wrong / bouncing email", "Mark 'Bad contact' on blast -> lead asks the company for correct contact."),
        ("No Calendly slots / interviewer busy", "Note it; lead pings the company to open slots; SLA clock paused with a note."),
        ("Candidate wants to reschedule", "Send link again; they self-serve; counts as scheduled once booked."),
        ("Candidate not interested / dropped", "Mark 'Not interested/Declined' -> closed + removed from SLA denominator."),
        ("Candidate says STOP / do not contact", "Stop all outreach immediately; mark Declined; note 'DND'; tell lead."),
        ("Duplicate request for same person", "'Dup?' column flags it; process only the first, ignore the rest."),
        ("Multiple rounds for same candidate", "Each round is its own row with its own link; scheduled independently."),
        ("Request arrives late at night", "Cadence respects 9-21 working hours; night touches roll to 9:00 next day, still < 24h."),
    ]
    for i, (a, b) in enumerate(edge):
        ra = ws.cell(row=next_row, column=2, value=a)
        # Put the two columns side by side using a merged-ish layout: col B = case, col C = action.
        # Re-add width for column C so it is readable.
        ws.column_dimensions["C"].width = 70
        rb = ws.cell(row=next_row, column=3, value=b)
        if i == 0:
            style_header(ra, fill=GREY)
            style_header(rb, fill=GREY)
        else:
            ra.font = Font(bold=True, size=10, color=GREY)
            ra.alignment = Alignment(wrap_text=True, vertical="top")
            ra.border = BORDER
            rb.font = Font(size=10, color=GREY)
            rb.alignment = Alignment(wrap_text=True, vertical="top")
            rb.border = BORDER
        next_row += 1
    return ws


# --------------------------------------------------------------------------- #
# Sheet 3: Playbook (SOP)
# --------------------------------------------------------------------------- #
def build_playbook(wb):
    ws = wb.create_sheet("Playbook (SOP)")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 26, "C": 96})

    def section(title):
        r = ws.max_row + 1
        c = ws.cell(row=r, column=2, value=title)
        c.font = Font(bold=True, size=12, color=WHITE)
        c.fill = PatternFill("solid", fgColor=BLUE)
        ws.cell(row=r, column=3).fill = PatternFill("solid", fgColor=BLUE)
        return r

    def kv(code, text):
        r = ws.max_row + 1
        a = ws.cell(row=r, column=2, value=code)
        a.font = Font(bold=True, size=10, color=NAVY)
        a.alignment = Alignment(vertical="top", wrap_text=True)
        b = ws.cell(row=r, column=3, value=text)
        b.font = Font(size=10, color=GREY)
        b.alignment = Alignment(wrap_text=True, vertical="top")

    t = ws.cell(row=1, column=2, value="Playbook (Standard Operating Procedure)")
    t.font = Font(bold=True, size=18, color=NAVY)
    ws.cell(row=2, column=2, value="Copy-paste these. Replace [BRACKETS] with the values from the row you are working.").font = Font(
        italic=True, size=11, color=GREEN)

    section("MESSAGE SCRIPTS (Ram)")
    kv("M1 - Initial WhatsApp", "Hi [Candidate], this is [Your name] from Weekday. Your [Round] interview "
       "with [Company] is confirmed. Please pick a time here: [Calendly Link]. It takes 30 seconds. "
       "Kindly book within today. Reply here if you face any issue.")
    kv("M1 - Initial Email", "Subject: Action needed: book your [Company] [Round] interview\n\nHi [Candidate], "
       "you are confirmed for the [Round] interview with [Company]. Please choose a slot here: "
       "[Calendly Link]. Booking today keeps your process on track. Reply if the link does not work.")
    kv("M2 - Nudge WhatsApp (T+2h)", "Hi [Candidate], gentle reminder to lock your [Company] interview slot: "
       "[Calendly Link]. Takes under a minute. Want me to suggest a time?")
    kv("M3 - SMS fallback", "Weekday: Please book your [Company] [Round] interview: [Calendly Link]. "
       "Reply CALL if you want us to book for you.")

    section("CALL SCRIPTS (Shyam)")
    kv("C1 - Call 1 opener", "Hi [Candidate], calling from Weekday about your confirmed [Company] [Round] "
       "interview. I can text you the booking link now, or I can book a slot for you right now - "
       "I have [slot A], [slot B], [slot C]. Which works?")
    kv("C1 - If they pick a slot", "Book it on Calendly yourself on their behalf, confirm by WhatsApp, "
       "then mark 'Booked on their behalf'.")
    kv("C2 - Call 2 (final)", "Hi [Candidate], last reminder so you do not lose your [Company] slot. "
       "Shall I book [slot A] or [slot B] for you right now?")
    kv("VM - Voicemail / no answer", "Send M3 SMS immediately after a missed call, then try again in the "
       "next golden hour (12:00-14:00 or 18:00-20:00).")

    section("DECISION TABLE - what each dropdown outcome means and what happens next")
    dt = [
        ("Where", "You select", "What it means / next action"),
        ("Ram S1 Blast", "Sent", "Blast delivered. Wait for T+2h check."),
        ("Ram S1 Blast", "Bad contact", "Email bounced AND no WhatsApp. Row blocks -> lead gets correct contact."),
        ("Ram S2 Check", "Scheduled", "Candidate booked. Row is DONE. Stop."),
        ("Ram S2 Check", "Not yet", "Not booked. Send nudge M2 (S3)."),
        ("Ram S3 Nudge", "Nudged", "Reminder sent. Wait for T+4h final check."),
        ("Ram S4 Final", "Scheduled", "Booked. Row is DONE."),
        ("Ram S4 Final", "Escalate to call", "Auto-sends row to Shyam's call queue."),
        ("Shyam Call 1", "Scheduled on call", "They booked while on call. DONE."),
        ("Shyam Call 1", "Booked on their behalf", "You booked for them. DONE. Confirm by WhatsApp."),
        ("Shyam Call 1", "No answer", "Send SMS, retry in next golden hour -> Call 2."),
        ("Shyam Call 1", "Invalid number", "Wrong number -> escalate to lead for correct details."),
        ("Shyam Call 1", "Call later", "They asked to be called later -> Call 2 slot."),
        ("Shyam Call 1", "Not interested/Declined", "Candidate dropped. CLOSED + tell lead."),
        ("Shyam Call 2", "Scheduled on call / Booked on their behalf", "DONE."),
        ("Shyam Call 2", "Still no answer", "Escalate to AM/lead before T+24h."),
        ("Shyam Call 2", "Escalate to AM", "Lead takes over the last-mile attempt."),
    ]
    for i, row in enumerate(dt):
        r = ws.max_row + 1
        cells = [ws.cell(row=r, column=2, value=row[0]),
                 ws.cell(row=r, column=3, value=row[1] + "  ->  " + row[2] if i else row[1] + "  |  " + row[2])]
        if i == 0:
            ws.cell(row=r, column=2).value = row[0]
            ws.cell(row=r, column=3).value = f"{row[1]}  |  {row[2]}"
            for c in cells:
                style_header(c, fill=GREY)
        else:
            cells[0].font = Font(bold=True, size=10, color=NAVY)
            cells[0].alignment = Alignment(wrap_text=True, vertical="top")
            cells[0].border = BORDER
            cells[1].font = Font(size=10, color=GREY)
            cells[1].alignment = Alignment(wrap_text=True, vertical="top")
            cells[1].border = BORDER

    section("DEFINITIONS")
    kv("Golden hours", "Best call pickup times: 12:00-14:00 and 18:00-20:00.")
    kv("Book on their behalf", "With verbal consent, you open Calendly and pick the slot for the candidate.")
    kv("SLA", "Service Level Agreement = the 24h window from 'Added On' to schedule the interview.")
    return ws


# --------------------------------------------------------------------------- #
# Sheet 4: Ram - Digital ToDo
# --------------------------------------------------------------------------- #
def build_ram(wb, requests):
    ws = wb.create_sheet(RAM_SHEET)
    headers = [
        "Req ID", "Added On", "SLA Deadline", "Hours Left", "NEXT STEP", "Do By",
        "Candidate", "Company", "Round", "Phone", "Email", "Calendly Link (this round)",
        "Dup?", "S1 Blast", "S2 Calendly Check", "S3 Nudge", "S4 Final Check",
        "Ram Status", "Notes",
    ]
    ws.append(headers)
    for c in ws[1]:
        style_header(c, fill=GREEN)

    for i, req in enumerate(requests):
        r = i + 2
        ws.append([
            f"REQ-{i+1:04d}",
            req["added_on"],
            f"=B{r}+1",
            f"=(C{r}-{CLOCK_CELL})*24",
            # NEXT STEP
            (f'=IF(N{r}="","STEP 1: Send WhatsApp + Email blast now (Playbook M1)",'
             f'IF(N{r}="Bad contact","STOP: report bad contact to lead (Playbook)",'
             f'IF(O{r}="Scheduled","DONE: scheduled via Calendly",'
             f'IF(O{r}="","STEP 2: Check Calendly now",'
             f'IF(P{r}="","STEP 3: Send WhatsApp nudge (Playbook M2)",'
             f'IF(Q{r}="Scheduled","DONE: scheduled via Calendly",'
             f'IF(Q{r}="Escalate to call","HANDED to Shyam (call queue)",'
             f'"STEP 4: Final Calendly check")))))))'),
            # Do By
            (f'=IF(N{r}="",B{r},'
             f'IF(O{r}="",B{r}+2/24,'
             f'IF(AND(O{r}="Not yet",P{r}=""),B{r}+2/24,'
             f'IF(Q{r}="",B{r}+4/24,""))))'),
            req["candidate"],
            req["company"],
            req["round"],
            req["phone"],
            req["email"],
            req["calendly"],
            req["dup"],
            None, None, None, None,
            # Ram Status
            (f'=IF(OR(O{r}="Scheduled",Q{r}="Scheduled"),"SCHEDULED",'
             f'IF(Q{r}="Escalate to call","ESCALATED TO CALL",'
             f'IF(N{r}="Bad contact","BAD CONTACT",'
             f'IF(N{r}="","NOT STARTED","IN PROGRESS"))))'),
            None,
        ])

    _finalise_queue(ws, requests, phone_col="J", datetime_cols=("B", "C", "F"),
                    hours_col="D", dup_col="M")

    # Data validations
    dv_s1 = DataValidation(type="list", formula1='"Sent,Bad contact"', allow_blank=True)
    dv_s2 = DataValidation(type="list", formula1='"Scheduled,Not yet"', allow_blank=True)
    dv_s3 = DataValidation(type="list", formula1='"Nudged,N/A"', allow_blank=True)
    dv_s4 = DataValidation(type="list", formula1='"Scheduled,Escalate to call"', allow_blank=True)
    last = len(requests) + 1
    for dv, col in ((dv_s1, "N"), (dv_s2, "O"), (dv_s3, "P"), (dv_s4, "Q")):
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{last}")

    set_widths(ws, {
        "A": 9, "B": 16, "C": 16, "D": 9, "E": 42, "F": 16, "G": 12, "H": 11, "I": 9,
        "J": 15, "K": 22, "L": 26, "M": 6, "N": 12, "O": 14, "P": 10, "Q": 14, "R": 16, "S": 22,
    })
    return ws


# --------------------------------------------------------------------------- #
# Sheet 5: Shyam - Call ToDo
# --------------------------------------------------------------------------- #
def build_shyam(wb, requests):
    ws = wb.create_sheet(SHYAM_SHEET)
    headers = [
        "Req ID", "Added On", "SLA Deadline", "Hours Left", "NEEDS CALL?", "NEXT STEP", "Call By",
        "Candidate", "Company", "Round", "Phone", "Calendly Link (this round)",
        "Call 1", "Call 2", "Shyam Status", "FINAL OUTCOME", "Notes",
    ]
    ws.append(headers)
    for c in ws[1]:
        style_header(c, fill=RED)

    ram = f"'{RAM_SHEET}'"
    for i, req in enumerate(requests):
        r = i + 2
        ws.append([
            f"REQ-{i+1:04d}",
            req["added_on"],
            f"=B{r}+1",
            f"=(C{r}-{CLOCK_CELL})*24",
            # NEEDS CALL?  (reads Ram's escalation for the same request)
            (f'=IF({ram}!Q{r}="Escalate to call","CALL NOW",'
             f'IF({ram}!N{r}="Bad contact","CALL NOW (digital failed)",'
             f'"Ram is handling - no action"))'),
            # NEXT STEP
            (f'=IF(E{r}="Ram is handling - no action","No action needed",'
             f'IF(M{r}="","CALL 1: ring candidate, offer to book on call (Playbook C1)",'
             f'IF(OR(M{r}="Scheduled on call",M{r}="Booked on their behalf"),"DONE: scheduled on call",'
             f'IF(M{r}="Not interested/Declined","CLOSED: candidate declined (log + tell lead)",'
             f'IF(N{r}="","CALL 2: final attempt (Playbook C2)",'
             f'IF(OR(N{r}="Scheduled on call",N{r}="Booked on their behalf"),"DONE: scheduled on call",'
             f'IF(N{r}="Escalate to AM","ESCALATED TO LEAD","Review")))))))'),
            # Call By
            (f'=IF(E{r}="Ram is handling - no action","",'
             f'IF(M{r}="",B{r}+6/24,'
             f'IF(AND(M{r}<>"Scheduled on call",M{r}<>"Booked on their behalf",'
             f'M{r}<>"Not interested/Declined",N{r}=""),B{r}+20/24,"")))'),
            req["candidate"],
            req["company"],
            req["round"],
            req["phone"],
            req["calendly"],
            None, None,
            # Shyam Status
            (f'=IF(OR(M{r}="Scheduled on call",M{r}="Booked on their behalf",'
             f'N{r}="Scheduled on call",N{r}="Booked on their behalf"),"SCHEDULED (call)",'
             f'IF(OR(M{r}="Not interested/Declined",N{r}="Not interested/Declined"),"DECLINED",'
             f'IF(N{r}="Escalate to AM","ESCALATED TO LEAD",'
             f'IF(E{r}="Ram is handling - no action","N/A",'
             f'IF(M{r}="","PENDING CALL 1","IN PROGRESS")))))'),
            # FINAL OUTCOME (single source of truth for KPI; no < or > chars for COUNTIF safety)
            (f'=IF(OR({ram}!O{r}="Scheduled",{ram}!Q{r}="Scheduled"),"Scheduled within SLA (self-serve)",'
             f'IF(OR(M{r}="Scheduled on call",M{r}="Booked on their behalf",'
             f'N{r}="Scheduled on call",N{r}="Booked on their behalf"),"Scheduled within SLA (call)",'
             f'IF(OR(M{r}="Not interested/Declined",N{r}="Not interested/Declined"),"Declined",'
             f'IF(N{r}="Escalate to AM","SLA breach - unresolved","In progress"))))'),
            None,
        ])

    _finalise_queue(ws, requests, phone_col="K", datetime_cols=("B", "C", "G"),
                    hours_col="D", dup_col=None)

    dv_c1 = DataValidation(
        type="list",
        formula1='"Scheduled on call,Booked on their behalf,No answer,Invalid number,Call later,Not interested/Declined"',
        allow_blank=True)
    dv_c2 = DataValidation(
        type="list",
        formula1='"Scheduled on call,Booked on their behalf,Still no answer,Not interested/Declined,Escalate to AM"',
        allow_blank=True)
    last = len(requests) + 1
    for dv, col in ((dv_c1, "M"), (dv_c2, "N")):
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{last}")

    set_widths(ws, {
        "A": 9, "B": 16, "C": 16, "D": 9, "E": 22, "F": 44, "G": 16, "H": 12, "I": 11,
        "J": 9, "K": 15, "L": 22, "M": 22, "N": 20, "O": 16, "P": 28, "Q": 22,
    })
    return ws


def _finalise_queue(ws, requests, phone_col, datetime_cols, hours_col, dup_col):
    """Common formatting: number formats, borders, freeze panes, autofilter, heat-map."""
    last = len(requests) + 1
    # phone as text
    for r in range(2, last + 1):
        ws[f"{phone_col}{r}"].number_format = "@"
    # datetime cols
    for col in datetime_cols:
        for r in range(2, last + 1):
            ws[f"{col}{r}"].number_format = "yyyy-mm-dd hh:mm"
    # hours left number format
    for r in range(2, last + 1):
        ws[f"{hours_col}{r}"].number_format = "0.0"
    # light borders + top alignment on all data cells
    for row in ws.iter_rows(min_row=2, max_row=last, max_col=ws.max_column):
        for c in row:
            c.border = BORDER
            if c.alignment is None or not c.alignment.wrap_text:
                c.alignment = Alignment(vertical="center", wrap_text=True)
    # duplicate rows shaded
    if dup_col:
        for r in range(2, last + 1):
            if ws[f"{dup_col}{r}"].value == "Yes":
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=col_idx).fill = PatternFill("solid", fgColor="F2F2F2")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{last}"
    # Hours Left heat-map: red (urgent) -> green (relaxed)
    ws.conditional_formatting.add(
        f"{hours_col}2:{hours_col}{last}",
        ColorScaleRule(start_type="num", start_value=0, start_color=RED,
                       mid_type="num", mid_value=12, mid_color="FFD966",
                       end_type="num", end_value=24, end_color=GREEN),
    )


# --------------------------------------------------------------------------- #
# Sheet 6: KPI Dashboard
# --------------------------------------------------------------------------- #
def build_dashboard(wb, requests):
    ws = wb.create_sheet("KPI Dashboard")
    ws.sheet_view.showGridLines = False
    set_widths(ws, {"A": 3, "B": 44, "C": 18, "D": 60})
    last = len(requests) + 1
    sh = f"'{SHYAM_SHEET}'"
    rm = f"'{RAM_SHEET}'"
    fo = f"{sh}!$P$2:$P${last}"      # FINAL OUTCOME range
    ram_s1 = f"{rm}!$N$2:$N${last}"  # Ram blast result
    ram_s4 = f"{rm}!$Q$2:$Q${last}"  # Ram final check
    shy_c1 = f"{sh}!$M$2:$M${last}"  # Shyam call 1
    dup_rng = f"{rm}!$M$2:$M${last}"

    t = ws.cell(row=1, column=2, value="KPI Dashboard")
    t.font = Font(bold=True, size=18, color=NAVY)
    ws.cell(row=2, column=2, value="Target: 80% of eligible requests scheduled within 24 hours.").font = Font(
        italic=True, size=11, color=GREEN)

    # Editable "system clock". Hours Left in both queues counts down from this
    # cell, so the demo is sensible even though the sample data is from 2022.
    set_widths(ws, {"F": 30, "G": 20})
    lbl = ws.cell(row=2, column=6, value="TREAT 'NOW' AS (edit to advance time):")
    lbl.font = Font(bold=True, size=11, color=NAVY)
    lbl.alignment = Alignment(horizontal="right", vertical="center")
    clk = ws.cell(row=2, column=7, value=CLOCK_DEFAULT)
    clk.number_format = "yyyy-mm-dd hh:mm"
    clk.font = Font(bold=True, size=12, color=WHITE)
    clk.fill = PatternFill("solid", fgColor=AMBER)
    clk.alignment = Alignment(horizontal="center", vertical="center")
    clk.border = BORDER
    ws.cell(row=3, column=6, value="(All 'Hours Left' countdowns are measured from this moment.)").font = Font(
        italic=True, size=9, color=GREY)

    rows = [
        ("Metric", "Value", "Notes", "header"),                                    # r4
        ("Total requests", f"=COUNTA({sh}!$A$2:$A${last})",
         "All rows in the queue.", ""),                                            # r5
        ("Possible duplicates flagged", f'=COUNTIF({dup_rng},"Yes")',
         "Process the first only.", ""),                                           # r6
        ("Scheduled - self-serve (Ram)", f'=COUNTIF({fo},"Scheduled within SLA (self-serve)")',
         "Booked via Calendly after digital outreach.", ""),                       # r7
        ("Scheduled - on call (Shyam)", f'=COUNTIF({fo},"Scheduled within SLA (call)")',
         "Booked during/after a phone call.", ""),                                 # r8
        ("TOTAL SCHEDULED within 24h", "=C7+C8", "Self-serve + on call.", "good"), # r9
        ("Declined / not interested", f'=COUNTIF({fo},"Declined")',
         "Real drop-offs; excluded from SLA denominator.", ""),                    # r10
        ("Unresolved (SLA breach)", f'=COUNTIF({fo},"SLA breach - unresolved")',
         "Escalated to lead; missed the 24h window.", "bad"),                      # r11
        ("In progress", f'=COUNTIF({fo},"In progress")',
         "Still inside the cadence.", ""),                                         # r12
        ("Eligible for SLA (Total - Declined)", "=C5-C10",
         "Denominator for the % metric.", ""),                                     # r13
        ("% SCHEDULED WITHIN 24h", "=IFERROR(C9/C13,0)",
         "This is the headline number.", "kpi"),                                   # r14
        ('Status vs 80% target', '=IF(C14>=0.8,"ON TRACK (>=80%)","BELOW TARGET - push harder")',
         "Auto-checks the goal.", "status"),                                       # r15
        ("", "", "", "gap"),                                                       # r16
        ("WORKLOAD SPLIT", "", "", "header"),                                      # r17
        ("Ram - blasts sent", f'=COUNTIF({ram_s1},"Sent")',
         "Digital outreach volume.", ""),                                          # r18
        ("Ram - bad contacts", f'=COUNTIF({ram_s1},"Bad contact")',
         "Need correct details from company.", ""),                                # r19
        ("Escalated from Ram to Shyam", f'=COUNTIF({ram_s4},"Escalate to call")',
         "Hand-off volume.", ""),                                                  # r20
        ("Shyam - calls logged (Call 1)", f"=COUNTA({shy_c1})",
         "How many first calls Shyam made.", ""),                                  # r21
    ]

    r = 4
    for label, value, note, kind in rows:
        bcell = ws.cell(row=r, column=2, value=label)
        ccell = ws.cell(row=r, column=3, value=value)
        dcell = ws.cell(row=r, column=4, value=note)
        if kind == "header":
            for c in (bcell, ccell, dcell):
                style_header(c, fill=BLUE)
        elif kind == "gap":
            pass
        else:
            bcell.font = Font(bold=True, size=11, color=GREY)
            bcell.alignment = Alignment(vertical="center", wrap_text=True)
            bcell.border = BORDER
            ccell.alignment = Alignment(horizontal="center", vertical="center")
            ccell.border = BORDER
            dcell.font = Font(size=10, color=GREY)
            dcell.alignment = Alignment(vertical="center", wrap_text=True)
            dcell.border = BORDER
            if kind == "good":
                ccell.font = Font(bold=True, size=12, color=GREEN)
                bcell.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
            elif kind == "bad":
                ccell.font = Font(bold=True, size=12, color=RED)
                bcell.fill = PatternFill("solid", fgColor=LIGHT_RED)
            elif kind == "kpi":
                ccell.number_format = "0.0%"
                ccell.font = Font(bold=True, size=16, color=NAVY)
                bcell.font = Font(bold=True, size=12, color=NAVY)
            elif kind == "status":
                ccell.font = Font(bold=True, size=12, color=NAVY)
        r += 1

    # Conditional format on the KPI % cell (row 14 -> C14)
    ws.conditional_formatting.add("C14", CellIsRule(operator="greaterThanOrEqual",
                                  formula=["0.8"], fill=PatternFill("solid", fgColor=LIGHT_GREEN)))
    ws.conditional_formatting.add("C14", CellIsRule(operator="lessThan",
                                  formula=["0.8"], fill=PatternFill("solid", fgColor=LIGHT_RED)))
    return ws


# --------------------------------------------------------------------------- #
# Sheet 7: INPUT (clean)
# --------------------------------------------------------------------------- #
def build_input_clean(wb, requests):
    ws = wb.create_sheet("INPUT (clean)")
    headers = ["Req ID", "Company", "Interviewer", "Interviewer Email", "Candidate",
               "Candidate Email", "Candidate Phone", "Round", "Calendly (this round)",
               "Added On", "Dup?"]
    ws.append(headers)
    for c in ws[1]:
        style_header(c, fill=NAVY)
    for i, req in enumerate(requests):
        ws.append([
            f"REQ-{i+1:04d}", req["company"], req["interviewer"], req["interviewer_email"],
            req["candidate"], req["email"], req["phone"], req["round"], req["calendly"],
            req["added_on"], req["dup"],
        ])
    last = len(requests) + 1
    for r in range(2, last + 1):
        ws[f"G{r}"].number_format = "@"
        ws[f"J{r}"].number_format = "yyyy-mm-dd hh:mm"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{last}"
    set_widths(ws, {"A": 9, "B": 12, "C": 14, "D": 22, "E": 14, "F": 22, "G": 15,
                    "H": 9, "I": 22, "J": 16, "K": 6})
    return ws


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def create_solution_excel():
    requests = load_requests()
    wb = Workbook()
    build_start_here(wb)
    build_assumptions(wb)
    build_playbook(wb)
    build_ram(wb, requests)
    build_shyam(wb, requests)
    build_dashboard(wb, requests)
    build_input_clean(wb, requests)
    wb.save(OUTPUT_FILE)
    print(f"Created {OUTPUT_FILE} with {len(requests)} requests across {len(wb.sheetnames)} sheets:")
    for s in wb.sheetnames:
        print(f"  - {s}")


if __name__ == "__main__":
    create_solution_excel()
