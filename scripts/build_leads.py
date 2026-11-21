#!/usr/bin/env python3
"""
Build the UK health & wellness lead pack for video editing outreach.

Input : leads/research/batch_*.json  (hand-verified records only)
Output: output/UK_Health_Wellness_Video_Leads.xlsx
        output/leads_master.csv
        output/hot_leads_call_list.csv
        output/whatsapp_outreach.csv

Every record in the input files was found through live research and confirmed
against a real business page or official directory. No records are generated,
pattern-guessed, or padded. Records that could not be corroborated as UK-based
were rejected during research and never entered the ledger.
"""

import csv
import glob
import json
import os
import re
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RESEARCH = os.path.join(ROOT, "leads", "research")
OUT = os.path.join(ROOT, "output")

# ----------------------------------------------------------------------------
# Reference data
# ----------------------------------------------------------------------------

# How much original video a business in this category can realistically use.
VIDEO_NEED = {
    "Aesthetics / Skin Clinic": 5,
    "Aesthetics / Dermatology": 5,
    "Aesthetics / Skin": 5,
    "Yoga Studio": 5,
    "Pilates Studio": 5,
    "Gym / Personal Training": 5,
    "Gym / Fitness": 5,
    "Gym / CrossFit": 5,
    "Gym / Physiotherapy / Spa": 5,
    "Personal Training": 5,
    "Personal Training / Nutrition": 5,
    "Personal Training / Sports Therapy": 5,
    "Dance / Fitness Studio": 5,
    "Boxing / Fitness": 5,
    "Pole / Aerial Fitness Studio": 4,
    "Spa / Massage": 4,
    "Spa / Wellness": 4,
    "Spa / Beauty": 4,
    "Massage Therapy": 4,
    "Sports Massage": 4,
    "Sports Massage / Personal Training": 4,
    "Massage / CBT / Hypnotherapy": 4,
    "Massage / Women's Health": 4,
    "Physiotherapy": 4,
    "Physiotherapy / Chiropractic": 4,
    "Physiotherapy / Osteopathy": 4,
    "Osteopathy": 4,
    "Osteopathy / Pilates": 4,
    "Osteopathy / Sports Massage": 4,
    "Osteopathy / Physiotherapy": 4,
    "Osteopathy / MSK": 4,
    "Chiropractic": 4,
    "Acupuncture": 3,
    "Acupuncture / Chinese Medicine": 3,
    "Nutrition": 4,
    "Nutrition / Dietetics": 4,
    "Nutrition / Wellbeing": 4,
    "Hypnotherapy": 4,
    "Hypnotherapy / CBT": 4,
    "Counselling / Therapy": 3,
    "Meditation / Wellbeing": 3,
    "Mindfulness / Wellbeing": 3,
    "Wellness Centre": 4,
    "Wellness Centre / Nutrition": 4,
    "Wellness / Beauty": 4,
    "Yoga / Movement Studio": 5,
    "Yoga / Wellbeing": 4,
    "Dance / Fitness": 5,
    "Health Club": 3,
}

# What to pitch them. Written per category so the sheet is immediately usable
# as an outreach script rather than just a list of names.
VIDEO_ANGLE = {
    "Aesthetics / Skin Clinic": "Before/after reels, treatment explainers, practitioner trust-building",
    "Aesthetics / Dermatology": "Condition explainers, before/after, consultant credibility clips",
    "Yoga Studio": "Class reels, teacher intros, course/retreat promos",
    "Pilates Studio": "Form-correction reels, member progress, membership promos",
    "Gym / Personal Training": "Member transformations, coach intros, challenge promos",
    "Gym / Fitness": "Member stories, facility tours, class promos",
    "Gym / CrossFit": "Workout clips, competition days, community stories",
    "Personal Training": "Client transformations, technique breakdowns, intro offers",
    "Spa / Massage": "Atmosphere reels, package promos, gift-voucher ads",
    "Spa / Wellness": "Ambience reels, seasonal packages, gift vouchers",
    "Massage Therapy": "Treatment explainers, pain-relief testimonials",
    "Sports Massage": "Recovery explainers, athlete testimonials",
    "Physiotherapy": "Condition explainers, rehab progress, patient testimonials",
    "Osteopathy": "Posture/pain explainers, patient stories",
    "Chiropractic": "Posture explainers, adjustment clips (compliance-aware), testimonials",
    "Acupuncture": "Treatment explainers, myth-busting, patient stories",
    "Nutrition": "Client transformations, myth-busting reels, programme promos",
    "Nutrition / Dietetics": "Evidence-led explainers, programme promos, practitioner credibility",
    "Hypnotherapy": "Anxiety/stress explainers, session walkthroughs, free-consult promos",
    "Hypnotherapy / CBT": "Symptom explainers, free-consult promos",
    "Counselling / Therapy": "Practitioner intros, workplace-wellbeing pitches, destigmatising content",
    "Meditation / Wellbeing": "Guided-practice clips, course promos",
    "Mindfulness / Wellbeing": "Course promos, corporate wellbeing pitches, guided sessions",
    "Wellness Centre": "Facility tours, therapy-menu explainers, package promos",
    "Dance / Fitness Studio": "Class energy reels, showcase edits, beginner-friendly promos",
    "Boxing / Fitness": "Training montages, fight-night edits, beginner-programme promos",
    "Pole / Aerial Fitness Studio": "Performance reels, hen-party/challenge promos",
    "Health Club": "Facility tours, class schedules, membership campaigns",
}

# Records deliberately rejected during research. Kept visible so the client can
# see the vetting was real rather than claimed.
EXCLUDED = [
    ("Hot Yoga Manchester", "Manchester, New Hampshire, USA - not UK"),
    ("Bristol Chiropractic", "Bristol, Rhode Island, USA - not UK"),
    ("Blue Yoga Birmingham", "Birmingham, Michigan, USA - not UK"),
    ("In Bloom Yoga", "Birmingham, Alabama, USA - not UK"),
    ("True40 Fitness Birmingham", "Birmingham, Alabama, USA - not UK"),
    ("Birmingham Ultimate Fitness", "Birmingham, Michigan, USA - not UK"),
    ("Birmingham Clinical Hypnotherapy", "Birmingham, Michigan, USA - not UK"),
    ("Path Foundation Hypnosis Clinic", "Birmingham, Alabama, USA - not UK"),
    ("Yoga London (yogalondon.us)", "Ammon, Idaho, USA - not UK"),
    ("Karen Bishop Nutrition", "Birmingham, Alabama, USA - not UK"),
    ("Sophia Burge Nutrition", "Birmingham, Alabama, USA - not UK"),
    ("Casey Bramble Nutrition", "Birmingham, Alabama, USA - not UK"),
    ("Blue Yoga / In Bloom Yoga listings", "US directory results returned for a UK city query - rejected"),
    ("Manchester hot-yoga listicle studios", "Phone numbers ran in false sequence (834 6284/6287/6289/6291) - generated data, rejected"),
    ("Release Reformer Pilates", "Published phone '0161 171133' is malformed - unverifiable, rejected"),
    ("Box in Balance", "Phone '+44 7700 900123' is the standard placeholder number - rejected"),
]


# ----------------------------------------------------------------------------
# Normalisation helpers
# ----------------------------------------------------------------------------

def digits_only(value: str) -> str:
    return re.sub(r"[^0-9]", "", value or "")


def normalise_phone(raw: str):
    """Return (e164, local, kind, issue)."""
    d = digits_only(raw)
    if not d:
        return "", "", "", ""
    if d.startswith("0044"):
        d = d[4:]
    elif d.startswith("44") and len(d) in (11, 12):  # already intl
        d = d[2:]
    d = d.lstrip("0") if d.startswith("00") else d
    if not d.startswith("0") and len(d) == 10:
        d = "0" + d
    kind, issue = "", ""
    if d.startswith("07") and len(d) == 11:
        kind = "Mobile"
    elif d.startswith(("01", "02", "03")) and len(d) in (10, 11):
        kind = "Landline"
    elif d.startswith("08") and len(d) == 11:
        kind = "Freephone/Non-geographic"
    else:
        kind = "Check format"
        issue = f"Raw number '{raw}' did not normalise cleanly - confirm before dialling"
    e164 = "+44" + d[1:] if d.startswith("0") else "+" + d
    return e164, d, kind, issue


def wa_number(raw: str) -> str:
    """wa.me requires the full international number with no + and no leading 0."""
    d = digits_only(raw)
    if not d:
        return ""
    if d.startswith("0044"):
        d = d[4:]
    elif d.startswith("44"):
        d = d[2:]
    d = d.lstrip("0")
    return "44" + d


def whatsapp_state(rec, phone_kind):
    """WhatsApp is only meaningful on a UK mobile."""
    explicit = (rec.get("whatsapp") or "").strip()
    notes = (rec.get("notes") or "").lower()
    if explicit:
        return "Confirmed on their website", f"https://wa.me/{wa_number(explicit)}"
    if phone_kind == "Mobile" and ("whatsapp" in notes or "text" in notes):
        return "Confirmed in their own wording", f"https://wa.me/{wa_number(rec['phone'])}"
    if phone_kind == "Mobile":
        return "Likely (UK mobile)", f"https://wa.me/{wa_number(rec['phone'])}"
    return "Not applicable", ""


def social_handle(url: str) -> str:
    if not url:
        return ""
    m = re.search(r"(?:instagram\.com|facebook\.com)/([^/?#]+)", url)
    return "@" + m.group(1).lstrip("@") if m else ""


def city_label(rec):
    """Normalise the city field; keep Greater Manchester/Bolton honest."""
    city = rec.get("city", "").strip()
    addr = (rec.get("address") or "").lower()
    if city == "Manchester" and ("bolton" in addr or "bury" in addr):
        return "Greater Manchester"
    return city


def score_lead(rec, e164, email, website, insta, fb, wa_state, video_need, status):
    """Transparent 0-100 score. Documented on the 'How To Use' sheet."""
    s = 0
    if status == "Verified":
        s += 10
    elif status == "Partial":
        s += 4
    if email:
        s += 15
    if e164:
        s += 12
    if wa_state.startswith("Confirmed"):
        s += 10
    elif wa_state.startswith("Likely"):
        s += 5
    if website:
        s += 10
    if insta:
        s += 8
    if fb:
        s += 3
    s += video_need * 4  # up to 20
    return min(s, 100)


def tier(score, has_email, has_phone, wa_state, video_need, status):
    """Hot = reachable by two independent channels AND clear video appetite."""
    channels = sum([bool(has_email), bool(has_phone), wa_state.startswith(("Confirmed", "Likely"))])
    if score >= 70 and channels >= 2 and video_need >= 4 and status == "Verified":
        return "A - Hot"
    if score >= 52 and channels >= 2 and video_need >= 3:
        return "B - Warm"
    return "C - Nurture"


# ----------------------------------------------------------------------------
# Load
# ----------------------------------------------------------------------------

def load_records():
    decoder = json.JSONDecoder()
    records = []
    for path in sorted(glob.glob(os.path.join(RESEARCH, "batch_*.json"))):
        with open(path, encoding="utf-8") as fh:
            raw = fh.read()
        i = 0
        while i < len(raw):
            while i < len(raw) and raw[i] in " \n\r\t":
                i += 1
            if i >= len(raw):
                break
            chunk, i = decoder.raw_decode(raw, i)
            for rec in chunk:
                rec["_source_file"] = os.path.basename(path)
                records.append(rec)
    return records


def build_rows(records):
    rows = []
    for rec in records:
        phone = (rec.get("phone") or "").strip()
        e164, local, kind, issue = normalise_phone(phone)
        email = (rec.get("email") or "").strip()
        website = (rec.get("website") or "").strip()
        insta = (rec.get("instagram") or "").strip()
        fb = (rec.get("facebook") or "").strip()
        wa_state, wa_link = whatsapp_state(rec, kind)
        status = rec.get("status", "Partial")
        category = rec.get("category", "").strip()
        video_need = VIDEO_NEED.get(category, 3)
        sc = score_lead(rec, e164, email, website, insta, fb, wa_state, video_need, status)
        tr = tier(sc, email, e164, wa_state, video_need, status)

        flags = []
        if issue:
            flags.append(issue)
        if status == "Partial":
            missing = [n for n, v in (("email", email), ("phone", phone), ("website", website)) if not v]
            if missing:
                flags.append("Missing " + "/".join(missing) + " - capture on first contact")
        if not (insta or fb):
            flags.append("No social profile found - pitch social setup as an add-on")

        rows.append({
            "Priority": tr,
            "Score": sc,
            "Business": rec.get("name", "").strip(),
            "Category": category,
            "City": city_label(rec),
            "Address": rec.get("address", "").strip(),
            "Postcode": rec.get("postcode", "").strip(),
            "Phone": phone,
            "PhoneE164": e164,
            "PhoneType": kind,
            "WhatsApp": wa_link,
            "WhatsAppStatus": wa_state,
            "Email": email,
            "Website": website,
            "Instagram": insta,
            "InstagramHandle": social_handle(insta),
            "Facebook": fb,
            "FacebookPage": social_handle(fb),
            "VideoNeed": video_need,
            "VideoAngle": VIDEO_ANGLE.get(category, "Class promos, testimonials, brand story"),
            "Status": status,
            "SourceURL": rec.get("source", ""),
            "Notes": rec.get("notes", "").strip(),
            "Flags": "; ".join(flags),
            "DateFound": date.today().isoformat(),
        })
    rows.sort(key=lambda r: (-r["Score"], r["City"], r["Business"]))
    return rows


# ----------------------------------------------------------------------------
# CSV writers
# ----------------------------------------------------------------------------

MASTER_COLS = [
    ("Priority", "Priority"), ("Score", "Score"), ("Business", "Business Name"),
    ("Category", "Category"), ("City", "City"), ("Address", "Address"),
    ("Postcode", "Postcode"), ("Phone", "Phone"), ("PhoneE164", "Phone (E.164)"),
    ("PhoneType", "Phone Type"), ("WhatsApp", "WhatsApp Link"),
    ("WhatsAppStatus", "WhatsApp Status"), ("Email", "Email"), ("Website", "Website"),
    ("Instagram", "Instagram"), ("InstagramHandle", "Instagram Handle"),
    ("Facebook", "Facebook"), ("FacebookPage", "Facebook Page"),
    ("VideoNeed", "Video Need (1-5)"), ("VideoAngle", "Suggested Video Angle"),
    ("Status", "Verification"), ("SourceURL", "Source Verified At"),
    ("Notes", "Notes"), ("Flags", "Flags"), ("DateFound", "Date Found"),
]


def write_csv(path, rows, cols):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow([label for _, label in cols])
        for r in rows:
            w.writerow([r.get(k, "") for k, _ in cols])


# ----------------------------------------------------------------------------
# XLSX writer
# ----------------------------------------------------------------------------

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TIER_FILL = {
    "A - Hot": PatternFill("solid", fgColor="C6EFCE"),
    "B - Warm": PatternFill("solid", fgColor="FFEB9C"),
    "C - Nurture": PatternFill("solid", fgColor="F2F2F2"),
}

WIDTHS = {
    "Priority": 11, "Score": 7, "Business Name": 34, "Category": 26, "City": 18,
    "Address": 46, "Postcode": 10, "Phone": 16, "Phone (E.164)": 15,
    "Phone Type": 16, "WhatsApp Link": 30, "WhatsApp Status": 24, "Email": 34,
    "Website": 38, "Instagram": 38, "Instagram Handle": 24, "Facebook": 34,
    "Facebook Page": 26, "Video Need (1-5)": 9, "Suggested Video Angle": 48,
    "Verification": 12, "Source Verified At": 40, "Notes": 60, "Flags": 44,
    "Date Found": 12,
}


def write_sheet(ws, rows, cols):
    headers = [label for _, label in cols]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS.get(headers[c - 1], 18)
    ws.row_dimensions[1].height = 30

    for r in rows:
        ws.append([r.get(k, "") for k, _ in cols])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        prio = row[0].value
        fill = TIER_FILL.get(prio)
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=False)
            if fill:
                cell.fill = fill
        row[0].font = Font(bold=True, size=10)

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def summary(rows, key):
    agg = {}
    for r in rows:
        k = r[key]
        b = agg.setdefault(k, {"total": 0, "hot": 0, "warm": 0, "wa": 0, "email": 0})
        b["total"] += 1
        if r["Priority"] == "A - Hot":
            b["hot"] += 1
        elif r["Priority"] == "B - Warm":
            b["warm"] += 1
        if r["WhatsApp"]:
            b["wa"] += 1
        if r["Email"]:
            b["email"] += 1
    return sorted(agg.items(), key=lambda kv: (-kv[1]["total"], kv[0]))


def write_xlsx(path, rows):
    wb = Workbook()

    ws = wb.active
    ws.title = "All Leads"
    write_sheet(ws, rows, MASTER_COLS)

    hot = [r for r in rows if r["Priority"] == "A - Hot"]
    write_sheet(wb.create_sheet("Hot Leads (A)"), hot, MASTER_COLS)

    call_cols = [
        ("Priority", "Priority"), ("Business", "Business Name"), ("City", "City"),
        ("Category", "Category"), ("Phone", "Phone to Dial"),
        ("PhoneType", "Phone Type"), ("WhatsApp", "WhatsApp Link"),
        ("Email", "Email"), ("VideoAngle", "What To Pitch"),
    ]
    write_sheet(wb.create_sheet("Call List"), hot + [r for r in rows if r["Priority"] == "B - Warm"], call_cols)

    wa = [r for r in rows if r["WhatsApp"]]
    write_sheet(wb.create_sheet("WhatsApp Outreach"), wa, [
        ("Business", "Business Name"), ("City", "City"), ("Category", "Category"),
        ("WhatsApp", "WhatsApp Link"), ("WhatsAppStatus", "How Confirmed"),
        ("Phone", "Number"), ("VideoAngle", "What To Pitch"),
    ])

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "UK Health & Wellness Video Editing Leads"
    ws2["A1"].font = Font(bold=True, size=14, color="1F3864")
    ws2["A2"] = f"Built {date.today().isoformat()} - {len(rows)} verified businesses across 5 UK cities"
    ws2["A2"].font = Font(italic=True, size=10)

    ws2["A4"] = "By city"
    ws2["A4"].font = Font(bold=True, size=11)
    ws2.append([])
    hdr_row = 6
    for i, h in enumerate(["City", "Total", "Hot (A)", "Warm (B)", "WhatsApp", "Email"], start=1):
        c = ws2.cell(row=hdr_row, column=i, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
    r = hdr_row + 1
    for city, b in summary(rows, "City"):
        ws2.cell(row=r, column=1, value=city)
        ws2.cell(row=r, column=2, value=b["total"])
        ws2.cell(row=r, column=3, value=b["hot"])
        ws2.cell(row=r, column=4, value=b["warm"])
        ws2.cell(row=r, column=5, value=b["wa"])
        ws2.cell(row=r, column=6, value=b["email"])
        r += 1
    ws2.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    ws2.cell(row=r, column=2, value=len(rows)).font = Font(bold=True)
    ws2.cell(row=r, column=3, value=len(hot)).font = Font(bold=True)

    r += 3
    ws2.cell(row=r, column=1, value="By category").font = Font(bold=True, size=11)
    r += 2
    cat_hdr = r
    for i, h in enumerate(["Category", "Total", "Hot (A)", "Warm (B)", "WhatsApp", "Email"], start=1):
        c = ws2.cell(row=cat_hdr, column=i, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
    r = cat_hdr + 1
    for cat, b in summary(rows, "Category"):
        ws2.cell(row=r, column=1, value=cat)
        ws2.cell(row=r, column=2, value=b["total"])
        ws2.cell(row=r, column=3, value=b["hot"])
        ws2.cell(row=r, column=4, value=b["warm"])
        ws2.cell(row=r, column=5, value=b["wa"])
        ws2.cell(row=r, column=6, value=b["email"])
        r += 1
    for col, w in zip("ABCDEF", [42, 10, 10, 10, 12, 10]):
        ws2.column_dimensions[col].width = w

    ws3 = wb.create_sheet("How To Use")
    ws3.column_dimensions["A"].width = 34
    ws3.column_dimensions["B"].width = 104
    lines = [
        ("HOW THIS LIST WAS BUILT", ""),
        ("", ""),
        ("Method", "Each business was found through targeted research, then confirmed against its own website or an official directory. "
                   "Real trading names, real published phone numbers and emails. Nothing is generated, pattern-guessed or padded."),
        ("Count", f"{len(rows)} businesses. That is the number of UK health & wellness businesses that could actually be verified one-by-one in this "
                  "session - not 3,000. Padding to 3,000 would mean inventing numbers, which is exactly what you asked me not to do."),
        ("Volume reality", "A verified UK list of this kind takes roughly 1-3 minutes of human research per business. 3,000 verified records is "
                           "about 75-150 hours of research, or a paid data source such as Google Places API / Companies House / a licensed "
                           "outbound database. This pack is the verified seed: use it to land the first clients, then scale."),
        ("Fake leads removed", "See the 'Rejected' sheet. Several US businesses (Birmingham Alabama/Michigan, Bristol Rhode Island, Manchester "
                               "New Hampshire) rank for these city searches. They were all rejected."),
        ("", ""),
        ("COLUMN DEFINITIONS", ""),
        ("Priority", "A - Hot: two or more working contact channels, verification passed, and a clear ongoing video need. Call these first."),
        ("", "B - Warm: reachable but one channel missing or video need less obvious."),
        ("", "C - Nurture: thin contact data. Keep for email nurture only."),
        ("Score", "0-100. Verified status +10, email +15, phone +12, confirmed WhatsApp +10 (likely +5), website +10, "
                  "Instagram +8, Facebook +3, video need x4 (max 20)."),
        ("Phone (E.164)", "Normalised international format for diallers and WhatsApp. Original formatting is kept in 'Phone'."),
        ("WhatsApp Link", "wa.me click-to-chat. 'Confirmed on their website' means the business publishes a WhatsApp number itself. "
                          "'Likely' means it is a UK mobile, which usually has WhatsApp but is not guaranteed."),
        ("Verification", "Verified = phone or email confirmed on the business's own page. Partial = found via a directory or one field "
                         "is hidden/obfuscated - capture it on first contact."),
        ("Source Verified At", "The exact page the details came from. Keep it: it lets you re-check a lead before you quote."),
        ("", ""),
        ("HOW TO RUN THIS", ""),
        ("1", "Work the Call List sheet top to bottom. A leads first, and within A, businesses with a confirmed WhatsApp before landlines."),
        ("2", "Open their website and watch 60 seconds of what they already post. The 'What To Pitch' column tells you the angle."),
        ("3", "Lead with a specific observation, not a generic pitch. Example: 'Saw you post class clips filmed on a phone - I cut a 3-reel "
              "test from your last month of footage, want it free?'"),
        ("4", "Ask which number is best for WhatsApp, and confirm the email on the call. Fix any row flagged 'Check format' or 'Missing'."),
        ("5", "Log every outcome in the Status column so this sheet becomes your CRM, not just a one-off download."),
        ("", ""),
        ("COMPLIANCE / GOOD PRACTICE", ""),
        ("UK GDPR & PECR", "These are business contact details, which is lawful for B2B outreach - but you must identify yourself, say where "
                           "you got the data, and honour opt-outs immediately. Do not bulk-message consumer numbers via WhatsApp; cold "
                           "WhatsApp marketing needs prior consent under PECR. A personal, specific message to a business owner is lower risk "
                           "but still identify yourself and keep it short."),
        ("Health claims", "Health and aesthetics clients are regulated. Do not promise results in your edits, and be careful with "
                          "before/after content - it needs the client's consent and, for medical claims, their compliance sign-off."),
        ("Re-verify", f"Details go stale. Anything with 'Date Found' more than 90 days old should be re-checked before you pitch."),
    ]
    for i, (a, b) in enumerate(lines, start=1):
        ws3.cell(row=i, column=1, value=a)
        ws3.cell(row=i, column=2, value=b)
        ws3.cell(row=i, column=1).font = Font(bold=a.isupper() or a in {"Method", "Count", "Volume reality", "Fake leads removed"})
        ws3.cell(row=i, column=2).alignment = Alignment(wrap_text=True, vertical="top")
        ws3.cell(row=i, column=1).alignment = Alignment(vertical="top")
        if a.isupper() and a:
            ws3.cell(row=i, column=1).font = Font(bold=True, size=11, color="1F3864")

    ws4 = wb.create_sheet("Rejected")
    ws4.append(["Business", "Why It Was Rejected"])
    for c in ("A1", "B1"):
        ws4[c].fill = HDR_FILL
        ws4[c].font = HDR_FONT
    for name, reason in EXCLUDED:
        ws4.append([name, reason])
    ws4.column_dimensions["A"].width = 40
    ws4.column_dimensions["B"].width = 80
    ws4.freeze_panes = "A2"
    ws4.auto_filter.ref = f"A1:B{ws4.max_row}"

    wb.save(path)


def main():
    os.makedirs(OUT, exist_ok=True)
    records = load_records()
    rows = build_rows(records)

    write_csv(os.path.join(OUT, "leads_master.csv"), rows, MASTER_COLS)
    write_csv(os.path.join(OUT, "hot_leads_call_list.csv"),
              [r for r in rows if r["Priority"] == "A - Hot"], MASTER_COLS)
    write_csv(os.path.join(OUT, "whatsapp_outreach.csv"),
              [r for r in rows if r["WhatsApp"]], [
                  ("Business", "Business Name"), ("City", "City"), ("Category", "Category"),
                  ("WhatsApp", "WhatsApp Link"), ("WhatsAppStatus", "How Confirmed"),
                  ("Phone", "Number"), ("Email", "Email"), ("VideoAngle", "What To Pitch"),
              ])
    write_xlsx(os.path.join(OUT, "UK_Health_Wellness_Video_Leads.xlsx"), rows)

    hot = [r for r in rows if r["Priority"] == "A - Hot"]
    warm = [r for r in rows if r["Priority"] == "B - Warm"]
    print(f"total        : {len(rows)}")
    print(f"hot (A)      : {len(hot)}")
    print(f"warm (B)     : {len(warm)}")
    print(f"whatsapp     : {sum(1 for r in rows if r['WhatsApp'])}")
    print(f"confirmed wa : {sum(1 for r in rows if r['WhatsAppStatus'].startswith('Confirmed'))}")
    print(f"with email   : {sum(1 for r in rows if r['Email'])}")
    print(f"with phone   : {sum(1 for r in rows if r['Phone'])}")
    print(f"with website : {sum(1 for r in rows if r['Website'])}")
    print(f"with social  : {sum(1 for r in rows if r['Instagram'] or r['Facebook'])}")
    print("cities       :", {c: b['total'] for c, b in summary(rows, 'City')})


if __name__ == "__main__":
    main()
