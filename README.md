# UK Health & Wellness — Video Editing Client Leads

Verified outreach list for selling **video editing services** to health & wellness
businesses in **Bristol, Manchester, Leeds, Birmingham and London**.

**159 real businesses. 48 hot leads. 11 confirmed WhatsApp numbers. Zero invented data.**

---

## What's in `output/`

| File | What it is |
|---|---|
| `UK_Health_Wellness_Video_Leads.xlsx` | The main deliverable. 7 sheets, colour-coded by priority. |
| `leads_master.csv` | Same data, flat CSV for any CRM. |
| `hot_leads_call_list.csv` | Just the 48 A-tier leads, phone-first. |
| `whatsapp_outreach.csv` | The 84 leads with a usable `wa.me` click-to-chat link. |

### Workbook sheets

- **All Leads** — full 159 with 25 columns, filterable, frozen panes.
- **Hot Leads (A)** — green. Two or more working channels *and* a clear video need.
- **Call List** — priority-ordered, with a "What To Pitch" column per category.
- **WhatsApp Outreach** — every lead with a click-to-chat link plus how it was confirmed.
- **Summary** — counts by city and by category.
- **How To Use** — column definitions, scoring, the outreach method, GDPR notes.
- **Rejected** — businesses found and thrown out, with reasons.

---

## The straight answer on "3,000 leads"

You asked for 3,000. This pack has 159, and that is deliberate.

Every row here was found through live research and confirmed against the business's
own website or an official directory: real trading names, real published phone
numbers, real emails. Doing that takes roughly 1–3 minutes per business.

3,000 records at that standard is **75–150 hours of human research**. Anything that
claims to hand you 3,000 "verified" UK wellness leads in one pass is either scraped
and unverified, or generated. Generated is the dangerous case — those are the lists
full of dead numbers and businesses that closed years ago, and they burn your
sender reputation on the first blast.

So this is the verified seed instead. **48 hot leads is roughly 4–6 weeks of real
outreach** for a solo video editor, and at a normal 5–10% close rate on warm B2B
outreach that is your first handful of retainers.

### To genuinely reach 3,000

| Route | Realistic output | Notes |
|---|---|---|
| Scale this process | 40–80 verified/day per researcher | Keep the `leads/research/*.json` format, re-run the script |
| Google Places API | Thousands, ~£0.02–0.03 per place | Gives name, address, phone, website. Does **not** give email, Instagram or WhatsApp |
| Companies House API | Full UK company registry | Free. Good for confirming a business is real and active |
| Licensed B2B data (Apollo, Lusha, Cognism) | 1k–50k | Fastest path, but verify a 50-row sample before paying |

Merge those into `leads/research/batch_*.json` using the existing schema and run
`python scripts/build_leads.py` — the scoring, WhatsApp links, dedupe and workbook
rebuild automatically.

---

## Why this list is trustworthy

The city searches for these five UK cities are polluted with **US businesses of the
same name**. All of these were found and rejected (see the `Rejected` sheet):

- "Bristol Chiropractic" — Bristol, **Rhode Island, USA**
- "Blue Yoga Birmingham", "Birmingham Ultimate Fitness" — Birmingham, **Michigan**
- "In Bloom Yoga", "True40 Birmingham", "Karen Bishop Nutrition" — Birmingham, **Alabama**
- "Hot Yoga Manchester" — Manchester, **New Hampshire**
- "Yoga London" — Ammon, **Idaho**

Also rejected: a Manchester studio listicle whose phone numbers ran
`834 6284 / 6287 / 6289 / 6291` — sequential numbers mean generated data, not real
listings. And a gym whose number was `+44 7700 900123`, the standard placeholder.

### Honesty flags kept in the data

Rows that need care are marked rather than quietly cleaned up:

- **Nutrishan** (Leeds) — their own page shows two different mobiles. Confirm on first contact.
- **Bristol Mindfulness** — number read from Facebook and shared with another local listing.
- **Birdwell Osteopathy** (Bristol) — published number is also shown as a fax elsewhere.
- **Mindful Lives** (Bristol) — mobile sits next to a French number; email is the safer route.
- **Skin Clinix** (Manchester) — their site misspells their own email on one page.
- **24 C-tier rows** — missing a website, social or email. Contact data is real, marketing presence is thin.

`Verification = Partial` means found via a directory or one field is obfuscated.
`Verified` means confirmed on the business's own page.

---

## Using it

1. Open **Call List**. Work A-tier first, and within A, WhatsApp before landlines.
2. Check what they already post before you call — the `What To Pitch` column gives the angle.
3. Lead with something specific, not a generic pitch. Free sample edit beats a rate card.
4. Fix any row flagged in `Flags` as you go, and log the outcome so the sheet becomes your CRM.

**Compliance:** these are business contact details, lawful for B2B outreach under UK
GDPR, but you must identify yourself and honour opt-outs. Cold WhatsApp marketing to
consumers needs prior consent under PECR — keep messages personal, short and
identifiable. Health and aesthetics are regulated sectors: don't promise outcomes in
your edits, and get client sign-off on before/after content.

---

## Rebuilding

```bash
python3 -m venv .venv && ./.venv/bin/pip install openpyxl
./.venv/bin/python scripts/build_leads.py
```

Source records live in `leads/research/batch_*.json`. Add a file there and re-run to
extend the list. The script derives phone types, E.164 numbers, `wa.me` links,
scores, tiers and flags — and fails loudly on duplicates or malformed numbers.
