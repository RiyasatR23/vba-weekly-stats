"""Generates the sample workbook data/Echo_Weekly_Report.xlsx.

All names, emails and figures are made up. The raw sheets are deliberately messy
(mixed date formats, "R 4 500,00" text amounts, inconsistent casing, duplicates,
blank rows) because that is what real exports look like and what the macro cleans.

Usage: python tools/make_sample_data.py
"""
import datetime as dt
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

random.seed(23)
START = dt.date(2026, 9, 14)          # Monday
WEEKS = 8                             # ends Sunday 2026-11-08
OUT = Path(__file__).resolve().parent.parent / "data" / "Echo_Weekly_Report.xlsx"

FIRST = ["Thabo", "Aisha", "Ruan", "Naledi", "Yusuf", "Lerato", "Pieter", "Fatima", "Sipho",
         "Zanele", "Imraan", "Karabo", "Megan", "Tshepo", "Nadia", "Kagiso", "Ayesha", "Bongani"]
LAST = ["Mokoena", "Patel", "van Wyk", "Dlamini", "Khan", "Nkosi", "Botha", "Ismail", "Ndlovu",
        "Mahlangu", "Moosa", "Molefe", "Smith", "Sithole", "Essop", "Mabaso", "Naidoo", "Zulu"]
BUSINESS = ["Studio", "Auto", "Dental", "Bakery", "Fitness", "Logistics", "Salon", "Plumbing",
            "Interiors", "Coffee", "Legal", "Academy", "Motors", "Events"]
SERVICES = {  # service: (lead value range, invoice amounts)
    "Website": ((4500, 15500), [4500, 8500, 13500, 15500]),
    "AI assistant": ((3500, 9000), [3500, 6000]),
    "Automation": ((5000, 12000), [5500, 7500, 9500]),
    "SEO": ((2500, 4500), [3500]),
    "Hosting & care": ((450, 900), [450, 650]),
}
SOURCES = ["Website", "WhatsApp", "Referral", "Instagram", "Google Business"]
SOURCE_W = [40, 22, 18, 12, 8]
SERVICE_W = [45, 15, 15, 12, 13]

# messy spellings the macro has to normalise
SOURCE_MESS = {"Website": ["Website", "website", "Website ", "WEBSITE"],
               "WhatsApp": ["WhatsApp", "whatsapp", "Whatsapp "],
               "Referral": ["Referral", "referral"],
               "Instagram": ["Instagram", "instagram", "IG"],
               "Google Business": ["Google Business", "google business", "GBP"]}
SCORE_MESS = {"hot": ["hot", "Hot", "HOT ", " hot"], "warm": ["warm", "Warm", "WARM"],
              "cold": ["cold", "Cold", "cold "]}


def person():
    f, l = random.choice(FIRST), random.choice(LAST)
    biz = f"{l.split()[-1]} {random.choice(BUSINESS)}"
    email = f"{f.lower()}@{l.split()[-1].lower()}{random.choice(BUSINESS).lower()}.example"
    return f"{f} {l}", email, biz


def rand_money_text(amount):
    """Same amount, written the different ways people and exports write it."""
    whole = f"{amount:,}".replace(",", " ")
    return random.choice([amount, amount, f"R{whole}", f"R {whole},00", f"R{amount}.00", f"{amount}"])


leads, invoices, web = [], [], []
inv_no = 1041
for w in range(WEEKS):
    week_leads = 9 + w * 2 + random.randint(-2, 3)            # growing pipeline
    for _ in range(week_leads):
        day = START + dt.timedelta(days=7 * w + random.randint(0, 6))
        ts = dt.datetime.combine(day, dt.time(random.randint(7, 21), random.randint(0, 59), random.randint(0, 59)))
        name, email, biz = person()
        src = random.choices(SOURCES, SOURCE_W)[0]
        svc = random.choices(list(SERVICES), SERVICE_W)[0]
        score = random.choices(["hot", "warm", "cold"], [25 + w * 2, 45, 30 - w * 2])[0]
        lo, hi = SERVICES[svc][0]
        value = round(random.randint(lo, hi), -2)
        status = random.choices(["new", "contacted", "won", "lost"],
                                {"hot": [15, 30, 40, 15], "warm": [30, 45, 12, 13], "cold": [45, 30, 3, 22]}[score])[0]
        score_txt = random.choice(SCORE_MESS[score]) if random.random() > 0.05 else ""   # some unscored
        # created_at exported as ISO text, like a Supabase CSV
        row = [ts.strftime("%Y-%m-%dT%H:%M:%SZ"), name, email, random.choice(SOURCE_MESS[src]),
               svc, score_txt, status, rand_money_text(value)]
        leads.append(row)
        if random.random() < 0.08:                       # double-submitted form
            leads.append(list(row))
        if status == "won":                             # a won lead becomes an invoice 0-6 days later
            inv_day = day + dt.timedelta(days=random.randint(0, 6))
            if inv_day <= START + dt.timedelta(days=7 * WEEKS - 1):
                amt = random.choice(SERVICES[svc][1])
                paid = random.random() < (0.85 if inv_day < START + dt.timedelta(days=7 * (WEEKS - 1)) else 0.45)
                invoices.append([f"INV-{inv_no}", inv_day, biz, svc, amt, "Paid" if paid else "Unpaid"])
                inv_no += 1
    # retainers (hosting/SEO) billed every Monday
    for _ in range(3 + w // 2):
        svc = random.choice(["Hosting & care", "SEO"])
        amt = random.choice(SERVICES[svc][1])
        invoices.append([f"INV-{inv_no}", START + dt.timedelta(days=7 * w), person()[2], svc, amt,
                         "Paid" if random.random() < 0.9 else "Unpaid"])
        inv_no += 1

invoices.sort(key=lambda r: (r[1], r[0]))
# mess up the invoice export: mixed date formats, money as text, inconsistent status casing
messy_inv = []
for r in invoices:
    d = r[1]
    date_val = random.choice([d.strftime("%d/%m/%Y"), d.strftime("%d/%m/%Y"), d.strftime("%Y-%m-%d"), d])
    status = random.choice({"Paid": ["Paid", "paid", "PAID ", "Paid"], "Unpaid": ["Unpaid", "unpaid", "Unpaid "]}[r[5]])
    messy_inv.append([r[0], date_val, r[2], r[3], rand_money_text(r[4]), status])
# an invoice that was exported twice
messy_inv.insert(14, list(messy_inv[13]))

for i in range(7 * WEEKS):
    day = START + dt.timedelta(days=i)
    w = i // 7
    base = 70 + w * 9 + (25 if day.weekday() < 5 else -20)
    sessions = max(10, int(base + random.randint(-15, 15)))
    enquiries = max(0, int(sessions * random.uniform(0.012, 0.03)))
    resp = round(random.uniform(1.0, 6.5) - w * 0.35, 1)
    web.append([day, sessions, enquiries, max(0.3, resp)])

wb = Workbook()
ctrl = wb.active
ctrl.title = "Control"
purple = PatternFill("solid", fgColor="4B2BD6")
ctrl["B2"] = "Echo Corp · Weekly operations report"
ctrl["B2"].font = Font(size=18, bold=True, color="4B2BD6")
ctrl["B4"] = "Report week (any date in the week)"
ctrl["B4"].font = Font(bold=True)
ctrl["C4"].number_format = "dd mmm yyyy"
ctrl["D4"] = "Leave blank to use the latest full week in the data"
ctrl["D4"].font = Font(italic=True, color="777777")
ctrl["B6"] = "How to use"
ctrl["B6"].font = Font(bold=True)
steps = ["1. Paste this week's exports into Raw_Leads, Raw_Invoices and Raw_Web (keep the headers).",
         "2. Click Generate weekly report.",
         "3. The Report sheet is rebuilt and a PDF is saved next to this workbook."]
for i, s in enumerate(steps):
    ctrl.cell(row=7 + i, column=2, value=s)
ctrl.column_dimensions["A"].width = 3
ctrl.column_dimensions["B"].width = 36
ctrl.column_dimensions["C"].width = 16

def sheet(name, headers, rows, widths):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = purple
    for r in rows:
        ws.append(r)
    for i, wdt in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = wdt
    ws.freeze_panes = "A2"
    return ws

sheet("Raw_Leads", ["created_at", "name", "email", "source", "service", "score", "status", "est_value"],
      leads, [22, 20, 34, 16, 16, 8, 11, 12])
ws_inv = sheet("Raw_Invoices", ["invoice_no", "invoice_date", "client", "service", "amount", "status"],
               messy_inv, [12, 13, 24, 16, 14, 9])
for row in ws_inv.iter_rows(min_row=2, min_col=2, max_col=2):
    for c in row:
        if isinstance(c.value, dt.date):
            c.number_format = "yyyy/mm/dd"
web_rows = list(web)
web_rows.insert(20, [None, None, None, None])        # blank row in the middle of the export
ws_web = sheet("Raw_Web", ["date", "sessions", "enquiries", "avg_response_hrs"], web_rows, [12, 10, 10, 18])
for row in ws_web.iter_rows(min_row=2, min_col=1, max_col=1):
    for c in row:
        c.number_format = "yyyy-mm-dd"
wb.create_sheet("Report")

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print(f"{OUT}: {len(leads)} lead rows, {len(messy_inv)} invoice rows, {len(web_rows)} web rows")
