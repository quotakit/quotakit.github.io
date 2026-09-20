"""Compose Etsy-ready listing images (2000x2000) from rendered sheet PNGs."""
import os, glob
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageOps

ROOT = os.environ.get("QK_IMAGES", os.path.join(os.path.dirname(__file__), "..", "images"))
REN = os.path.join(ROOT, "render")
OUT = os.path.join(ROOT, "listing")
os.makedirs(OUT, exist_ok=True)
FONT_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
FONT_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
NAVY = (31, 42, 68); TEAL = (14, 159, 142); TEAL_L = (217, 242, 239); OFF = (244, 246, 248); WHITE = (255, 255, 255)

PRODUCTS = {
    "Sales-Commission-Calculator": ("Sales Commission\nCalculator", "Tiered rates · quota bonus · rep statements", ["Dashboard", "Statement", "Settings"]),
    "Incentive-Plan-Simulator": ("Sales Incentive\nPlan Simulator", "Payout curves · 3 scenarios · cost of plan", ["Dashboard", "Payout_Curve", "Plan_Design"]),
    "Sales-KPI-Tracker": ("Sales KPI\nDashboard", "Monthly tracker · funnel · RAG scorecard", ["Dashboard", "KPI_Calc", "Monthly_Input"]),
    "Pipeline-Deal-Tracker": ("Sales Pipeline &\nDeal Tracker", "Weighted forecast · stages · overdue flags", ["Dashboard", "Deals", "Settings"]),
    "Territory-Coverage-Planner": ("Territory & Account\nCoverage Planner", "A/B/C visit plans · coverage % · rep league", ["Dashboard", "Rep_Summary", "Accounts"]),
}


def crop_content(img, pad=30):
    """Trim white margins from a rendered page."""
    g = ImageOps.invert(img.convert("L"))
    bbox = g.point(lambda p: 255 if p > 8 else 0).getbbox()
    if not bbox:
        return img
    l, t, r, b = bbox
    return img.crop((max(0, l - pad), max(0, t - pad), min(img.width, r + pad), min(img.height, b + pad)))


def window(img, width):
    """Put a screenshot inside a mac-style window frame with a soft shadow."""
    ratio = width / img.width
    img = img.resize((width, int(img.height * ratio)), Image.LANCZOS)
    bar = 44
    frame = Image.new("RGB", (img.width + 2, img.height + bar + 2), (225, 228, 233))
    frame.paste((236, 238, 241), (1, 1, img.width + 1, bar))
    d = ImageDraw.Draw(frame)
    for i, col in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
        d.ellipse((18 + i * 26, 14, 34 + i * 26, 30), fill=col)
    frame.paste(img, (1, bar + 1))
    # rounded mask + shadow
    mask = Image.new("L", frame.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, frame.width - 1, frame.height - 1), radius=22, fill=255)
    shadow = Image.new("RGBA", (frame.width + 120, frame.height + 120), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.rounded_rectangle((60, 75, frame.width + 60, frame.height + 75), radius=24, fill=(0, 0, 0, 90))
    shadow = shadow.filter(ImageFilter.GaussianBlur(28))
    out = shadow.copy()
    out.paste(frame, (60, 60), mask)
    return out


def text_block(draw, xy, lines, font, fill, spacing=10):
    x, y = xy
    for ln in lines:
        draw.text((x, y), ln, font=font, fill=fill)
        y += font.getbbox(ln)[3] + spacing
    return y


def badge(draw, xy, text, font, fg, bg):
    x, y = xy
    w = draw.textlength(text, font=font) + 40
    h = font.getbbox(text)[3] + 26
    draw.rounded_rectangle((x, y, x + w, y + h), radius=h // 2, fill=bg)
    draw.text((x + 20, y + 11), text, font=font, fill=fg)
    return x + w + 16


def main_image(prod, title, sub, shot):
    canvas = Image.new("RGB", (2000, 2000), OFF)
    d = ImageDraw.Draw(canvas)
    # header band
    d.rectangle((0, 0, 2000, 620), fill=NAVY)
    d.rectangle((0, 620, 2000, 632), fill=TEAL)
    fT = ImageFont.truetype(FONT_B, 110); fS = ImageFont.truetype(FONT_R, 46); fB = ImageFont.truetype(FONT_B, 34); fK = ImageFont.truetype(FONT_B, 40)
    d.text((100, 70), "QUOTAKIT", font=fK, fill=TEAL)
    y = text_block(d, (100, 140), title.split("\n"), fT, WHITE, 6)
    d.text((100, y + 10), sub, font=fS, fill=(201, 211, 224))
    # screenshot window
    win = window(shot, 1720)
    canvas.paste(win, ((2000 - win.width) // 2, 700), win)
    # badges
    x = 100
    for t in ["Excel + Google Sheets", "Instant download", "No macros"]:
        x = badge(d, (x, 1860), t, fB, NAVY, TEAL_L)
    return canvas


def secondary_image(prod, caption, shot, k):
    canvas = Image.new("RGB", (2000, 2000), OFF)
    d = ImageDraw.Draw(canvas)
    fC = ImageFont.truetype(FONT_B, 64); fK = ImageFont.truetype(FONT_B, 36); fN = ImageFont.truetype(FONT_R, 40)
    d.rectangle((0, 0, 2000, 24), fill=TEAL)
    d.text((100, 90), "QUOTAKIT", font=fK, fill=TEAL)
    d.text((100, 150), caption, font=fC, fill=NAVY)
    win = window(shot, 1760)
    win.thumbnail((1800, 1500))
    canvas.paste(win, ((2000 - win.width) // 2, 300), win)
    d.text((100, 1900), f"Screen {k} of 4  ·  Sample data shown", font=fN, fill=(102, 112, 133))
    return canvas


def feature_image(prod, title, bullets):
    canvas = Image.new("RGB", (2000, 2000), NAVY)
    d = ImageDraw.Draw(canvas)
    fT = ImageFont.truetype(FONT_B, 96); fB = ImageFont.truetype(FONT_R, 52); fK = ImageFont.truetype(FONT_B, 40)
    d.text((120, 120), "QUOTAKIT", font=fK, fill=TEAL)
    d.text((120, 190), "What's inside", font=fT, fill=WHITE)
    y = 380
    for b in bullets:
        d.rounded_rectangle((120, y + 16, 150, y + 46), radius=8, fill=TEAL)
        lines = []
        words = b.split(); cur = ""
        for w in words:
            if d.textlength(cur + " " + w, font=fB) > 1650:
                lines.append(cur); cur = w
            else:
                cur = (cur + " " + w).strip()
        lines.append(cur)
        for ln in lines:
            d.text((190, y), ln, font=fB, fill=WHITE); y += 70
        y += 30
    d.text((120, 1860), "Created with AI assistance · designed and tested by QuotaKit", font=ImageFont.truetype(FONT_R, 36), fill=(201, 211, 224))
    return canvas


FEATURES = {
    "Sales-Commission-Calculator": ["Up to 5 commission tiers with marginal or whole-amount payout, switchable in one cell",
                                    "Rep roster (50 reps) with monthly quotas, quota bonus and an optional monthly cap",
                                    "Sales log for 1,000 transactions; monthly sales and commission matrices build themselves",
                                    "Print-ready statement for any rep and month with the tier-by-tier breakdown",
                                    "Dashboard: YTD sales, commission, bonus, effective rate, monthly chart, top 10 reps"],
    "Incentive-Plan-Simulator": ["Design a payout curve with up to 8 points: threshold, accelerators, cap and kicker",
                                 "Step (slab) or linear interpolation — switch with one dropdown",
                                 "100-rep team, three scenarios (Base, Upside, Downside) with levers to stress-test cost",
                                 "Payout Curve tab with chart and a what-if calculator for any attainment",
                                 "Dashboard: plan cost, cost-to-sales ratio, reps paid, reps at cap, attainment distribution"],
    "Sales-KPI-Tracker": ["Seven inputs per month: leads, meetings, opportunities, won, lost, revenue, new customers",
                          "Automatic conversion rates, win rate, deal size, revenue per lead, MoM growth, run-rate",
                          "Benchmarks you set, with red/amber/green status per month",
                          "Dashboard: YTD tiles, revenue vs target chart, funnel chart with benchmark comparison",
                          "Works for one rep or a whole team"],
    "Pipeline-Deal-Tracker": ["300-deal tracker with owner and stage dropdowns, next step and notes",
                              "Eight stages with editable win probabilities; weighted value per deal",
                              "Age in days, overdue flags and close-month grouping calculated for you",
                              "Six-month weighted forecast from your chosen start month",
                              "Dashboard: open and weighted pipeline, win rate, stage chart, top 10 open deals"],
    "Territory-Coverage-Planner": ["300 accounts with A/B/C tiers and visit frequency targets per tier",
                                   "Log actual visits per month; planned, actual, coverage % and gap calculate per account",
                                   "'Missed' flag for A/B accounts with zero visits in the latest month",
                                   "Rep Summary: accounts by tier, coverage %, missed accounts, potential at risk",
                                   "Dashboard: coverage tiles, planned vs actual chart, tier coverage, rep league table"],
}

for prod, (title, sub, sheets) in PRODUCTS.items():
    shots = []
    for sh in sheets:
        p = os.path.join(REN, f"{prod}__{sh}-1.png")
        shots.append(crop_content(Image.open(p).convert("RGB")))
    main_image(prod, title, sub, shots[0]).save(os.path.join(OUT, f"{prod}_01_main.png"), optimize=True)
    feature_image(prod, title, FEATURES[prod]).save(os.path.join(OUT, f"{prod}_02_features.png"), optimize=True)
    secondary_image(prod, sheets[1].replace("_", " "), shots[1], 3).save(os.path.join(OUT, f"{prod}_03_{sheets[1]}.png"), optimize=True)
    secondary_image(prod, sheets[2].replace("_", " "), shots[2], 4).save(os.path.join(OUT, f"{prod}_04_{sheets[2]}.png"), optimize=True)
    print("done", prod)
