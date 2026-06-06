"""
Advertorial Research Hub — Professional Excel Workbook Builder
"""
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import SheetView
import datetime

OUTPUT_PATH = "/mnt/user-data/outputs/Advertorial_Research_Tracker.xlsx"

# ─── Color Palette ────────────────────────────────────────────────────────────
WHITE       = "FFFFFFFF"
LIGHT_GRAY  = "FFF5F7FA"
DARK_SLATE  = "FF1E293B"
DARK_INK    = "FF0F172A"
MUTED_GRAY  = "FF64748B"
BORDER_CLR  = "FFE4E7EC"

COBALT      = "FF0047AB"
COBALT_LIGHT= "FFEBF0FB"
COBALT_MED  = "FFB8CEFB"

TEAL        = "FF0F766E"
TEAL_LIGHT  = "FFE6FAFA"
TEAL_MED    = "FFB2E0DC"

SLATE       = "FF1E293B"
SLATE_LIGHT = "FFF1F5F9"

# Status pills
S_DONE_TXT   = "FF065F46"; S_DONE_BG   = "FFD1FAE5"
S_PROG_TXT   = "FFD97706"; S_PROG_BG   = "FFFEF3C7"
S_TODO_TXT   = "FF1D62D1"; S_TODO_BG   = "FFEBF0FB"
S_BLOK_TXT   = "FF9F1239"; S_BLOK_BG   = "FFFFE4E6"

# Priority pills
P_HIGH_TXT = "FF9F1239"; P_HIGH_BG = "FFFFE4E6"
P_MED_TXT  = "FFD97706"; P_MED_BG  = "FFFEF3C7"
P_LOW_TXT  = "FF065F46"; P_LOW_BG  = "FFD1FAE5"


# ─── Helpers ──────────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(name="Aptos", size=10, bold=False, italic=False, color="FF000000"):
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def bottom_border(color=BORDER_CLR):
    s = Side(style="thin", color=color)
    return Border(bottom=s)

def left_border(color, thick=False):
    s = Side(style="medium" if thick else "thin", color=color)
    return Border(left=s)

def top_border_medium(color):
    s = Side(style="medium", color=color)
    return Border(top=s)

def no_border():
    return Border()

def set_col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width

def disable_gridlines(ws):
    ws.sheet_view.showGridLines = False

def apply_row_fill(ws, row, col_start, col_end, hex_color):
    f = fill(hex_color)
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = f

def write_cell(ws, row, col, value, fnt=None, aln=None, bdr=None, flll=None):
    cell = ws.cell(row=row, column=col, value=value)
    if fnt:  cell.font      = fnt
    if aln:  cell.alignment = aln
    if bdr:  cell.border    = bdr
    if flll: cell.fill      = flll
    return cell


# ─── Shared Layout Elements ───────────────────────────────────────────────────
def build_title_area(ws, title, subtitle, primary_color, tag="ADVERTORIAL RESEARCH HUB"):
    """Rows 1-7: accent stripe col A, title row 5, subtitle row 6, rule row 4."""
    ws.row_dimensions[1].height = 6
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 3
    ws.row_dimensions[5].height = 34
    ws.row_dimensions[6].height = 18
    ws.row_dimensions[7].height = 10

    # Col A accent stripe (rows 1-40)
    for r in range(1, 60):
        ws.cell(row=r, column=1).fill = fill(primary_color)

    # Top rule row 4
    for c in range(1, 20):
        ws.cell(row=4, column=c).fill = fill(primary_color)

    # Title (col C, row 5)
    cell = ws.cell(row=5, column=3, value=title)
    cell.font      = font("Aptos", 22, bold=True, color=DARK_INK)
    cell.alignment = align("left", "center")

    # Branded tag badge (col M–N, row 5)
    tag_cell = ws.cell(row=5, column=13, value=tag)
    tag_cell.font      = font("Aptos", 8, bold=True, color="FF" + primary_color[2:])
    tag_cell.alignment = align("right", "center")
    tag_cell.fill      = fill(primary_color[0:2] + primary_color[2:].replace(primary_color[2:], _lighten(primary_color)))

    # Subtitle (col C, row 6)
    sub_cell = ws.cell(row=6, column=3, value=subtitle)
    sub_cell.font      = font("Aptos", 10, italic=True, color=MUTED_GRAY)
    sub_cell.alignment = align("left", "center")


def _lighten(hex_color):
    """Return a tinted version of a hex color (very light)."""
    mapping = {
        COBALT: COBALT_LIGHT[2:],
        TEAL:   TEAL_LIGHT[2:],
        SLATE:  SLATE_LIGHT[2:],
    }
    return mapping.get(hex_color, "F1F5F9")


def build_header_row(ws, row, columns, primary_color):
    """columns = list of (col_index, label)"""
    ws.row_dimensions[row].height = 28
    hdr_fill = fill(DARK_SLATE)
    for col_idx, label in columns:
        cell = ws.cell(row=row, column=col_idx, value=label.upper())
        cell.font      = font("Aptos", 8, bold=True, color="FFFFFFFF")
        cell.fill      = hdr_fill
        cell.alignment = align("left", "center")
        cell.border    = Border(
            bottom=Side(style="thin", color=primary_color[2:])
        )


def pill_cell(ws, row, col, label, txt_color, bg_color):
    cell = ws.cell(row=row, column=col, value=label)
    cell.font      = font("Aptos", 9, bold=True, color=txt_color)
    cell.fill      = fill(bg_color)
    cell.alignment = align("center", "center")
    cell.border    = bottom_border()
    return cell


def status_pill(ws, row, col, status):
    mapping = {
        "Done":        ("Done",        S_DONE_TXT, S_DONE_BG),
        "In Progress": ("In Progress", S_PROG_TXT, S_PROG_BG),
        "To Do":       ("To Do",       S_TODO_TXT, S_TODO_BG),
        "Blocked":     ("Blocked",     S_BLOK_TXT, S_BLOK_BG),
    }
    lbl, tc, bc = mapping.get(status, ("To Do", S_TODO_TXT, S_TODO_BG))
    pill_cell(ws, row, col, lbl, tc, bc)


def priority_pill(ws, row, col, priority):
    mapping = {
        "High":   ("▲ HIGH", P_HIGH_TXT, P_HIGH_BG),
        "Medium": ("● MED",  P_MED_TXT,  P_MED_BG),
        "Low":    ("▼ LOW",  P_LOW_TXT,  P_LOW_BG),
    }
    lbl, tc, bc = mapping.get(priority, ("● MED", P_MED_TXT, P_MED_BG))
    pill_cell(ws, row, col, lbl, tc, bc)


def data_cell(ws, row, col, value, row_color=WHITE, wrap=False, bold=False, size=9):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = font("Aptos", size, bold=bold, color=DARK_INK)
    cell.fill      = fill(row_color)
    cell.alignment = align("left", "center", wrap=wrap)
    cell.border    = bottom_border()
    return cell


def footer_bar(ws, row, col_start, col_end, text, primary_color, light_color):
    ws.row_dimensions[row].height = 22
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(light_color)
        cell.border = Border(left=Side(style="thick", color=primary_color[2:]) if c == col_start else Side())
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font      = font("Aptos", 9, italic=True, color=MUTED_GRAY)
    cell.alignment = align("left", "center")


# ─── Sheet 1: Product Research ────────────────────────────────────────────────
PROD_COLS = [
    (2,  "#"),
    (3,  "Product Name"),
    (4,  "Category"),
    (5,  "Target Audience"),
    (6,  "Key Claims & USPs"),
    (7,  "Proof & Sources"),
    (8,  "Angle Ideas"),
    (9,  "Competitor Refs"),
    (10, "Price"),
    (11, "Status"),
    (12, "Priority"),
    (13, "Notes"),
    (14, "Date Added"),
]

PROD_WIDTHS = {
    1: 1.2, 2: 5, 3: 22, 4: 16, 5: 20, 6: 30,
    7: 28, 8: 26, 9: 20, 10: 12, 11: 16, 12: 12, 13: 28, 14: 13
}

PROD_SAMPLES = [
    (1, "OmegaShield Pro", "Supplements", "Men 35–55, health-conscious",
     "Clinically proven to reduce inflammation by 42%; 3rd-party tested; vegan omega-3",
     "JAMA study 2023; NSF Certified; 4,800 verified reviews",
     "Doctor-validated longevity angle; 'What your cardiologist wishes you knew'",
     "Nordic Naturals, Thorne, Viva Naturals", "$49.95/mo",
     "In Progress", "High",
     "Focus on bioavailability vs fish oil; emphasize no fishy aftertaste",
     "2025-06-01"),
    (2, "LumaTrack Band", "Wellness Tech", "Women 28–45, fitness enthusiasts",
     "FDA Class II cleared; 98% sleep-stage accuracy; 14-day battery",
     "Clinical trial NCT04821 (n=340); CE marked; FDA 510(k)",
     "Sleep-debt epidemic angle; 'The $3,000 sleep problem fixed for $149'",
     "Whoop, Oura Ring, Fitbit Sense", "$149 one-time",
     "Done", "High",
     "Emphasize no subscription unlike competitors; dermatologist-tested band",
     "2025-05-28"),
    (3, "ContentFlow AI", "SaaS / MarTech", "Freelance writers, content teams",
     "10× output speed; SEO-optimized drafts; integrates with WordPress & Notion",
     "Beta user case studies; G2 4.8/5 (612 reviews); SOC 2 Type II",
     "Burnout angle for writers; 'Write a month of content in one afternoon'",
     "Jasper, Copy.ai, Writesonic", "$79/mo (Pro)",
     "To Do", "Medium",
     "Highlight human-in-the-loop editing; not a 'replace writers' pitch",
     "2025-06-02"),
    (4, "DermaClear Serum", "Skincare / Beauty", "Women 22–40, acne-prone skin",
     "93% saw clearer skin in 28 days; 2% salicylic acid + niacinamide blend",
     "IRB-approved consumer study (n=180); Allure Best of Beauty 2024",
     "Confidence/transformation angle; 'The serum that changed my morning routine'",
     "Paula's Choice, The Ordinary, CeraVe", "$38 / 30ml",
     "In Progress", "Medium",
     "Before/after UGC available; strong micro-influencer social proof",
     "2025-06-03"),
    (5, "FocusNest App", "Productivity SaaS", "Remote workers, ADHD adults 20–40",
     "62% reduction in task-switching; Pomodoro + AI prioritisation; offline-first",
     "Peer-reviewed productivity study (Univ. of Surrey 2024); 50k MAU",
     "ADHD-friendly design angle; 'Finally, a to-do app built for how you actually think'",
     "Todoist, Motion, Reclaim.ai", "$12/mo",
     "To Do", "Low",
     "Free tier available — good for demo CTA; B2B team plan upsell opportunity",
     "2025-06-04"),
]


def build_sheet1(wb):
    ws = wb.create_sheet("Product Research")
    disable_gridlines(ws)

    for col, w in PROD_WIDTHS.items():
        set_col_width(ws, get_column_letter(col), w)

    build_title_area(ws,
        "Product Research Tracker",
        "Validate claims, proof points, pricing & competitor landscape before writing",
        COBALT)

    HDR_ROW = 8
    ws.row_dimensions[HDR_ROW].height = 28
    build_header_row(ws, HDR_ROW, PROD_COLS, COBALT)

    DATA_START = 9
    today = datetime.date.today()

    for i, sample in enumerate(PROD_SAMPLES):
        r = DATA_START + i
        row_color = WHITE if i % 2 == 0 else LIGHT_GRAY
        ws.row_dimensions[r].height = 42

        num, name, cat, audience, claims, proof, angles, comps, price, status, priority, notes, date = sample

        data_cell(ws, r, 2,  num,      row_color, bold=True)
        data_cell(ws, r, 3,  name,     row_color, bold=True)
        data_cell(ws, r, 4,  cat,      row_color)
        data_cell(ws, r, 5,  audience, row_color)
        data_cell(ws, r, 6,  claims,   row_color, wrap=True)
        data_cell(ws, r, 7,  proof,    row_color, wrap=True)
        data_cell(ws, r, 8,  angles,   row_color, wrap=True)
        data_cell(ws, r, 9,  comps,    row_color, wrap=True)
        data_cell(ws, r, 10, price,    row_color)
        status_pill(ws,   r, 11, status)
        ws.cell(row=r, column=11).fill = fill(
            {"Done": S_DONE_BG, "In Progress": S_PROG_BG,
             "To Do": S_TODO_BG, "Blocked": S_BLOK_BG}.get(status, S_TODO_BG))
        priority_pill(ws, r, 12, priority)
        ws.cell(row=r, column=12).fill = fill(
            {"High": P_HIGH_BG, "Medium": P_MED_BG, "Low": P_LOW_BG}.get(priority, P_MED_BG))
        data_cell(ws, r, 13, notes, row_color, wrap=True)
        data_cell(ws, r, 14, date,  row_color)

        # Left accent stripe
        ws.cell(row=r, column=1).fill = fill(COBALT)

    # Empty rows
    for i in range(5):
        r = DATA_START + 5 + i
        row_color = WHITE if i % 2 == 0 else LIGHT_GRAY
        ws.row_dimensions[r].height = 32
        for c in range(2, 15):
            cell = ws.cell(row=r, column=c)
            cell.fill   = fill(row_color)
            cell.border = bottom_border()
        ws.cell(row=r, column=1).fill = fill(COBALT)

    footer_bar(ws, DATA_START + 11,
               2, 14,
               "  ✦  Tip: Use the Status pills to track research progress. Add competitor URLs in 'Proof & Sources' for quick reference during writing.",
               COBALT, COBALT_LIGHT[2:])
    ws.cell(row=DATA_START + 11, column=1).fill = fill(COBALT)

    return ws


# ─── Sheet 2: Creative Research ───────────────────────────────────────────────
CREA_COLS = [
    (2,  "#"),
    (3,  "Product / Campaign"),
    (4,  "Headline Idea"),
    (5,  "Hook / Opening Line"),
    (6,  "Emotional Angle"),
    (7,  "Format"),
    (8,  "Reference Ad / URL"),
    (9,  "CTA Copy Idea"),
    (10, "Tone"),
    (11, "Status"),
    (12, "Priority"),
    (13, "Creative Notes"),
    (14, "Date"),
]

CREA_WIDTHS = {
    1: 1.2, 2: 5, 3: 22, 4: 26, 5: 30, 6: 22,
    7: 17, 8: 26, 9: 20, 10: 14, 11: 16, 12: 12, 13: 28, 14: 13
}

CREA_SAMPLES = [
    (1, "OmegaShield Pro",
     "The Inflammation Lie Your Doctor Never Told You",
     "Every morning, millions of men over 40 are quietly losing the battle against inflammation — and most don't know it yet.",
     "Fear of silent decline; desire for control over aging",
     "Long-form advertorial",
     "https://ref.example.com/omega-ad-1",
     "Try OmegaShield Free for 30 Days →",
     "Authoritative, empathetic",
     "In Progress", "High",
     "Lead with the 'invisible enemy' concept; pivot to solution at paragraph 3",
     "2025-06-01"),
    (2, "LumaTrack Band",
     "I Was Losing 2 Hours of Deep Sleep Every Night. Here's What Fixed It.",
     "I used to think I was just 'not a morning person.' Turns out, I was barely sleeping at all.",
     "Relatability; relief; transformation",
     "First-person story advertorial",
     "https://ref.example.com/lumatrack-story",
     "Claim Your LumaTrack — Ships in 2 Days →",
     "Conversational, personal",
     "Done", "High",
     "Use authentic voice — avoid over-clinical language. Include screenshot of app data.",
     "2025-05-28"),
    (3, "ContentFlow AI",
     "How a Burned-Out Freelancer Now Delivers 10× More Content (Without the Burnout)",
     "Six months ago, I was ready to quit freelance writing. Now I'm turning down clients.",
     "Burnout → breakthrough; professional pride",
     "Case-study / hybrid advertorial",
     "https://ref.example.com/contentflow-case",
     "Start Your Free 14-Day Trial →",
     "Aspirational, peer-to-peer",
     "To Do", "Medium",
     "Pair with real freelancer testimonial if available; avoid 'AI writes for you' framing",
     "2025-06-02"),
    (4, "DermaClear Serum",
     "28 Days. One Serum. This Is What Happened to My Skin.",
     "I've spent $600 on skincare this year. This $38 serum did more than all of it combined.",
     "Skepticism → conviction; self-worth",
     "Transformation story + product breakdown",
     "https://ref.example.com/dermaclear-ugc",
     "Get Clearer Skin in 28 Days — Guaranteed →",
     "Warm, relatable, empowering",
     "In Progress", "Medium",
     "Use before/after imagery hook in opening paragraph; dermatologist quote in body",
     "2025-06-03"),
    (5, "FocusNest App",
     "This App Was Designed for People Who Hate Productivity Apps",
     "If you've tried every to-do app and still feel behind, it's not you. It's the apps.",
     "Vindication; finally understood; hope",
     "Problem/solution advertorial",
     "https://ref.example.com/focusnest-ad",
     "Try FocusNest Free — No Credit Card →",
     "Empathetic, slightly irreverent",
     "To Do", "Low",
     "ADHD-community tone works well; avoid hustle-culture language entirely",
     "2025-06-04"),
]


def build_sheet2(wb):
    ws = wb.create_sheet("Creative Research")
    disable_gridlines(ws)

    for col, w in CREA_WIDTHS.items():
        set_col_width(ws, get_column_letter(col), w)

    build_title_area(ws,
        "Creative Research Tracker",
        "Hooks, headlines, emotional angles, ad formats & reference creative — all in one place",
        TEAL)

    HDR_ROW = 8
    build_header_row(ws, HDR_ROW, CREA_COLS, TEAL)

    DATA_START = 9

    for i, sample in enumerate(CREA_SAMPLES):
        r = DATA_START + i
        row_color = WHITE if i % 2 == 0 else LIGHT_GRAY
        ws.row_dimensions[r].height = 44

        num, prod, headline, hook, emotion, fmt, ref, cta, tone, status, priority, notes, date = sample

        data_cell(ws, r, 2,  num,      row_color, bold=True)
        data_cell(ws, r, 3,  prod,     row_color, bold=True)
        data_cell(ws, r, 4,  headline, row_color, wrap=True)
        data_cell(ws, r, 5,  hook,     row_color, wrap=True)
        data_cell(ws, r, 6,  emotion,  row_color, wrap=True)
        data_cell(ws, r, 7,  fmt,      row_color)
        data_cell(ws, r, 8,  ref,      row_color)
        data_cell(ws, r, 9,  cta,      row_color, wrap=True)
        data_cell(ws, r, 10, tone,     row_color)
        status_pill(ws,   r, 11, status)
        ws.cell(row=r, column=11).fill = fill(
            {"Done": S_DONE_BG, "In Progress": S_PROG_BG,
             "To Do": S_TODO_BG, "Blocked": S_BLOK_BG}.get(status, S_TODO_BG))
        priority_pill(ws, r, 12, priority)
        ws.cell(row=r, column=12).fill = fill(
            {"High": P_HIGH_BG, "Medium": P_MED_BG, "Low": P_LOW_BG}.get(priority, P_MED_BG))
        data_cell(ws, r, 13, notes, row_color, wrap=True)
        data_cell(ws, r, 14, date,  row_color)

        ws.cell(row=r, column=1).fill = fill(TEAL)

    for i in range(5):
        r = DATA_START + 5 + i
        row_color = WHITE if i % 2 == 0 else LIGHT_GRAY
        ws.row_dimensions[r].height = 32
        for c in range(2, 15):
            cell = ws.cell(row=r, column=c)
            cell.fill   = fill(row_color)
            cell.border = bottom_border()
        ws.cell(row=r, column=1).fill = fill(TEAL)

    footer_bar(ws, DATA_START + 11,
               2, 14,
               "  ✦  Tip: Paste winning competitor ad URLs in 'Reference Ad / URL'. The best advertorials borrow structure, never copy. Always verify emotional angle matches the audience.",
               TEAL, TEAL_LIGHT[2:])
    ws.cell(row=DATA_START + 11, column=1).fill = fill(TEAL)

    return ws


# ─── Sheet 3: Dashboard ───────────────────────────────────────────────────────
WORKFLOW_STEPS = [
    (1, COBALT,   "Product Research",
     "Validate claims, USPs, pricing, proof points & competitors. Fill Sheet 1."),
    (2, "FF0369A1","Audience Research",
     "Define ICP, pain points, buying triggers & language patterns. Build empathy map."),
    (3, TEAL,     "Creative Research",
     "Source hooks, reference ads, emotional angles & format inspiration. Fill Sheet 2."),
    (4, "FF7C3AED","Advertorial Draft",
     "Write first draft using the Hook → Problem → Agitate → Solution → Proof → CTA framework."),
    (5, "FF16A34A","Fact-Check & Compliance",
     "Verify every claim has a cited source. Flag FTC-sensitive language. Legal sign-off."),
    (6, "FFD97706","Design & Layout Brief",
     "Create visual brief: hero image, pull-quotes, call-out boxes, button styles."),
    (7, SLATE,    "Publish, Track & Iterate",
     "Launch with UTM tracking. Monitor conversion rate, time-on-page, and scroll depth."),
]

PRINCIPLES = [
    ("Claim First, Proof Always",
     "Every claim in your advertorial must be backed by a source in Sheet 1. No source = no claim. Fabricated proof destroys trust and creates legal liability."),
    ("One Big Idea Per Advertorial",
     "Your advertorial should have a single, clear thesis. If you can't state your angle in one sentence, keep researching until you can."),
    ("The Reader, Not the Product",
     "The best advertorials read as journalism. Lead with the reader's world — their pain, their aspiration — not the product's features."),
    ("Iterate on What Works",
     "Track performance by format, emotional angle, and hook style. Duplicate winning structures across new products before reinventing the wheel."),
]


def build_sheet3(wb):
    ws = wb.create_sheet("Dashboard")
    disable_gridlines(ws)

    # Column widths
    col_widths = {1: 1.2, 2: 3, 3: 18, 4: 16, 5: 2, 6: 18, 7: 16, 8: 2,
                  9: 18, 10: 16, 11: 2, 12: 18, 13: 16, 14: 2, 15: 4, 16: 30}
    for c, w in col_widths.items():
        set_col_width(ws, get_column_letter(c), w)

    build_title_area(ws,
        "Advertorial Research Dashboard",
        "Strategic overview — track pipeline health, workflow progress & guiding principles",
        SLATE)

    # ── KPI Cards (row 9–14) ──────────────────────────────────────────────────
    KPI_DATA = [
        ("Products Researched", "5", COBALT),
        ("Creatives in Pipeline", "5", TEAL),
        ("Ready to Write", "2", "FF16A34A"),
        ("High Priority Items", "2", "FF9F1239"),
    ]

    CARD_COLS = [(2, 4), (6, 8), (9, 11), (12, 14)]   # (label_col, val_col) groups

    ws.row_dimensions[8].height = 12
    ws.row_dimensions[9].height = 6

    card_positions = [
        (2, 4),   # card 1: cols B–D
        (6, 8),   # card 2: cols F–H
        (9, 11),  # card 3: cols I–K
        (12, 14), # card 4: cols L–N
    ]

    for idx, ((start_col, end_col), (label, value, color)) in enumerate(
            zip(card_positions, KPI_DATA)):

        # Top accent border row
        for r in range(9, 14):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = fill("FFFFFFFF")

        # Top border accent
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=9, column=c)
            cell.border = Border(top=Side(style="thick", color=color[2:]))

        # Big number
        num_cell = ws.cell(row=10, column=start_col, value=value)
        num_cell.font      = font("Aptos", 32, bold=True, color=color)
        num_cell.alignment = align("left", "center")
        ws.row_dimensions[10].height = 44

        # Label
        lbl_cell = ws.cell(row=11, column=start_col, value=label.upper())
        lbl_cell.font      = font("Aptos", 8, bold=True, color=MUTED_GRAY)
        lbl_cell.alignment = align("left", "top")
        ws.row_dimensions[11].height = 16

        ws.row_dimensions[12].height = 6

    # Col A accent
    for r in range(9, 15):
        ws.cell(row=r, column=1).fill = fill(SLATE)

    # ── Workflow Section ──────────────────────────────────────────────────────
    ws.row_dimensions[14].height = 22
    sec_title = ws.cell(row=14, column=2, value="7-Step Advertorial Workflow")
    sec_title.font      = font("Aptos", 14, bold=True, color=DARK_INK)
    sec_title.alignment = align("left", "center")
    ws.cell(row=14, column=1).fill = fill(SLATE)

    ws.row_dimensions[15].height = 20
    # Header
    for col, label in [(2, "STEP"), (3, "PHASE"), (4, "✓"), (5, ""), (6, "DESCRIPTION")]:
        cell = ws.cell(row=15, column=col, value=label)
        cell.font      = font("Aptos", 8, bold=True, color="FFFFFFFF")
        cell.fill      = fill(DARK_SLATE)
        cell.alignment = align("center" if col in (2, 4) else "left", "center")
    ws.cell(row=15, column=1).fill = fill(SLATE)

    for i, (num, color, phase, desc) in enumerate(WORKFLOW_STEPS):
        r = 16 + i
        ws.row_dimensions[r].height = 28
        row_color = WHITE if i % 2 == 0 else LIGHT_GRAY

        # Step badge
        badge = ws.cell(row=r, column=2, value=str(num))
        badge.font      = font("Aptos", 11, bold=True, color="FFFFFFFF")
        badge.fill      = fill(color)
        badge.alignment = align("center", "center")

        # Phase
        ph = ws.cell(row=r, column=3, value=phase)
        ph.font      = font("Aptos", 10, bold=True, color=color)
        ph.fill      = fill(row_color)
        ph.alignment = align("left", "center")
        ph.border    = bottom_border()

        # Checkbox
        chk = ws.cell(row=r, column=4, value="☐")
        chk.font      = font("Aptos", 12, color=MUTED_GRAY)
        chk.fill      = fill(row_color)
        chk.alignment = align("center", "center")
        chk.border    = bottom_border()

        # Description
        desc_cell = ws.cell(row=r, column=6, value=desc)
        desc_cell.font      = font("Aptos", 9, color=MUTED_GRAY)
        desc_cell.fill      = fill(row_color)
        desc_cell.alignment = align("left", "center", wrap=True)
        desc_cell.border    = bottom_border()

        ws.cell(row=r, column=1).fill = fill(SLATE)

    # ── Strategic Principles ──────────────────────────────────────────────────
    PRIN_START = 16 + len(WORKFLOW_STEPS) + 2
    ws.row_dimensions[PRIN_START - 1].height = 24
    sec2_title = ws.cell(row=PRIN_START - 1, column=2, value="Strategic Principles")
    sec2_title.font      = font("Aptos", 14, bold=True, color=DARK_INK)
    sec2_title.alignment = align("left", "center")
    ws.cell(row=PRIN_START - 1, column=1).fill = fill(SLATE)

    for i, (title, explanation) in enumerate(PRINCIPLES):
        r = PRIN_START + i
        ws.row_dimensions[r].height = 38
        row_color = WHITE if i % 2 == 0 else LIGHT_GRAY

        title_cell = ws.cell(row=r, column=2, value=title)
        title_cell.font      = font("Aptos", 10, bold=True, color=DARK_INK)
        title_cell.fill      = fill(row_color)
        title_cell.alignment = align("left", "center")
        title_cell.border    = Border(
            left=Side(style="thick", color=COBALT[2:]),
            bottom=Side(style="thin", color=BORDER_CLR)
        )

        # Merge cols 3–8 for explanation
        exp_cell = ws.cell(row=r, column=4, value=explanation)
        exp_cell.font      = font("Aptos", 9, color=MUTED_GRAY)
        exp_cell.fill      = fill(row_color)
        exp_cell.alignment = align("left", "center", wrap=True)
        exp_cell.border    = bottom_border()

        ws.cell(row=r, column=1).fill = fill(SLATE)

    # Fill remaining col A cells
    for r in range(1, PRIN_START + len(PRINCIPLES) + 5):
        ws.cell(row=r, column=1).fill = fill(SLATE)

    return ws


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    wb = Workbook()
    # Remove default sheet
    default = wb.active
    wb.remove(default)

    build_sheet1(wb)
    build_sheet2(wb)
    build_sheet3(wb)

    wb.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
