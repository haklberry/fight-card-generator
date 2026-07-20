#!/usr/bin/env python3
"""
Fight card graphic generator for combat sports events.

Usage:
    python generate_card.py --excel fights.xlsx --logo logo.png --background background.png
    python generate_card.py --excel fights.xlsx --logo logo.png --background background.png --output ./output
"""

import argparse
import os
import re
import sys
import urllib.request
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency — run: pip install openpyxl")

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    sys.exit("Missing dependency — run: pip install 'Pillow>=9.2.0'")


# ─── Palette ──────────────────────────────────────────────────────────────────
NAVY   = (27,  42,  74)
RED    = (228,  2,  58)
WHITE  = (255, 255, 255)
ROW_W  = (255, 255, 255)   # even rows
ROW_A         = (218, 222, 232)   # odd rows — noticeably darker cool grey
BORDER        = (218, 222, 230)   # subtle row card border
PLACEHOLDER   = (175, 180, 195)   # muted blue-grey for placeholder text
SECTION_LINE  = (80,  100, 150)   # decorative line for section separators

# ─── Output formats (width × height in pixels) ────────────────────────────────
FORMATS = {
    "a4":        (2480, 3508),   # 210×297 mm @ 300 DPI
    "stories":   (1080, 1920),   # Instagram Stories
    "prispevek": (1080, 1350),   # Instagram příspěvek 4:5
}

# Vertical padding (top, bottom) in pixels added to each format
VERT_PAD = {
    "a4":        (0,   0),
    "stories":   (200, 200),
    "prispevek": (50,  50),
}

MAX_PER_PAGE = 12

# Font — bundled copy ships alongside this script in fonts/
_SCRIPT_DIR  = Path(__file__).resolve().parent
BUNDLED_FONT = _SCRIPT_DIR / "fonts" / "Montserrat[wght].ttf"
# Cache fallback for when the script is run from a copy without the fonts/ folder
FONT_CACHE   = Path.home() / ".cache" / "fightcard_fonts"
FONT_FILE    = FONT_CACHE / "Montserrat_variable.ttf"
FONT_URL     = "https://raw.githubusercontent.com/google/fonts/main/ofl/montserrat/Montserrat%5Bwght%5D.ttf"

# System TrueType fonts tried when Montserrat is unavailable (keeps correct size)
_SYSTEM_FONTS = [
    "/System/Library/Fonts/Helvetica.ttc",           # macOS
    "/System/Library/Fonts/SFNS.ttf",
    "/System/Library/Fonts/SFNSText.ttf",
    "C:/Windows/Fonts/arial.ttf",                    # Windows
    "C:/Windows/Fonts/calibri.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",    # Linux
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
]

# wght axis values for each role
WEIGHT_AXES = {
    "regular":   400,
    "medium":    500,
    "semibold":  600,
    "bold":      700,
    "extrabold": 800,
}

REQUIRED_COLS = {
    "fight_number", "red_name", "red_gym", "red_weight",
    "time", "blue_weight", "blue_name", "blue_gym",
}


# ─── Font loading ─────────────────────────────────────────────────────────────

def ensure_fonts() -> Path:
    """Return path to Montserrat variable font; bundled copy preferred, cache fallback."""
    if BUNDLED_FONT.exists():
        return BUNDLED_FONT
    # Not bundled — try the user cache, download if missing
    FONT_CACHE.mkdir(parents=True, exist_ok=True)
    if not FONT_FILE.exists():
        print("  Downloading Montserrat variable font…", end=" ", flush=True)
        try:
            urllib.request.urlretrieve(FONT_URL, FONT_FILE)
            print("ok")
        except Exception as exc:
            print(f"failed ({exc})")
    return FONT_FILE


def get_font(font_file: Path, weight: str, size: int) -> ImageFont.FreeTypeFont:
    """Load Montserrat at the given weight and size; fall back to a system TrueType font."""
    wght = WEIGHT_AXES.get(weight, 400)
    if font_file.exists():
        try:
            f = ImageFont.truetype(str(font_file), max(1, size))
            f.set_variation_by_axes([wght])
            return f
        except Exception:
            pass
    # Montserrat unavailable — try system fonts at the correct pt size
    for path in _SYSTEM_FONTS:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, max(1, size))
            except Exception:
                pass
    # Last resort: PIL bitmap default (size will be wrong but at least renders)
    return ImageFont.load_default()


# ─── Excel parsing ────────────────────────────────────────────────────────────

def load_excel(path: str) -> Tuple[Dict[str, str], List[Dict]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    if len(rows) < 2:
        sys.exit("Excel file has no data rows.")

    # Find the fight table header row (must contain "fight_number")
    fight_header_idx = None
    for i, row in enumerate(rows):
        cells_lower = [str(c).strip().lower() if c is not None else "" for c in row]
        if "fight_number" in cells_lower:
            fight_header_idx = i
            break

    if fight_header_idx is None:
        sys.exit("Excel is missing a header row with 'fight_number'.")

    # Read event info from the section above the fight table.
    # Labels use [title], [date], [ring] markers; values are in the next row, same column.
    event: Dict[str, str] = {"title": "", "date": "", "ring": ""}
    for row_i in range(fight_header_idx):
        row = rows[row_i]
        for col_i, cell in enumerate(row):
            if cell is None:
                continue
            cell_lower = str(cell).strip().lower()
            for key in ("title", "date", "ring"):
                if f"[{key}]" in cell_lower:
                    next_row_i = row_i + 1
                    if next_row_i < fight_header_idx:
                        val_row = rows[next_row_i]
                        if col_i < len(val_row) and val_row[col_i] is not None:
                            v = val_row[col_i]
                            event[key] = (
                                str(int(v)) if isinstance(v, float) and v.is_integer()
                                else str(v).strip()
                            )

    # Build column index from fight table header
    headers = [str(c).strip().lower() if c is not None else "" for c in rows[fight_header_idx]]
    missing = REQUIRED_COLS - set(headers)
    if missing:
        sys.exit(f"Excel is missing required columns: {', '.join(sorted(missing))}")

    col_idx = {h: i for i, h in enumerate(headers)}

    def val(row: tuple, name: str) -> str:
        i = col_idx.get(name)
        v = row[i] if i is not None and i < len(row) else None
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v).strip()

    items: List[Dict] = []
    current_section: str = ""

    for row in rows[fight_header_idx + 1:]:
        if not any(c for c in row if c is not None):
            continue
        red_name = val(row, "red_name")
        if not red_name:
            continue  # skip empty fight rows (pre-filled numbers but no fighters)
        num = val(row, "fight_number")
        if not num or not str(num).strip().lstrip("0").isdigit():
            continue  # skip hint rows whose fight_number is not a real number

        # Section header — insert when section column has a new non-empty value
        section = val(row, "section").strip() if "section" in col_idx else ""
        if section and section != current_section:
            current_section = section
            items.append({"type": "section", "name": section})

        items.append({
            "type":        "fight",
            "num":         num,
            "red_name":    red_name.upper(),
            "red_gym":     val(row, "red_gym").upper(),
            "red_weight":  val(row, "red_weight"),
            "time":        val(row, "time"),
            "blue_weight": val(row, "blue_weight"),
            "blue_name":   val(row, "blue_name").upper(),
            "blue_gym":    val(row, "blue_gym").upper(),
            "note":        val(row, "note"),
        })

    return event, items


# ─── Pagination ───────────────────────────────────────────────────────────────

def paginate_items(
    items: List[Dict],
    max_per_page: int = MAX_PER_PAGE,
    max_p2_plus: Optional[int] = None,
) -> List[List[Dict]]:
    """Split items into pages. Page 1 uses max_per_page; page 2+ uses max_p2_plus
    (defaults to max_per_page). Section headers are never orphaned at page end."""
    if max_p2_plus is None:
        max_p2_plus = max_per_page

    pages: List[List[Dict]] = []
    current: List[Dict] = []
    fight_count = 0

    def _max() -> int:
        return max_per_page if len(pages) == 0 else max_p2_plus

    for item in items:
        if item["type"] == "section":
            if fight_count >= _max():
                pages.append(current)
                current = []
                fight_count = 0
            current.append(item)
        else:  # fight
            if fight_count >= _max():
                if current and current[-1]["type"] == "section":
                    orphan = current.pop()
                    pages.append(current)
                    current = [orphan]
                else:
                    pages.append(current)
                    current = []
                fight_count = 0
            current.append(item)
            fight_count += 1

    if current:
        pages.append(current)

    return pages


# ─── Output filename ──────────────────────────────────────────────────────────

def out_filename(event: Dict, fmt: str, page: int, total: int) -> str:
    title = event.get("title", "CARD")
    date  = event.get("date", "").replace("/", "-")
    ring  = re.sub(r"\s+", "_", event.get("ring", "").strip())

    m = re.match(r"(\d+)\.\s*KOLO", title, re.IGNORECASE)
    prefix = f"{m.group(1)}_KOLO" if m else re.sub(r"[^\w]", "_", title)[:16]

    labels = {"a4": "A4", "stories": "stories", "prispevek": "prispevek"}
    page_suf = f"_page{page}" if total > 1 else ""
    ext = "pdf" if fmt == "a4" else "png"
    return f"{prefix}_{date}_{ring}_{labels[fmt]}{page_suf}.{ext}"


# ─── Layout metrics helpers ───────────────────────────────────────────────────

def _page_row_h(
    w: int, h: int, n_fights: int, n_sections: int,
    font_file: Path, pad_top: int, pad_bottom: int, show_header: bool,
) -> int:
    """Compute the row_h render_page would use — pure geometry, no drawing."""
    s = w / 1080
    _edge = round(59 * s)
    pt = pad_top    if pad_top    > 0 else _edge
    pb = pad_bottom if pad_bottom > 0 else _edge

    _msr = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    def _TH(f: ImageFont.FreeTypeFont) -> int:
        bb = _msr.textbbox((0, 0), "Ag", font=f)
        return bb[3] - bb[1]

    header_h     = round(210 * s) if show_header else 0
    f_ch         = get_font(font_file, "semibold",  max(1, round(COLHDR_SZ  * s)))
    colhdr_strip = _TH(f_ch) + round(22 * s)
    y_rows       = pt + header_h + colhdr_strip + round(8 * s)

    f_nm  = get_font(font_file, "extrabold", max(1, round(NAME_SZ * s)))
    f_gym = get_font(font_file, "medium",    max(1, round(GYM_SZ  * s)))
    v_gap    = round(4 * s)
    ideal_rh = _TH(f_nm) + v_gap + _TH(f_gym) + round(16 * s)

    f_pn     = get_font(font_file, "medium",   max(1, round(PAGENUM_SZ * s)))
    pnum_res = _TH(f_pn) + round(PNUM_MARGIN_ABOVE * s) + round(PNUM_MARGIN_BELOW * s)

    f_sec    = get_font(font_file, "semibold", max(1, round(18 * s)))
    sec_slot = round(12 * s) + _TH(f_sec) + round(12 * s)
    row_gap  = round(5 * s)

    y_end  = h - pb - pnum_res
    avail  = y_end - y_rows - n_sections * sec_slot
    raw_rh = (avail - row_gap * max(n_fights - 1, 0)) // max(n_fights, 1)
    return max(ideal_rh, min(round(160 * s), raw_rh))


def calc_page1_row_h(
    w: int, h: int, page1_items: List[Dict],
    font_file: Path, pad_top: int = 0, pad_bottom: int = 0,
) -> int:
    """Row height page 1 would use — call before paginating page 2+."""
    n_f = sum(1 for x in page1_items if x["type"] == "fight")
    n_s = sum(1 for x in page1_items if x["type"] == "section")
    return _page_row_h(w, h, n_f, n_s, font_file, pad_top, pad_bottom, show_header=True)


def calc_max_fights_p2(
    w: int, h: int, fixed_row_h: int,
    font_file: Path, pad_top: int = 0, pad_bottom: int = 0,
) -> int:
    """Max fights that fit on page 2+ with the given fixed row_h (no header)."""
    s = w / 1080
    _edge = round(59 * s)
    pt = pad_top    if pad_top    > 0 else _edge
    pb = pad_bottom if pad_bottom > 0 else _edge

    _msr = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    def _TH(f: ImageFont.FreeTypeFont) -> int:
        bb = _msr.textbbox((0, 0), "Ag", font=f)
        return bb[3] - bb[1]

    f_ch         = get_font(font_file, "semibold", max(1, round(COLHDR_SZ * s)))
    colhdr_strip = _TH(f_ch) + round(22 * s)
    y_rows       = pt + colhdr_strip + round(8 * s)   # no header_h
    f_pn         = get_font(font_file, "medium",   max(1, round(PAGENUM_SZ * s)))
    pnum_res     = _TH(f_pn) + round(PNUM_MARGIN_ABOVE * s) + round(PNUM_MARGIN_BELOW * s)
    row_gap      = round(5 * s)
    available    = h - pb - pnum_res - y_rows
    return max(1, (available + row_gap) // (fixed_row_h + row_gap))


# ─── Drawing helpers ──────────────────────────────────────────────────────────

def text_size(draw: ImageDraw.Draw, text: str, f: ImageFont.FreeTypeFont) -> Tuple[int, int]:
    bb = draw.textbbox((0, 0), text, font=f)
    return bb[2] - bb[0], bb[3] - bb[1]


def clip_text(draw: ImageDraw.Draw, text: str, f: ImageFont.FreeTypeFont, max_w: int) -> str:
    if text_size(draw, text, f)[0] <= max_w:
        return text
    while len(text) > 1 and text_size(draw, text[:-1] + "…", f)[0] > max_w:
        text = text[:-1]
    return (text[:-1] + "…") if len(text) > 1 else text


def clip_name(draw: ImageDraw.Draw, full_name: str, f: ImageFont.FreeTypeFont, max_w: int) -> str:
    """Shorten first name to initial if name overflows; keep surname intact."""
    if text_size(draw, full_name, f)[0] <= max_w:
        return full_name
    parts = full_name.split(None, 1)
    if len(parts) == 2:
        first, rest = parts
        short = f"{first[0]}. {rest}"
        if text_size(draw, short, f)[0] <= max_w:
            return short
        if text_size(draw, rest, f)[0] <= max_w:
            return rest
        return clip_text(draw, rest, f, max_w)
    return clip_text(draw, full_name, f, max_w)


def fmt_time(t: str) -> str:
    """Strip trailing 'min' so '3x2 min' → '3x2'."""
    return t.replace(" min", "").replace("min", "").strip()


def draw_centered(draw: ImageDraw.Draw, cx: int, cy: int, text: str, f, fill):
    w, h = text_size(draw, text, f)
    draw.text((cx - w // 2, cy - h // 2), text, font=f, fill=fill)


def draw_halign(draw: ImageDraw.Draw, cx: int, ty: int, text: str, f, fill):
    """Draw text horizontally centred at cx, top edge at ty."""
    w, _ = text_size(draw, text, f)
    draw.text((cx - w // 2, ty), text, font=f, fill=fill)


# ─── Main renderer ────────────────────────────────────────────────────────────

# Font sizes at 1080 px base (before s-scaling). GYM_SZ kept, rest -2 pt.
NAME_SZ   = 28   # ExtraBold — fighter name
GYM_SZ    = 20   # Medium    — gym name (unchanged)
WT_SZ     = 26   # Bold      — weight value
TIME_SZ   = 24   # Bold      — round format
NUM_SZ    = 24   # Bold      — fight number
DATE_SZ           = 38   # Bold      — date / ring
COLHDR_SZ         = 22   # SemiBold  — column labels
TITLE_SZ          = 60   # ExtraBold — event title (auto-shrinks to fit)
PAGENUM_SZ        = 18   # Medium    — page number at bottom
PNUM_MARGIN_ABOVE = 20   # pts of clear space above page-number text (fights end here)
PNUM_MARGIN_BELOW =  8   # pts between page-number text and bottom-padding edge


def render_page(
    w: int, h: int,
    items: List[Dict],
    event: Dict[str, str],
    font_file: Path,
    bg: Image.Image,
    logo: Optional[Image.Image],
    pad_top: int = 0,
    pad_bottom: int = 0,
    show_header: bool = True,
    page_num: int = 1,
    total_pages: int = 1,
    row_h_override: Optional[int] = None,
) -> Image.Image:

    s  = w / 1080   # horizontal scale relative to 1080 px base

    # Uniform edge margin for A4 (pad_top/pad_bottom == 0); stories/prispevek keep their VERT_PAD
    _edge = round(59 * s)
    pt = pad_top    if pad_top    > 0 else _edge
    pb = pad_bottom if pad_bottom > 0 else _edge

    # ── Measurement helpers ───────────────────────────────────────────────────
    _msr = ImageDraw.Draw(Image.new("RGBA", (1, 1)))

    def F(wt: str, sz: float) -> ImageFont.FreeTypeFont:
        return get_font(font_file, wt, max(1, round(sz)))

    def TW(f, text: str) -> int:
        bb = _msr.textbbox((0, 0), text, font=f)
        return bb[2] - bb[0]

    def TH(f) -> int:
        bb = _msr.textbbox((0, 0), "Ag", font=f)
        return bb[3] - bb[1]

    # ── Step 1: layout geometry ───────────────────────────────────────────────
    pad_l      = round(30 * s)
    pad_r      = round(30 * s)
    card_pad_x = round(12 * s)
    header_h   = round(210 * s) if show_header else 0

    f_colhdr      = F("semibold", COLHDR_SZ * s)
    colhdr_strip  = TH(f_colhdr) + round(22 * s)
    _ch_bb        = _msr.textbbox((0, 0), "VÁHA", font=f_colhdr)
    _ch_h         = _ch_bb[3] - _ch_bb[1]
    _ch_top       = _ch_bb[1]
    colhdr_text_y = pt + header_h + (colhdr_strip - _ch_h) // 2 - _ch_top
    nav_total     = header_h + colhdr_strip          # height of navy section
    row_top_gap   = round(8 * s)
    y_rows        = pt + nav_total + row_top_gap     # absolute y of first row card

    f_name_ref = F("extrabold", NAME_SZ * s)
    f_gym_ref  = F("medium",    GYM_SZ  * s)
    v_gap      = round(4 * s)
    block_h    = TH(f_name_ref) + v_gap + TH(f_gym_ref)
    row_v_pad  = round(16 * s)
    ideal_rh   = block_h + row_v_pad

    f_pnum_fnt    = F("medium", round(PAGENUM_SZ * s))
    pnum_h        = TH(f_pnum_fnt)
    pnum_reserved = pnum_h + round(PNUM_MARGIN_ABOVE * s) + round(PNUM_MARGIN_BELOW * s)

    row_gap   = round(5 * s)
    y_end     = h - pb - pnum_reserved
    available = y_end - y_rows

    # Section header metrics (used for height budget and drawing)
    f_sec         = F("semibold", round(18 * s))
    sec_text_h    = TH(f_sec)
    section_slot_h = round(12 * s) + sec_text_h + round(12 * s)  # pre + text + post

    n_fights   = sum(1 for it in items if it["type"] == "fight")
    n_sections = sum(1 for it in items if it["type"] == "section")

    available_for_fights = available - n_sections * section_slot_h
    raw_rh = (available_for_fights - row_gap * max(n_fights - 1, 0)) // max(n_fights, 1)
    if row_h_override is not None:
        row_h = row_h_override
    else:
        row_h = max(ideal_rh, min(round(160 * s), raw_rh))

    rs = s * min(1.0, row_h / ideal_rh)

    # ── Step 2: row-content fonts ─────────────────────────────────────────────
    f_name = F("extrabold", NAME_SZ * rs)
    f_gym  = F("medium",    GYM_SZ  * rs)
    f_wt   = F("bold",      WT_SZ   * rs)
    f_time = F("bold",      TIME_SZ * rs)
    f_num  = F("bold",      NUM_SZ  * rs)

    # ── Step 3: column x positions ────────────────────────────────────────────
    tim_cx   = w // 2
    wt_ref   = TW(f_wt,   "99 kg")
    time_ref = TW(f_time, "3x2")
    col_gap  = round(16 * s)

    rwt_cx = tim_cx - (time_ref // 2 + col_gap + wt_ref // 2)
    bwt_cx = tim_cx + (time_ref // 2 + col_gap + wt_ref // 2)

    num_ref = TW(f_num, "15")
    red_lx  = pad_l + card_pad_x + num_ref + round(10 * s)
    num_cx  = (pad_l + red_lx) // 2   # center of number column
    red_rx  = rwt_cx - wt_ref // 2 - col_gap
    blue_lx = bwt_cx + wt_ref // 2 + col_gap
    blue_rx = w - pad_r - card_pad_x

    # ── Step 4: header fonts ──────────────────────────────────────────────────
    title_max_w = w - round(44 * s)
    f_title = F("extrabold", TITLE_SZ * s)
    tw_raw  = TW(f_title, event.get("title", ""))
    if tw_raw > title_max_w:
        f_title = F("extrabold", round(TITLE_SZ * s * title_max_w / tw_raw))
    f_date = F("bold", DATE_SZ * s)

    # ── Canvas — custom background or solid navy ─────────────────────────────
    if bg is not None:
        bg_w, bg_h = bg.size
        scale  = max(w / bg_w, h / bg_h)
        nw, nh = round(bg_w * scale), round(bg_h * scale)
        bg_fit = bg.convert("RGBA").resize((nw, nh), Image.LANCZOS)
        ox, oy = (nw - w) // 2, (nh - h) // 2
        canvas = bg_fit.crop((ox, oy, ox + w, oy + h)).copy()
    else:
        canvas = Image.new("RGBA", (w, h), NAVY + (255,))
    draw = ImageDraw.Draw(canvas)

    # ── Full navy block (title + date + column labels), offset by pt ──────────
    draw.rectangle([0, pt, w, pt + nav_total], fill=NAVY)

    if show_header:
        # Title — centred across full width
        title  = event.get("title", "")
        tw, th = text_size(draw, title, f_title)
        ty     = pt + round(36 * s)
        draw.text((w // 2 - tw // 2, ty), title, font=f_title, fill=WHITE)

        # Date | Ring — centred below title
        sub   = f"{event.get('date', '')}  |  {event.get('ring', '')}"
        sw, _ = text_size(draw, sub, f_date)
        dy    = ty + th + round(10 * s)
        draw.text((w // 2 - sw // 2, dy), sub, font=f_date, fill=RED)

    # ── Column header labels ───────────────────────────────────────────────────
    y_ch = colhdr_text_y

    draw.text((red_lx, y_ch), "ČERVENÝ ROH", font=f_colhdr, fill=RED)

    for cx, lbl in [(rwt_cx, "VÁHA"), (tim_cx, "ČAS"), (bwt_cx, "VÁHA")]:
        lw, _ = text_size(draw, lbl, f_colhdr)
        draw.text((cx - lw // 2, y_ch), lbl, font=f_colhdr, fill=WHITE)

    lbl_blue = "MODRÝ ROH"
    lw, _ = text_size(draw, lbl_blue, f_colhdr)
    draw.text((blue_rx - lw, y_ch), lbl_blue, font=f_colhdr, fill=WHITE)

    # ── Fight rows + section headers ───────────────────────────────────────────
    y         = y_rows
    fight_idx = 0   # separate counter for alternating row colour

    for item in items:

        if item["type"] == "section":
            # ── Section header: decorative line + centred label ──────────────
            sec_text = item["name"].upper()
            tw, _    = text_size(draw, sec_text, f_sec)
            y_text   = y + round(12 * s)          # text top inside slot
            y_mid    = y_text + sec_text_h // 2   # midpoint for horizontal lines
            pad_in   = round(10 * s)
            gap_txt  = round(14 * s)
            lw       = max(1, round(2 * s))
            lx1 = pad_l + pad_in
            lx2 = w // 2 - tw // 2 - gap_txt
            rx1 = w // 2 + tw // 2 + gap_txt
            rx2 = w - pad_r - pad_in
            if lx2 > lx1:
                draw.line([(lx1, y_mid), (lx2, y_mid)], fill=RED, width=lw)
            draw_halign(draw, w // 2, y_text, sec_text, f_sec, RED)
            if rx2 > rx1:
                draw.line([(rx1, y_mid), (rx2, y_mid)], fill=RED, width=lw)
            y += section_slot_h
            continue

        # ── Fight card ────────────────────────────────────────────────────────
        fight = item
        ry0, ry1 = y, y + row_h
        cy   = (ry0 + ry1) // 2
        rr   = round(10 * s)
        fill = ROW_W if fight_idx % 2 == 0 else ROW_A
        fight_idx += 1

        draw.rounded_rectangle([pad_l, ry0, w - pad_r, ry1], radius=rr, fill=fill)

        draw_centered(draw, num_cx, cy, fight["num"], f_num, NAVY)

        _, nnh = text_size(draw, "Ag", f_name)
        _, gnh = text_size(draw, "Ag", f_gym)
        v_blk  = nnh + v_gap + gnh
        ny     = cy - v_blk // 2
        gy     = ny + nnh + v_gap

        rname = clip_name(draw, fight["red_name"], f_name, red_rx - red_lx)
        rgym  = clip_text(draw, fight["red_gym"],  f_gym,  red_rx - red_lx)
        draw.text((red_lx, ny), rname, font=f_name, fill=RED)
        draw.text((red_lx, gy), rgym,  font=f_gym,  fill=NAVY)

        t_str = fmt_time(fight["time"])
        draw_halign(draw, rwt_cx, ny, f"{fight['red_weight']} kg",  f_wt,   NAVY)
        draw_halign(draw, tim_cx, ny, t_str,                        f_time, NAVY)
        draw_halign(draw, bwt_cx, ny, f"{fight['blue_weight']} kg", f_wt,   NAVY)

        f_note_lbl = F("extrabold", GYM_SZ * rs)
        note_txt   = fight.get("note", "").strip()
        note_disp  = note_txt if note_txt else "děti, chrániče"
        draw_halign(draw, tim_cx, gy, note_disp, f_note_lbl, RED)

        bname = clip_name(draw, fight["blue_name"], f_name, blue_rx - blue_lx)
        bgym  = clip_text(draw, fight["blue_gym"],  f_gym,  blue_rx - blue_lx)
        bw, _ = text_size(draw, bname, f_name)
        gw, _ = text_size(draw, bgym,  f_gym)
        draw.text((blue_rx - bw, ny), bname, font=f_name, fill=NAVY)
        draw.text((blue_rx - gw, gy), bgym,  font=f_gym,  fill=NAVY)

        y += row_h + row_gap

    # ── Page number at bottom centre ──────────────────────────────────────────
    pnum_text = f"str. {page_num} z {total_pages}"
    ptw, _    = text_size(draw, pnum_text, f_pnum_fnt)
    pnum_y    = h - pb - round(PNUM_MARGIN_BELOW * s) - pnum_h
    draw.text((w // 2 - ptw // 2, pnum_y), pnum_text, font=f_pnum_fnt, fill=WHITE)

    return canvas.convert("RGB")


# ─── Entry point ──────────────────────────────────────────────────────────────

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Generate fight card graphics (A4 PDF + Instagram Stories + Carousel)."
    )
    ap.add_argument("--excel",      required=True, help="Path to .xlsx fight data file")
    ap.add_argument("--logo",       default=None,  help="Path to logo image (PNG/JPG) — optional")
    ap.add_argument("--background", default=None,  help="Path to background image (PNG/JPG) — currently unused; solid navy is used")
    ap.add_argument("--output",     default=".",   help="Output directory (default: current)")
    args = ap.parse_args()

    if not os.path.exists(args.excel):
        sys.exit(f"Error: Excel file not found: {args.excel}")
    if args.logo and not os.path.exists(args.logo):
        sys.exit(f"Error: Logo not found: {args.logo}")

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    print("Loading fonts (cached after first run)…")
    font_file = ensure_fonts()

    print("Loading images…")
    bg_img   = Image.open(args.background).convert("RGBA") if args.background else None
    logo_img = Image.open(args.logo).convert("RGBA") if args.logo else None

    print("Reading fight data…")
    event, items = load_excel(args.excel)
    fights = [x for x in items if x["type"] == "fight"]
    if not fights:
        sys.exit("No fight rows found in the Excel file.")
    print(f"  {len(fights)} fights  ·  {event.get('title')}  ·  {event.get('date')}  ·  {event.get('ring')}")

    pages = paginate_items(items)

    for pnum, page_items in enumerate(pages, 1):
        page_fights = [x for x in page_items if x["type"] == "fight"]
        print(f"\nPage {pnum}/{len(pages)}  ({len(page_fights)} fights):")
        for fmt, (fw, fh) in FORMATS.items():
            print(f"  [{fmt:<10}]  {fw}×{fh} px … ", end="", flush=True)
            pt, pb = VERT_PAD.get(fmt, (0, 0))
            img    = render_page(fw, fh, page_items, event, font_file, bg_img, logo_img,
                                 pad_top=pt, pad_bottom=pb,
                                 show_header=(pnum == 1),
                                 page_num=pnum, total_pages=len(pages))
            name   = out_filename(event, fmt, pnum, len(pages))
            dest = out_dir / name
            if fmt == "a4":
                img.save(str(dest), "PDF", resolution=300)
            else:
                img.save(str(dest), "PNG", optimize=True)
            print(f"saved → {name}")

    print("\nAll done.")


if __name__ == "__main__":
    main()
