#!/usr/bin/env python3
"""
Fight Card Generator — webová aplikace
Spustit: streamlit run app_karta.py
"""

import io
import os
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
from PIL import Image

from generate_card import (
    FONT_FILE, FORMATS, VERT_PAD,
    ensure_fonts, load_excel, paginate_items, render_page,
    calc_page1_row_h, calc_max_fights_p2,
)

# ── Konfigurace stránky ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="Fight Card Generator",
    page_icon="🥊",
    layout="centered",
)

st.markdown("""
<style>
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {
    background-color: #1B2A4A !important;
    border-color: #1B2A4A !important;
    color: white !important;
}
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover {
    background-color: #243660 !important;
    border-color: #243660 !important;
}
</style>
""", unsafe_allow_html=True)

st.title("Fight Card Generator 🥊")
st.markdown(
    "Nahraj vyplněnou Excel tabulku se zápasy a stáhni hotové grafiky "
    "pro všechny formáty (A4 PDF, Stories, Příspěvek 4:5). "
    "Šablonu tabulky stáhneš tlačítkem níže — stačí ji vyplnit a nahrát zpět."
)

# ── Fonty (stáhnout jednou, pak z cache) ─────────────────────────────────────
with st.spinner("Načítám fonty (jen při prvním spuštění)…"):
    font_file = ensure_fonts()

# ── Tlačítko pro stažení šablony ─────────────────────────────────────────────
template_path = Path(__file__).parent / "template_zapasy.xlsx"
if template_path.exists():
    with open(template_path, "rb") as f:
        st.download_button(
            "⬇️ Stáhnout šablonu Excel",
            data=f.read(),
            file_name="template_zapasy.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Vzorová tabulka se správnými sloupci a ukázkovými daty",
        )

st.divider()

# ── Nahrání souboru ───────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Nahraj vyplněnou Excel tabulku (.xlsx)",
    type=["xlsx"],
    label_visibility="visible",
)

bg_uploaded = st.file_uploader(
    "Vlastní pozadí — volitelné (PNG nebo JPG)",
    type=["png", "jpg", "jpeg"],
    label_visibility="visible",
    help="Doporučená velikost: min. 2480 × 3508 px (pokryje A4 formát). "
         "Pozadí se ořízne ze středu pro každý formát zvlášť — nebude roztaženo.",
)
bg_image = None
if bg_uploaded:
    from PIL import Image as _PIL
    bg_image = _PIL.open(io.BytesIO(bg_uploaded.getvalue())).convert("RGBA")

if not uploaded:
    st.info("👆 Nahraj Excel tabulku a klikni **Generovat grafiky**.")
    st.stop()

# ── Načtení dat ───────────────────────────────────────────────────────────────
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
    tmp.write(uploaded.getvalue())
    tmp_path = tmp.name

try:
    event, items = load_excel(tmp_path)
    fights = [x for x in items if x["type"] == "fight"]
except SystemExit as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"Chyba při čtení souboru: {e}")
    st.stop()
finally:
    os.unlink(tmp_path)

if not fights:
    st.error("V tabulce nejsou žádné zápasy.")
    st.stop()

# ── Přehled načtených dat ─────────────────────────────────────────────────────
st.success(f"✅ Načteno **{len(fights)} zápasů**")

st.markdown(f"**Název akce:** {event.get('title', '—')}")
col_d, col_r = st.columns(2)
col_d.markdown(f"**Datum:** {event.get('date', '—')}")
col_r.markdown(f"**Popis akce:** {event.get('ring', '—')}")

# ── Výpočet fixní výšky řádku a kapacity stran 2+ ───────────────────────────
# Fáze 1: dočasné stránkování → zjistíme obsah str. 1
_pages_tmp = paginate_items(items)
_page1_items = _pages_tmp[0] if _pages_tmp else []

# Pro každý formát spočítáme row_h str. 1 a kapacitu stran 2+
_row_h_per_fmt: dict = {}
_p2_fits_per_fmt: dict = {}
for _fmt, (_fw, _fh) in FORMATS.items():
    _pt, _pb = VERT_PAD.get(_fmt, (0, 0))
    _rh = calc_page1_row_h(_fw, _fh, _page1_items, font_file, _pt, _pb)
    _fits = calc_max_fights_p2(_fw, _fh, _rh, font_file, _pt, _pb)
    _row_h_per_fmt[_fmt] = _rh
    _p2_fits_per_fmt[_fmt] = _fits

_max_p2 = min(_p2_fits_per_fmt.values())   # nejmenší kapacita přes formáty

# Fáze 2: finální stránkování s reálnou kapacitou stran 2+
pages = paginate_items(items, max_p2_plus=_max_p2)

if len(pages) > 1:
    st.info(f"Celkem {len(fights)} zápasů → vygenerují se **{len(pages)} stránky** pro každý formát.")

st.divider()

# ── Generování ────────────────────────────────────────────────────────────────
# Pokud se nahraje nový soubor, zrušíme starý výsledek.
_file_id = getattr(uploaded, "file_id", None)
if st.session_state.get("_last_file_id") != _file_id:
    st.session_state.pop("gen", None)
    st.session_state["_last_file_id"] = _file_id

if st.button("🎨 Generovat grafiky", type="primary", use_container_width=True):
    progress_bar = st.progress(0, text="Generuji…")
    total_steps  = len(pages) * len(FORMATS)
    step         = 0

    rendered = {fmt: [] for fmt in FORMATS}

    for pnum, page_items in enumerate(pages, 1):
        for fmt, (fw, fh) in FORMATS.items():
            pt, pb = VERT_PAD.get(fmt, (0, 0))
            progress_bar.progress(
                step / total_steps,
                text=f"Generuji {fmt.upper()} — strana {pnum}/{len(pages)}…",
            )
            img = render_page(fw, fh, page_items, event, font_file,
                              bg_image, None, pad_top=pt, pad_bottom=pb,
                              show_header=(pnum == 1),
                              page_num=pnum, total_pages=len(pages),
                              row_h_override=_row_h_per_fmt[fmt])
            rendered[fmt].append(img)
            step += 1

    progress_bar.progress(0.95, text="Ukládám soubory…")

    # PDF — všechny stránky jako jeden soubor
    a4_imgs  = [img.convert("RGB") for img in rendered["a4"]]
    pdf_buf  = io.BytesIO()
    a4_imgs[0].save(pdf_buf, "PDF", resolution=300, save_all=True,
                    append_images=a4_imgs[1:])
    pdf_bytes = pdf_buf.getvalue()

    # Stories PNG
    stories_files = []
    for i, img in enumerate(rendered["stories"], 1):
        b = io.BytesIO(); img.save(b, "PNG", optimize=True)
        stories_files.append((f"Pribeh_str{i}.png", b.getvalue()))

    # Příspěvek PNG
    prispevek_files = []
    for i, img in enumerate(rendered["prispevek"], 1):
        b = io.BytesIO(); img.save(b, "PNG", optimize=True)
        prispevek_files.append((f"Prispevek_{i}.png", b.getvalue()))

    # Celkový ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("PDF A4 tisk.pdf", pdf_bytes)
        for name, data in stories_files:
            zf.writestr(name, data)
        for name, data in prispevek_files:
            zf.writestr(name, data)

    # Náhledy — plná kvalita, max. 1080 px šířka
    def _preview_png(img):
        if img.width > 1080:
            nh = round(1080 * img.height / img.width)
            img = img.resize((1080, nh), Image.LANCZOS)
        b = io.BytesIO(); img.save(b, "PNG"); return b.getvalue()

    st.session_state["gen"] = {
        "zip":       zip_buf.getvalue(),
        "pdf":       pdf_bytes,
        "stories":   stories_files,
        "prispevek": prispevek_files,
        "previews":  [(fmt, _preview_png(rendered[fmt][0]))
                      for fmt in ("a4", "stories", "prispevek")],
    }

    progress_bar.progress(1.0, text="Hotovo!")

# ── Výsledky (přetrvají i po kliknutí na stažení) ────────────────────────────
if "gen" in st.session_state:
    gen = st.session_state["gen"]

    def _fmt_zip(files):
        b = io.BytesIO()
        with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in files:
                zf.writestr(name, data)
        return b.getvalue()

    st.divider()

    # Řádek 1: Stáhnout vše
    st.download_button(
        "⬇️ Stáhnout vše (ZIP)",
        data=gen["zip"], file_name="fight_card.zip",
        mime="application/zip", type="primary", use_container_width=True,
    )

    # Řádek 2: Stažení po typech
    col_pdf, col_st, col_pr = st.columns(3)

    col_pdf.download_button(
        "⬇️ PDF A4",
        data=gen["pdf"], file_name="PDF A4 tisk.pdf",
        mime="application/pdf", type="primary", use_container_width=True,
    )

    if len(gen["stories"]) == 1:
        col_st.download_button(
            "⬇️ Příběh (Stories)",
            data=gen["stories"][0][1], file_name=gen["stories"][0][0],
            mime="image/png", type="primary", use_container_width=True,
        )
    else:
        col_st.download_button(
            "⬇️ Příběh (Stories ZIP)",
            data=_fmt_zip(gen["stories"]), file_name="Pribeh.zip",
            mime="application/zip", type="primary", use_container_width=True,
        )

    if len(gen["prispevek"]) == 1:
        col_pr.download_button(
            "⬇️ Příspěvek",
            data=gen["prispevek"][0][1], file_name=gen["prispevek"][0][0],
            mime="image/png", type="primary", use_container_width=True,
        )
    else:
        col_pr.download_button(
            "⬇️ Příspěvek (ZIP)",
            data=_fmt_zip(gen["prispevek"]), file_name="Prispevek.zip",
            mime="application/zip", type="primary", use_container_width=True,
        )

    # ── Náhled plná kvalita ───────────────────────────────────────────────────
    st.subheader("Náhled (strana 1)")
    fmt_labels = {"a4": "A4", "stories": "Stories", "prispevek": "Příspěvek"}
    cols = st.columns(3)
    for col, (fmt, png_data) in zip(cols, gen["previews"]):
        col.image(png_data, caption=fmt_labels[fmt], use_column_width=True)
