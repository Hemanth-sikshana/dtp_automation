"""
app.py - Streamlit front-end for dtp_pipeline.py

A point-and-click UI so a non-developer can:
  1. Upload a raw/unedited .docx
  2. Optionally tweak page size, colors, fonts, cover page, footer, contents
  3. Click Run
  4. Download the formatted result as .docx and/or .pdf

Run it with:
    pip install -r requirements.txt
    streamlit run app.py

Keep this file, dtp_pipeline.py, and the assets/ folder together in the
same directory -- the app imports dtp_pipeline.py directly and reuses the
bundled logos/illustration as sensible defaults.
"""

import os
import subprocess
import tempfile

import streamlit as st

import dtp_pipeline as dtp


st.set_page_config(page_title="Workbook DTP Formatter", page_icon="📘", layout="wide")

# ----------------------------------------------------------------------------
# Light visual polish (this styles the *app chrome*, not the generated
# document -- the document's own look is controlled by the settings below)
# ----------------------------------------------------------------------------
st.markdown("""
<style>
    .block-container {padding-top: 2rem; max-width: 1150px;}
    div[data-testid="stMetric"] {
        background: #F8F9FA; border: 1px solid #E5E7EB;
        border-radius: 10px; padding: 12px 14px;
    }
    .dtp-banner {
        background: linear-gradient(90deg, #1A2B4C 0%, #2E4670 100%);
        color: white; padding: 22px 28px; border-radius: 12px;
        margin-bottom: 22px;
    }
    .dtp-banner h1 {color: white; margin: 0; font-size: 1.6rem;}
    .dtp-banner p {color: #D7DEEA; margin: 6px 0 0 0; font-size: 0.95rem;}
    .asset-caption {font-size: 0.78rem; color: #6B7280; text-align: center;}
    section[data-testid="stSidebar"] .block-container {padding-top: 1.4rem;}
    div[data-testid="stFileUploaderDropzone"] {border-radius: 10px;}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Paper size presets (width_mm, height_mm) -- "any page size", not just A4
# ----------------------------------------------------------------------------
PAGE_SIZES = {
    "A3": (297, 420), "A4": (210, 297), "A5": (148, 210),
    "B4 (JIS)": (257, 364), "B5 (JIS)": (182, 257),
    "US Letter": (215.9, 279.4), "US Legal": (215.9, 355.6),
    "Custom...": None,
}

DC = dtp.DEFAULT_CONFIG

# ----------------------------------------------------------------------------
# Persistent per-session working directory (holds uploads + outputs)
# ----------------------------------------------------------------------------
if "work_dir" not in st.session_state:
    st.session_state.work_dir = tempfile.mkdtemp(prefix="dtp_session_")
WORK_DIR = st.session_state.work_dir


def save_uploaded(file, filename):
    """Write a Streamlit UploadedFile to disk in the session work dir."""
    if file is None:
        return None
    path = os.path.join(WORK_DIR, filename)
    with open(path, "wb") as f:
        f.write(file.getbuffer())
    return path


def asset_picker(label, key, default_path, help_text=""):
    """Show the currently-selected default image (if any) plus an uploader
    to replace it. Returns the path that should actually be used."""
    c_img, c_up = st.columns([1, 2])
    with c_img:
        if default_path and os.path.isfile(default_path):
            st.image(default_path, use_container_width=True)
            st.markdown('<p class="asset-caption">current default</p>', unsafe_allow_html=True)
        else:
            st.markdown('<p class="asset-caption">(none set)</p>', unsafe_allow_html=True)
    with c_up:
        upload = st.file_uploader(label, type=["png", "jpg", "jpeg"], key=key, help=help_text)
    if upload is not None:
        path = save_uploaded(upload, f"{key}{os.path.splitext(upload.name)[1]}")
        st.image(path, width=140, caption="will be used instead")
        return path
    return default_path


# ============================================================================
# HEADER
# ============================================================================
st.markdown("""
<div class="dtp-banner">
    <h1>📘 Workbook DTP Formatter</h1>
    <p>Upload a raw, unedited Word document and get a professionally
    designed workbook back -- cover page, styled tables, page numbers,
    a table of contents. No coding needed. Everything in the sidebar is
    optional; the built-in defaults already look good.</p>
</div>
""", unsafe_allow_html=True)

uploaded_docx = st.file_uploader("📤 Upload the raw .docx to format", type=["docx"])

# ============================================================================
# SIDEBAR -- all design settings, neatly grouped
# ============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Settings")
    st.caption("Change anything you like -- it only affects this run.")

    # ---- Page setup -------------------------------------------------------
    st.markdown("### 📐 Page setup")
    size_choice = st.selectbox("Paper size", list(PAGE_SIZES.keys()), index=1)
    if size_choice == "Custom...":
        cw1, cw2 = st.columns(2)
        width_mm = cw1.number_input("Width (mm)", min_value=50.0, max_value=1000.0,
                                     value=float(DC["document_setup"]["width_mm"]))
        height_mm = cw2.number_input("Height (mm)", min_value=50.0, max_value=1000.0,
                                      value=float(DC["document_setup"]["height_mm"]))
    else:
        width_mm, height_mm = PAGE_SIZES[size_choice]

    m1, m2 = st.columns(2)
    margin_tb = m1.number_input("Top/bottom margin (mm)", min_value=5.0, max_value=50.0,
                                 value=float(DC["document_setup"]["margins_mm"]["top"]))
    margin_lr = m2.number_input("Left/right margin (mm)", min_value=5.0, max_value=50.0,
                                 value=float(DC["document_setup"]["margins_mm"]["left"]))

    st.divider()

    # ---- Colors -------------------------------------------------------------
    st.markdown("### 🎨 Colors")
    primary = st.color_picker("Main heading", DC["color_palette"]["primary"])
    secondary = st.color_picker("Section heading / table header", DC["color_palette"]["secondary"])
    alt_row = st.color_picker("Table alternate-row", DC["components"]["tables"]["vocabulary_table"]["alternate_row_bg"])
    highlight_bg = st.color_picker("Highlight / objective box", DC["color_palette"]["accent_bg"])
    border_color = st.color_picker("Table & divider borders", DC["color_palette"]["border_color"])
    text_color = st.color_picker("Body text", DC["color_palette"]["dark_neutral"])

    st.divider()

    # ---- Fonts --------------------------------------------------------------
    st.markdown("### 🔤 Fonts")
    font_latin = st.text_input("English font", DC["typography"]["font_family"]["primary"])
    font_indic = st.text_input(
        "Kannada / Indic-script font", DC["typography"]["font_family"].get("indic", "Nirmala UI"),
        help="Used automatically for Kannada/Devanagari text -- the English "
             "font above usually has no Indic glyphs.",
    )

    st.divider()

    # ---- Cover page -----------------------------------------------------
    st.markdown("### 🖼️ Cover page")
    cover_enabled = st.checkbox("Add a designed cover page", value=True)
    title_text = st.text_input("Book / workbook title", "Workbook")
    subtitle_text = st.text_input("Subtitle (e.g. class/level)", "Class VII")
    fields_text = st.text_input("Blank fields on the cover (comma-separated)", "Name, Std, School")

    st.markdown("**Cover logos** (up to 3, shown left-to-right)")
    default_logos = [p for p in dtp.COVER["logos"] if p and os.path.isfile(p)]
    logo_cols = st.columns(max(len(default_logos), 1)) if default_logos else [st]
    for i, lp in enumerate(default_logos):
        with logo_cols[i]:
            st.image(lp, use_container_width=True)
    st.caption("current defaults shown above")
    logo_uploads = st.file_uploader(
        "Replace cover logos (upload 1-3 to override all of the above)",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="cover_logos_upload",
    )

    illustration_path = asset_picker(
        "Replace cover illustration", "cover_illustration_upload", dtp.COVER["illustration"],
    )

    st.divider()

    # ---- Footer -----------------------------------------------------------
    st.markdown("### 📌 Footer")
    footer_logo_path = asset_picker(
        "Replace footer logo", "footer_logo_upload", DC["footer"]["logo"],
    )
    page_number_corner = st.selectbox("Page number position", ["right", "left"], index=0)

    st.divider()

    # ---- Contents ---------------------------------------------------------
    st.markdown("### 📚 Contents page")
    add_contents = st.checkbox("Auto-generate a Contents page", value=True,
                                help='Inserts a real Word table-of-contents field. '
                                     'Open the result in Word and right-click it -> '
                                     '"Update Field" to fill in page numbers.')


# ============================================================================
# RUN
# ============================================================================
st.divider()
run_clicked = st.button("🚀 Format this document", type="primary",
                         disabled=uploaded_docx is None, use_container_width=True)

if run_clicked and uploaded_docx is not None:
    with st.spinner("Formatting document -- this takes a few seconds..."):
        input_path = save_uploaded(uploaded_docx, "input.docx")

        if logo_uploads:
            logo_paths = [
                save_uploaded(f, f"cover_logo_{i}{os.path.splitext(f.name)[1]}")
                for i, f in enumerate(logo_uploads[:3])
            ]
        else:
            logo_paths = default_logos

        overrides = {
            "document_setup": {
                "width_mm": width_mm, "height_mm": height_mm,
                "margins_mm": {"top": margin_tb, "bottom": margin_tb,
                                "left": margin_lr, "right": margin_lr},
            },
            "color_palette": {
                "primary": primary, "secondary": secondary,
                "accent_bg": highlight_bg, "light_neutral": highlight_bg,
                "border_color": border_color, "dark_neutral": text_color,
            },
            "typography": {"font_family": {"primary": font_latin, "indic": font_indic}},
            "components": {
                "tables": {
                    "vocabulary_table": {"header_bg": secondary, "alternate_row_bg": alt_row},
                    "grammar_table": {"header_bg": secondary},
                }
            },
            "footer": {"logo": footer_logo_path, "page_number_corner": page_number_corner},
        }
        cfg = dtp.deep_merge(dtp.DEFAULT_CONFIG, overrides)

        cover_cfg = {
            "enabled": cover_enabled,
            "logos": [p for p in logo_paths if p and os.path.isfile(p)],
            "illustration": illustration_path,
            "title": title_text,
            "subtitle": subtitle_text,
            "fields": [f.strip() for f in fields_text.split(",") if f.strip()],
        }

        output_docx = os.path.join(WORK_DIR, "formatted.docx")
        try:
            stats = dtp.build(
                input_path, output_docx, cfg, WORK_DIR,
                cover=cover_cfg, front_matter=[], add_contents=add_contents,
            )
            st.session_state.stats = stats
            st.session_state.output_docx = output_docx
            st.session_state.output_pdf = dtp.convert_to_pdf(output_docx)
            st.session_state.preview_png = None
            if st.session_state.output_pdf:
                try:
                    prefix = os.path.join(WORK_DIR, "preview")
                    subprocess.run(
                        ["pdftoppm", "-jpeg", "-r", "90", "-f", "1", "-l", "1",
                         st.session_state.output_pdf, prefix],
                        check=True, capture_output=True, timeout=60,
                    )
                    for fn in os.listdir(WORK_DIR):
                        if fn.startswith("preview") and fn.endswith(".jpg"):
                            st.session_state.preview_png = os.path.join(WORK_DIR, fn)
                except Exception:
                    pass
            st.session_state.error = None
        except Exception as e:
            st.session_state.error = str(e)

    if st.session_state.get("error"):
        st.error(f"Something went wrong: {st.session_state['error']}")
    else:
        st.success("Done! Your formatted document is ready below.")

# ============================================================================
# RESULTS
# ============================================================================
if st.session_state.get("output_docx") and os.path.exists(st.session_state["output_docx"]):
    st.subheader("Result")

    res_col, prev_col = st.columns([2, 1])

    with res_col:
        stats = st.session_state.get("stats", {})
        m = st.columns(4)
        m[0].metric("Section headings", stats.get("heading", 0))
        m[1].metric("Tables formatted", sum(
            stats.get(k, 0) for k in
            ("table_vocab", "table_dialogue", "table_wordgrid", "table_blank", "table_generic")
        ))
        m[2].metric("Images placed", stats.get("images", 0))
        m[3].metric("Titles", stats.get("title", 0))

        dcol, pcol = st.columns(2)
        with open(st.session_state["output_docx"], "rb") as f:
            dcol.download_button(
                "⬇️ Download DOCX", f, file_name="formatted_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        pdf_path = st.session_state.get("output_pdf")
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                pcol.download_button(
                    "⬇️ Download PDF", f, file_name="formatted_document.pdf",
                    mime="application/pdf", use_container_width=True,
                )
        else:
            pcol.info("PDF export needs LibreOffice on this server. The DOCX always works.")

        with st.expander("Full processing summary"):
            st.json(stats)

    with prev_col:
        if st.session_state.get("preview_png") and os.path.exists(st.session_state["preview_png"]):
            st.image(st.session_state["preview_png"], caption="Page 1 preview", use_container_width=True)
