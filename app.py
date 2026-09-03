"""
app.py - Streamlit front-end for dtp_pipeline.py

A point-and-click UI so a non-developer can:
  1. Upload a raw/unedited .docx
  2. Optionally tweak colors, fonts, cover page, footer, contents page
  3. Click Run
  4. Download the formatted result as .docx and/or .pdf

Run it with:
    pip install streamlit python-docx Pillow
    streamlit run app.py

Keep this file, dtp_pipeline.py, and the assets/ folder together in the
same directory -- the app imports dtp_pipeline.py directly and reuses the
same bundled logos/illustration as sensible defaults.
"""

import os
import tempfile
import shutil

import streamlit as st

import dtp_pipeline as dtp


st.set_page_config(page_title="Workbook DTP Formatter", page_icon="📘", layout="wide")


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


# ----------------------------------------------------------------------------
# Header
# ----------------------------------------------------------------------------
st.title("📘 Workbook DTP Formatter")
st.caption(
    "Upload a raw, unedited Word document. The tool applies consistent "
    "styling, tables, a cover page, and page numbers automatically -- no "
    "coding needed. Open **Design settings** below only if you want to "
    "change colors, fonts, logos, or the cover page."
)

uploaded_docx = st.file_uploader("1. Upload the raw .docx", type=["docx"])

# ----------------------------------------------------------------------------
# Design settings (all optional -- sensible defaults are pre-filled)
# ----------------------------------------------------------------------------
DC = dtp.DEFAULT_CONFIG

with st.expander("🎨 Design settings -- colors & fonts", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Colors**")
        primary = st.color_picker("Main heading color", DC["color_palette"]["primary"])
        secondary = st.color_picker("Section heading / table header color", DC["color_palette"]["secondary"])
        alt_row = st.color_picker(
            "Vocabulary table alternate-row color",
            DC["components"]["tables"]["vocabulary_table"]["alternate_row_bg"],
        )
        highlight_bg = st.color_picker("Highlight / objective box background", DC["color_palette"]["accent_bg"])
        border_color = st.color_picker("Table & divider border color", DC["color_palette"]["border_color"])
        text_color = st.color_picker("Body text color", DC["color_palette"]["dark_neutral"])
    with c2:
        st.markdown("**Fonts**")
        font_latin = st.text_input("English font", DC["typography"]["font_family"]["primary"])
        font_indic = st.text_input(
            "Kannada / Indic-script font",
            DC["typography"]["font_family"].get("indic", "Nirmala UI"),
            help="Used automatically for any Kannada/Devanagari text, since the "
                 "English font above usually has no Indic glyphs.",
        )
        st.markdown("**Page**")
        page_size = st.selectbox("Page size", ["A4", "US Letter"], index=0)

with st.expander("📄 Cover page", expanded=False):
    cover_enabled = st.checkbox("Add a designed cover page", value=True)
    title_text = st.text_input("Book / workbook title", "Workbook")
    subtitle_text = st.text_input("Subtitle (e.g. class/level)", "Class VII")
    fields_text = st.text_input(
        "Blank fields printed on the cover (comma-separated)", "Name, Std, School"
    )
    st.caption("Leave logo/illustration uploads empty to use the bundled defaults.")
    logo_uploads = st.file_uploader(
        "Cover logos (up to 3 images, shown left-to-right)",
        type=["png", "jpg", "jpeg"], accept_multiple_files=True,
    )
    illustration_upload = st.file_uploader(
        "Cover illustration (optional, large centered graphic)",
        type=["png", "jpg", "jpeg"],
    )

with st.expander("📑 Footer & Contents page", expanded=False):
    st.caption("Leave the footer logo empty to use the bundled default logo.")
    footer_logo_upload = st.file_uploader("Footer logo (centered in the footer)", type=["png", "jpg", "jpeg"])
    page_number_corner = st.selectbox("Page number position", ["right", "left"], index=0)
    add_contents = st.checkbox("Auto-generate a Contents page", value=True)

st.divider()
run_clicked = st.button("🚀 Format this document", type="primary", disabled=uploaded_docx is None)


# ----------------------------------------------------------------------------
# Run the pipeline
# ----------------------------------------------------------------------------
if run_clicked and uploaded_docx is not None:
    with st.spinner("Formatting document -- this takes a few seconds..."):
        input_path = save_uploaded(uploaded_docx, "input.docx")

        # Cover logos: use uploads if provided, else fall back to bundled assets
        if logo_uploads:
            logo_paths = [
                save_uploaded(f, f"cover_logo_{i}{os.path.splitext(f.name)[1]}")
                for i, f in enumerate(logo_uploads[:3])
            ]
        else:
            logo_paths = dtp.COVER["logos"]

        illustration_path = (
            save_uploaded(illustration_upload, f"cover_illustration{os.path.splitext(illustration_upload.name)[1]}")
            if illustration_upload else dtp.COVER["illustration"]
        )
        footer_logo_path = (
            save_uploaded(footer_logo_upload, f"footer_logo{os.path.splitext(footer_logo_upload.name)[1]}")
            if footer_logo_upload else DC["footer"]["logo"]
        )

        overrides = {
            "color_palette": {
                "primary": primary,
                "secondary": secondary,
                "accent_bg": highlight_bg,
                "light_neutral": highlight_bg,
                "border_color": border_color,
                "dark_neutral": text_color,
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
        if page_size == "US Letter":
            overrides["document_setup"] = {"width_mm": 216, "height_mm": 279}

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
            st.session_state.error = None
        except Exception as e:
            st.session_state.error = str(e)

    if st.session_state.get("error"):
        st.error(f"Something went wrong: {st.session_state['error']}")
    else:
        st.success("Done! Your formatted document is ready below.")


# ----------------------------------------------------------------------------
# Results + downloads (persist across reruns within the session)
# ----------------------------------------------------------------------------
if st.session_state.get("output_docx") and os.path.exists(st.session_state["output_docx"]):
    st.subheader("Result")

    stats = st.session_state.get("stats", {})
    cols = st.columns(4)
    cols[0].metric("Section headings", stats.get("heading", 0))
    cols[1].metric("Tables formatted", sum(
        stats.get(k, 0) for k in
        ("table_vocab", "table_dialogue", "table_wordgrid", "table_blank", "table_generic")
    ))
    cols[2].metric("Images placed", stats.get("images", 0))
    cols[3].metric("Titles", stats.get("title", 0))

    dcol, pcol = st.columns(2)
    with open(st.session_state["output_docx"], "rb") as f:
        dcol.download_button(
            "⬇️ Download formatted DOCX", f, file_name="formatted_document.docx",
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
        pcol.info("PDF export needs LibreOffice on this server. The DOCX above always works and can be exported to PDF from Word.")

    with st.expander("Full processing summary"):
        st.json(stats)
