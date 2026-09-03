# Workbook DTP Formatter

Turns a raw, unedited Word workbook into a professionally formatted DOCX
(cover page, styled headings/tables, footer, auto Contents page) -- either
from the command line or from a point-and-click web page.

## Folder contents

```
dtp_pipeline.py    the formatting engine (also runnable on its own)
app.py             Streamlit UI on top of dtp_pipeline.py
requirements.txt   what to pip install
assets/            bundled logos + cover illustration (defaults)
```

Keep all of these together in one folder -- the app looks for `assets/`
next to itself.

## One-time setup

```bash
pip install -r requirements.txt
```

(Optional, for PDF export) install LibreOffice and make sure the `soffice`
command is on your PATH. Without it, DOCX export still works fine -- you can
open the DOCX in Word and use "Save As PDF" instead.

## Option A -- non-developers: the web UI

```bash
streamlit run app.py
```

This opens a page in your browser:

1. Upload the raw `.docx`.
2. (Optional) open **Design settings** / **Cover page** / **Footer & Contents**
   to change colors, fonts, logos, or the cover text -- no coding needed.
3. Click **Format this document**.
4. Download the result as **DOCX** or **PDF**.

## Option B -- command line

Edit the three variables at the top of `dtp_pipeline.py`:

```python
INPUT_DOCX = "your_raw_file.docx"
OUTPUT_DOCX = None   # None = auto-name "<input>_formatted.docx"
STYLE_JSON = None    # None = use the built-in design system
```

then just run:

```bash
python3 dtp_pipeline.py
```

Or, without editing anything, pass it on the command line instead:

```bash
python3 dtp_pipeline.py your_raw_file.docx -o output.docx --pdf
```

## Customizing the design permanently

Anything changed in the Streamlit UI only affects that one run. To change
the *defaults* everyone gets (colors, fonts, logos, footer), either:

- edit the `DEFAULT_CONFIG` / `COVER` dictionaries near the top of
  `dtp_pipeline.py`, or
- pass your own JSON file with `-s your_style.json` (command line) -- you
  only need to include the keys you want to override, everything else
  falls back to the built-in defaults.
