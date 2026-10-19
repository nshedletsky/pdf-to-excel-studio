# PDF to Excel Studio

Convert multiple selectable-text PDFs into a single editable Excel workbook, with one worksheet per PDF.

## Features

- Upload up to 15 PDFs, each up to 15 MB and 100 pages.
- Review and edit extracted rows before downloading.
- Extract course codes, titles, and credits from many academic program sheets.
- Fall back to tables or text lines for other PDFs.
- Export formatted Excel worksheets with filters and frozen headers.
- Add automatic formulas to recognized credit-hour columns.

## Deploy to Streamlit Community Cloud

1. Visit https://share.streamlit.io and sign in with GitHub.
2. Select **Create app**, then **Deploy a public app from GitHub**.
3. Choose this repository, branch `main`, and main file `app.py`.
4. Select **Deploy**. Streamlit installs packages from `requirements.txt`.

## Local development (optional)

With Python 3.11 or newer installed:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Tests: `python -m unittest discover -s tests -v`.

## Limitations

- Scanned PDFs require OCR, which is not included in this version.
- Complex, multi-column layouts can cause errors. Always compare the result with the original.
- Automatic totals apply conservatively to credit-hour columns and do not verify published degree totals.
- On a hosted deployment, files are processed on the app server. Do not upload confidential records without an appropriate privacy and security review.
