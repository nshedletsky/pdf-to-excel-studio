"""PDF extraction and Excel workbook generation."""
from __future__ import annotations
import io
import re
from collections import Counter
from pathlib import Path
from typing import Any

import fitz
import pandas as pd
import pdfplumber
import xlsxwriter

COURSE_RE = re.compile(r"\b([A-Z]{2,5})\s*[- ]?\s*(\d{3}[A-Z]?)\b")
NUMERIC_RE = re.compile(r"^\(?-?\$?\d[\d,]*(?:\.\d+)?%?\)?$")
EXCLUDE_TOTAL = re.compile(r"\b(total|subtotal|hours|credit hours|credits required)\b", re.I)


def safe_sheet_name(name: str, used: set[str]) -> str:
    root = re.sub(r"[\[\]:*?/\\]", " ", Path(name).stem).strip().strip("'") or "Sheet"
    root = root[:31]
    candidate = root
    k = 2
    while candidate.casefold() in used:
        suffix = f" ({k})"
        candidate = root[:31-len(suffix)] + suffix
        k += 1
    used.add(candidate.casefold())
    return candidate


def _lines_in_column(words: list[tuple], x_min: float, x_max: float) -> list[tuple[float, float, str]]:
    """Group by vertical location, not PDF text blocks (table columns are often separate blocks)."""
    picked = [w for w in words if x_min <= (w[0]+w[2])/2 < x_max]
    picked.sort(key=lambda w: (w[1],w[0]))
    bands: list[list] = []
    for w in picked:
        if bands and abs(w[1]-bands[-1][0][1]) <= 3.0:
            bands[-1].append(w)
        else:
            bands.append([w])
    output=[]
    for arr in bands:
        arr.sort(key=lambda w:w[0])
        output.append((min(w[1] for w in arr),min(w[0] for w in arr)," ".join(w[4] for w in arr).strip()))
    return sorted(output,key=lambda t:(t[0],t[1]))


def _course_rows(pdf_bytes: bytes) -> list[dict[str, Any]]:
    """Extract course entries by their visual columns, never from trailing numbers."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    result: list[dict[str, Any]] = []
    for page_idx, page in enumerate(doc, 1):
        words = page.get_text("words")
        mid = page.rect.width / 2
        for side, lo, hi in (("Left", 0, mid-5), ("Right", mid-5, page.rect.width+1)):
            lines = _lines_in_column(words, lo, hi)
            section = "Uncategorized"
            for y, x, line in lines:
                m = re.match(r"^\s*(?:[_•☐✓]+\s*)?([A-Z]{2,5})\s*[- ]?\s*(\d{3}[A-Z]?)(?:[\*†]*)\b", line)
                if not m:
                    if line and len(line) < 92 and not re.search(r"student (name|id)|phone|email|mailing|entry date", line, re.I):
                        if re.search(r"(\b(?:major|core|concentration|electives|social sciences|humanities|religious studies|communication|natural sciences|prerequisites|term [1-9])\b)",line,re.I):
                            section = line
                    continue
                code=f"{m.group(1)} {m.group(2)}"
                row_words=[w for w in words if lo <= (w[0]+w[2])/2 < hi and abs(w[1]-y)<=3.0]
                row_words.sort(key=lambda w:w[0])
                start_x=next((w[0] for w in row_words if w[4]==m.group(1)), x)
                if side=="Left":
                    credit_band=(hi-33,hi-4)
                elif start_x >= mid+30:
                    credit_band=(page.rect.width-52,page.rect.width-25)
                else:
                    credit_band=(mid+165,mid+187)
                candidate=[]
                for w in row_words:
                    center=(w[0]+w[2])/2
                    t=w[4].strip()
                    if credit_band[0] <= center <= credit_band[1] and re.fullmatch(r"(?:[1-9]|1[0-2])(?:\.5)?", t):
                        candidate.append((center,float(t)))
                credit=candidate[-1][1] if len(candidate)==1 else None
                code_end=next((w[2] for w in row_words if re.match(r"^"+re.escape(m.group(2))+r"[\*†]*$",w[4])),start_x+45)
                title_words=[w[4] for w in row_words if w[0]>=code_end-1 and (w[0]+w[2])/2 < credit_band[0]-3]
                title=" ".join(title_words).strip(" -*:†")
                if not title or len(title)>160:
                    continue
                clinical=None
                if side=="Right":
                    far=[(w[0],w[4]) for w in row_words if w[0]>=page.rect.width-58 and re.fullmatch(r"\d+(?:\.\d+)?",w[4])]
                    if section.lower().startswith("term ") and far:
                        clinical=float(far[-1][1])
                result.append({"Page":page_idx,"Section":section,"Course":code,"Title":title,"Credits":credit,"Clinical Hours":clinical,"Source side":side})
    return result


def _tables(pdf_bytes: bytes) -> list[dict[str, Any]]:
    output = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as doc:
        for page_idx, page in enumerate(doc.pages, 1):
            for table in page.extract_tables() or []:
                if not table or len(table) < 2:
                    continue
                width = max(map(len, table))
                for row in table:
                    vals = [(str(v).strip() if v is not None else "") for v in (row + [None]*(width-len(row)))]
                    if any(vals):
                        output.append({"Page":page_idx, **{f"Column {i+1}": v for i,v in enumerate(vals)}})
    return output


def _text_rows(pdf_bytes: bytes) -> list[dict[str, Any]]:
    doc = fitz.open(stream=pdf_bytes,filetype="pdf")
    return [{"Page":i,"Text":line.strip()} for i,p in enumerate(doc,1) for line in p.get_text(sort=True).splitlines() if line.strip()]


def extract_pdf(pdf_bytes: bytes, filename: str) -> tuple[pd.DataFrame, dict]:
    """Return editable records and diagnosis. Scanned PDFs are reported, never invented."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = len(doc)
    if pages > 100:
        raise ValueError("PDF exceeds the 100-page prototype limit.")
    if not any(p.get_text().strip() for p in doc):
        raise ValueError("This PDF appears image-only. OCR is not included in this first prototype.")
    courses = _course_rows(pdf_bytes)
    # Prefer specialized course extraction only when this actually looks like a program sheet.
    if len(courses) >= 3:
        df = pd.DataFrame(courses).drop(columns=["Source side"])
        kind = "Program sheet"
    else:
        tables = _tables(pdf_bytes)
        if tables:
            df = pd.DataFrame(tables).fillna("")
            kind = "Tables"
        else:
            df = pd.DataFrame(_text_rows(pdf_bytes))
            kind = "Text lines"
    return df, {"filename":filename,"pages":pages,"mode":kind,"rows":len(df),"warning":("Review the extracted rows. Missing or duplicate courses suppress automatic totals. This is not a degree audit." if kind == "Program sheet" else "Check extraction against the PDF; layout may not be fully preserved.")}


def _as_number(value: Any) -> float | None:
    if isinstance(value, bool) or pd.isna(value):
        return None
    if isinstance(value, (int,float)):
        return float(value)
    s = str(value).strip()
    if not NUMERIC_RE.fullmatch(s) or s.endswith("%"):
        return None
    try:
        return float(s.replace("$","").replace(",", "").replace("(","-").replace(",", "").replace(")",""))
    except ValueError:
        return None


def infer_additive_columns(frame: pd.DataFrame, mode: str) -> list[str]:
    """Only total complete, unique credit-hour columns in program sheets."""
    if mode != "Program sheet" or "Credits" not in frame or "Course" not in frame:
        return []
    credits = [_as_number(v) for v in frame["Credits"]]
    if any(v is None for v in credits):
        return []
    if frame["Course"].astype(str).duplicated().any():
        return []
    return ["Credits"] if len(credits) >= 2 else []


def build_workbook(items: list[tuple[str,pd.DataFrame,str]], add_totals: bool=True) -> bytes:
    """One sheet per file; totals are actual Excel formulas, not static values."""
    output = io.BytesIO()
    used: set[str] = set()
    wb = xlsxwriter.Workbook(output, {"in_memory": True, "strings_to_formulas": False, "strings_to_urls": False})
    banner = wb.add_format({"bold":True,"font_color":"white","bg_color":"#183153","font_size":15,"valign":"vcenter"})
    header = wb.add_format({"bold":True,"font_color":"white","bg_color":"#2463A0","border":0,"text_wrap":True})
    cell = wb.add_format({"valign":"top","text_wrap":True,"bottom":1,"bottom_color":"#E6EAF0"})
    number = wb.add_format({"num_format":"0.##","valign":"top","bottom":1,"bottom_color":"#E6EAF0"})
    total = wb.add_format({"bold":True,"top":2,"top_color":"#2463A0","num_format":"0.##","bg_color":"#EAF3FA"})
    for fname, data, mode in items:
        if data is None or data.empty:
            continue
        sheet = wb.add_worksheet(safe_sheet_name(fname, used))
        sheet.hide_gridlines(2)
        cols = [str(c) for c in data.columns]
        n = len(cols)
        sheet.merge_range(0, 0, 0, max(n-1,1), Path(fname).name, banner)
        sheet.set_row(0, 28)
        for j, name in enumerate(cols):
            values = data.iloc[:,j].astype(str).fillna("")
            width = min(55, max(13, min(50,len(name)+2), int(values.str.len().quantile(.85)) + 3))
            sheet.set_column(j,j,width)
            sheet.write(2,j,name,header)
        for i, row in enumerate(data.itertuples(index=False,name=None),3):
            for j,v in enumerate(row):
                if pd.isna(v) or v=="":
                    sheet.write_blank(i,j,None,cell)
                elif isinstance(v,(int,float)) and not isinstance(v,bool):
                    sheet.write_number(i,j,float(v),number)
                else:
                    parsed = _as_number(v) if cols[j]=="Credits" else None
                    if parsed is not None:
                        sheet.write_number(i,j,parsed,number)
                    else:
                        # xlsxwriter string formatting prevents formulas from untrusted PDFs executing.
                        sheet.write_string(i,j,str(v)[:32767],cell)
            sheet.set_row(i, 30)
        last = 2+len(data)
        sheet.autofilter(2,0,last,n-1)
        sheet.freeze_panes(3,0)
        if add_totals:
            for c in infer_additive_columns(data,mode):
                j=cols.index(c)
                sheet.write(last+2, max(0,j-1),"Sum of extracted credits (not degree total)",total)
                letter=xlsxwriter.utility.xl_col_to_name(j)
                sheet.write_formula(last+2,j,f"=SUM({letter}4:{letter}{last+1})",total)
        sheet.set_landscape()
        sheet.fit_to_pages(1,0)
    wb.close()
    return output.getvalue()
