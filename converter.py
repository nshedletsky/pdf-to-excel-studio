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
    """Identify courses visually, using credit columns rather than trailing text numbers."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    result = []
    for page_no, page in enumerate(doc, 1):
        words = page.get_text("words")
        middle = page.rect.width / 2
        for side, low, high in (("Left", 0, middle), ("Right", middle, page.rect.width + 1)):
            section = "General"
            for y, x, line in _lines_in_column(words, low, high):
                # Course codes can be preceded by completion checkboxes/underscores.
                match = COURSE_RE.search(line)
                if not match:
                    if len(line) < 90 and re.search(r"\b(major|core|finance|concentration|electives|studies|communications?|humanities|social sciences|mathematics|prerequisites|term [1-9])\b",line,re.I):
                        section = line
                    continue
                # Avoid references to a course embedded in a paragraph/footnote.
                before = line[:match.start()].strip()
                if len(before) > 28 or any(c.isalpha() for c in before.replace("_", "")):
                    continue
                course = f"{match.group(1)} {match.group(2)}"
                row = sorted((w for w in words if low <= (w[0]+w[2])/2 < high and abs(w[1]-y) <= 3), key=lambda w:w[0])
                code_word = next((w for w in row if re.fullmatch(re.escape(match.group(2))+r"[\*†]*",w[4]) or re.fullmatch(re.escape(match.group(1))+re.escape(match.group(2))+r"[\*†]*",w[4])),None)
                if not code_word:
                    continue
                code_end = code_word[2]
                code_start = next((w[0] for w in row if w[4] == match.group(1)), code_word[0])
                if side == "Left":
                    credit_lo, credit_hi = middle-48, middle-6
                elif code_start < middle+35:
                    credit_lo, credit_hi = middle+145, middle+200
                else:
                    credit_lo, credit_hi = page.rect.width-65, page.rect.width-25
                matches = [w for w in row if credit_lo <= (w[0]+w[2])/2 <= credit_hi and re.fullmatch(r"(?:[1-9]|1[0-2])(?:\.5)?",w[4])]
                credits = float(matches[-1][4]) if len(matches)==1 else None
                title_parts = [w[4] for w in row if w[0] >= code_end-1 and (w[0]+w[2])/2 < credit_lo-3]
                title = " ".join(title_parts).strip(" -*:†_")
                # Some PDFs attach academic credit values to the final title word.
                # Never interpret clinical/contact-hour figures as course credits.
                if credits is None and not re.search(r"clinical|contact|lab hours", line, re.I):
                    glued = re.search(r"(?<=[A-Za-z])([1-9])$", title)
                    if glued:
                        credits = float(glued.group(1))
                        title = title[:-1].rstrip()
                if not title or len(title)>160:
                    continue
                clinical = None
                if side == "Right" and code_start < middle+35:
                    hours = [float(w[4]) for w in row if (w[0]+w[2])/2 > middle+230 and re.fullmatch(r"\d+(?:\.\d+)?",w[4])]
                    if hours:
                        clinical=hours[-1]
                result.append({"Page":page_no,"Section":section,"Course":course,"Title":title,"Credits":credits,"Clinical Hours":clinical,"Source side":side})
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


# Layout-preserving curriculum worksheet export. It complements the flat table export.
LAYOUT_CODE_RE = re.compile(r"\b([A-Z]{2,5})\s*(\d{3}[A-Z]?)\b")
SECTION_HINT_RE = re.compile(r"(\(\s*\d+\s*(?:credit\s*)?hours?\s*\)|university studies|business major|general electives|support courses|concentration|humanities|communication|social sciences|natural sciences|religious studies|major courses|prerequisites|term \d+)", re.I)

def layout_rows(pdf_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Read each page as two independent visual columns; retain blank elective lines."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    halves = [[], []]
    for pageno, page in enumerate(doc, 1):
        midpoint = page.rect.width / 2
        for index, (lo, hi) in enumerate(((0, midpoint), (midpoint, page.rect.width+1))):
            lines = _lines_in_column(page.get_text("words"), lo, hi)
            started = False
            for y, x, label in lines:
                # Student identity fields belong in the header, not curriculum requirements.
                if re.search(r"^(student (name|id)|entry date|phone number|email address|mailing address)", label, re.I):
                    continue
                if not label.strip() or re.fullmatch(r"[_\s]+", label):
                    continue
                text = re.sub(r"^\s*(?:(?:_{2,}\s*)+|[☐□]+\s*)", "", label).strip()
                # The title block isn't curriculum content. Keep it in the
                # workbook title, not as zero-credit sections on the left.
                if not started:
                    if index == 0 and re.search(r"^UNIVERSITY STUDIES PROGRAM",text,re.I):
                        started = True
                    elif index == 1 and re.search(r"^(?:Business Major|Major Courses|Required Prerequisites|Upper Division|Core Courses|Program Requirements)",text,re.I):
                        started = True
                    else:
                        continue
                m = LAYOUT_CODE_RE.search(text)
                placeholder = bool(re.match(r"^(?:_{2,}|[A-Z]{2,5}_{2,})(?:\s|$)", text))
                number = None
                code = ""
                kind = "Note"
                title = text
                if re.fullmatch(r"[1-9]|1[0-2]",text):
                    # Standalone amounts represent unnamed requirements/electives,
                    # not notes. These must remain editable.
                    kind = "Elective / blank"
                    title = "Elective / unspecified requirement"
                    number = int(text)
                elif m and m.start() < 22:
                    kind = "Course"
                    code = f"{m.group(1)} {m.group(2)}"
                    title = text[m.end():].strip(" _*†:-")
                elif placeholder or (re.search(r"_{3,}", text) and len(text)<95):
                    kind = "Elective / blank"
                    title = text
                elif SECTION_HINT_RE.search(text) and len(text)<95 and not re.search(r"recommended:|completed|must have|requirement to",text,re.I):
                    kind = "Section"
                # Only parse a single credit figure at the very end, and never
                # mistake clinical, contact, or total-hour values for course credits.
                if kind in ("Course", "Elective / blank") and number is None:
                    credit = re.search(r"\s+([1-9]|1[0-2])\s*$",title)
                    glued = re.search(r"(?<=[A-Za-z])([1-9])$",title)
                    if credit:
                        number=int(credit.group(1))
                        title=title[:credit.start()].strip()
                    elif glued and not re.search(r"clinical|contact|hours", title, re.I):
                        number=int(glued.group(1))
                        title=title[:-1].strip()
                    if re.search(r"clinical|contact|total hours", title, re.I):
                        number=None
                # Do not discard notes, since programs need to be able to edit them.
                halves[index].append({"Type":kind,"Code":code,"Course / Requirement":title,"Credits":number,"Page":pageno})
    return tuple(pd.DataFrame(items,columns=["Type","Code","Course / Requirement","Credits","Page"]) for items in halves)


def build_layout_workbook(items: list[tuple[str, pd.DataFrame, pd.DataFrame]]) -> bytes:
    """Print-ready, two-sided editable worksheets with independently calculated section totals."""
    buf=io.BytesIO()
    book=xlsxwriter.Workbook(buf,{"in_memory":True,"strings_to_formulas":False,"strings_to_urls":False})
    used=set()
    navy=book.add_format({"bold":True,"font_color":"white","bg_color":"#17324D","font_size":14,"align":"center","valign":"vcenter"})
    banner=book.add_format({"bold":True,"font_color":"white","bg_color":"#17324D","valign":"vcenter"})
    section=book.add_format({"bold":True,"bg_color":"#DDEBF5","font_color":"#17324D","valign":"vcenter","text_wrap":True})
    normal=book.add_format({"valign":"top","text_wrap":True})
    striped=book.add_format({"valign":"top","text_wrap":True,"bg_color":"#F3F6FA"})
    credit=book.add_format({"num_format":"0","bg_color":"#E9F5EC","valign":"top","align":"center"})
    subtotal=book.add_format({"bold":True,"bg_color":"#EEF2F7","font_color":"#17324D"})
    total=book.add_format({"bold":True,"bg_color":"#17324D","font_color":"white"})
    caution=book.add_format({"bg_color":"#FFF3D8","font_color":"#76520F","text_wrap":True,"valign":"vcenter"})
    for filename,left,right in items:
        if left.empty and right.empty:
            continue
        ws=book.add_worksheet(safe_sheet_name(filename,used))
        ws.hide_gridlines(2)
        ws.set_column("A:A",3);ws.set_column("B:B",10);ws.set_column("C:C",35)
        ws.set_column("D:D",7);ws.set_column("E:E",2)
        ws.set_column("F:F",3);ws.set_column("G:G",10);ws.set_column("H:H",35);ws.set_column("I:I",7)
        ws.merge_range("A1:I2",Path(filename).stem.replace("_"," ")+" | Editable Program Sheet",navy)
        ws.set_row(0,22);ws.set_row(1,22)
        ws.merge_range("A4:D4","UNIVERSITY STUDIES",banner)
        ws.merge_range("F4:I4","MAJOR, SUPPORT & ELECTIVES",banner)
        summaries=[]
        ends=[]
        side_values=[]
        for data, cols in ((left,("A","B","C","D")),(right,("F","G","H","I"))):
            flag,code,title,hours=cols
            row=5
            credit_cells=[]
            credit_values=[]
            groups=[]
            group_values=[]
            for idx,(_, item) in enumerate(data.reset_index(drop=True).iterrows()):
                kind=str(item["Type"])
                label=str(item["Course / Requirement"] if pd.notna(item["Course / Requirement"]) else "")
                number=item["Credits"]
                if kind=="Section":
                    if credit_cells:
                        ws.write(f"{title}{row}","Section subtotal",subtotal)
                        ws.write_formula(f"{hours}{row}","=SUM("+",".join(credit_cells)+")",subtotal,sum(credit_values))
                        groups.append(row)
                        group_values.append(sum(credit_values))
                        row+=2
                    credit_cells=[]
                    credit_values=[]
                    ws.merge_range(f"{flag}{row}:{hours}{row}",label,section)
                    ws.set_row(row-1,22)
                    row+=1
                    continue
                usefmt=striped if idx%2 else normal
                ws.write_string(f"{flag}{row}","☐" if kind in ("Course","Elective / blank") else "",usefmt)
                ws.write_string(f"{code}{row}",str(item["Code"] or ""),usefmt)
                ws.write_string(f"{title}{row}",label,usefmt)
                if number is not None and not pd.isna(number):
                    ws.write_number(f"{hours}{row}",float(number),credit)
                    credit_cells.append(f"{hours}{row}")
                    credit_values.append(float(number))
                else:
                    ws.write_blank(f"{hours}{row}",None,credit)
                ws.set_row(row-1,20 if len(label)<55 else 30)
                row+=1
            if credit_cells:
                subtotal_row = row
                ws.write(f"{title}{subtotal_row}","Section subtotal",subtotal)
                ws.write_formula(f"{hours}{subtotal_row}","=SUM("+",".join(credit_cells)+")",subtotal,sum(credit_values))
                groups.append(subtotal_row)
                group_values.append(sum(credit_values))
                row+=2
            ends.append(row)
            summaries.append((cols,groups))
            side_values.append(sum(group_values))
        bottom=max(ends)+1
        for ((cols, groups), cached), caption in zip(zip(summaries,side_values),("LEFT SIDE TOTAL","RIGHT SIDE TOTAL")):
            flag,code,title,hours=cols
            ws.merge_range(f"{flag}{bottom}:{title}{bottom}",caption,total)
            formula="=SUM("+",".join(f"{hours}{r}" for r in groups)+")" if groups else "=0"
            ws.write_formula(f"{hours}{bottom}",formula,total,cached)
        ws.merge_range(f"F{bottom+2}:H{bottom+2}","COMBINED EXTRACTED CREDITS",section)
        ws.write_formula(f"I{bottom+2}",f"=D{bottom}+I{bottom}",subtotal,sum(side_values))
        ws.merge_range(f"A{bottom+4}:I{bottom+5}","REVIEW BEFORE APPROVAL: Blank requirements, reused courses, OCR/layout ambiguities and electives may affect totals. Compare against the source PDF. These totals are not an official degree audit.",caution)
        ws.freeze_panes(4,0)
        ws.set_landscape();ws.set_paper(5);ws.fit_to_pages(1,0)
        ws.set_margins(.2,.2,.3,.3)
        ws.print_area(0,0,bottom+4,8)
        ws.center_horizontally()
        ws.set_footer("&CPage &P of &N")
    book.close()
    return buf.getvalue()
