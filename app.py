"""Run: streamlit run app.py"""
from __future__ import annotations
import hashlib
from pathlib import Path
import pandas as pd
import streamlit as st
from converter import extract_pdf, build_workbook, layout_rows, build_layout_workbook, graduate_rows, build_graduate_workbook

st.set_page_config(page_title="PDF to Excel for Jane!",page_icon="📊",layout="wide")
st.title("PDF to Excel for Jane")
st.write("Upload multiple PDFs and get one editable Excel workbook, with one worksheet per PDF.")
st.caption("Prototype: works best with selectable-text PDFs. Scanned/image-only PDFs are not supported yet.")

@st.cache_data
def load_course_catalog():
    catalog=pd.read_csv(Path(__file__).with_name("course_catalog.csv")).fillna("")
    catalog["Display"]=catalog["Course Code"]+" — "+catalog["Course Title"]
    return catalog

CATALOG=load_course_catalog()
COURSE_CHOICES=["(Type a new/proposed course)"]+CATALOG["Display"].tolist()
FIELDS=["Type","Code","Course / Requirement","Credits","Page"]

def blank_rows():
    return pd.DataFrame(columns=FIELDS)

def edit_requirements(data, identifier, label, page=1):
    """The displayed editor and the add-course controls share the same row data."""
    data_key=f"rows_{identifier}"
    revision_key=f"rev_{identifier}"
    if data_key not in st.session_state:
        st.session_state[data_key]=data.copy()
    if revision_key not in st.session_state:
        st.session_state[revision_key]=0
    st.markdown(f"**{label}**")
    frame=st.data_editor(
        st.session_state[data_key], key=f"edit_{identifier}_{st.session_state[revision_key]}",
        num_rows="dynamic",hide_index=True,use_container_width=True,
        column_config={"Type":st.column_config.SelectboxColumn("Type",options=["Section","Course","Elective / blank","Note"])},
    )
    st.session_state[data_key]=frame.copy()
    st.caption("Use the searchable course dropdown below to add courses directly to this table. You can also type in any cell.")
    with st.container(border=True):
        st.markdown("**Add courses or requirements**")
        kind=st.selectbox("What would you like to add?",["Course from catalog","Proposed / manual course","Section heading","Blank elective / requirement","Note"],key=f"kind_{identifier}")
        course=""
        if kind=="Course from catalog":
            course=st.selectbox("Select a course (type its number or name to search)",COURSE_CHOICES,key=f"catalog_{identifier}")
        manual_code=""
        title=""
        if kind=="Proposed / manual course":
            manual_code=st.text_input("Course code (optional)",key=f"code_{identifier}")
            title=st.text_input("Course title",key=f"title_{identifier}")
        if kind in ("Section heading","Blank elective / requirement","Note"):
            title=st.text_input("Section name or requirement",key=f"title_{identifier}")
        credits=None
        if kind in ("Course from catalog","Proposed / manual course","Blank elective / requirement"):
            entered=st.number_input("Credit hours (confirm from program catalog)",min_value=0.0,max_value=50.0,value=3.0,step=1.0,key=f"credits_{identifier}")
            credits=float(entered)
        position=st.selectbox("Where should this row go?",["At the end","Before an existing row"],key=f"position_{identifier}")
        insert_at=len(frame)
        if position=="Before an existing row" and len(frame):
            insert_at=st.number_input("Insert before row number",min_value=1,max_value=len(frame),value=len(frame),step=1,key=f"index_{identifier}")-1
        if st.button("Add selected item to table",key=f"add_{identifier}",type="primary"):
            if kind=="Course from catalog":
                if course==COURSE_CHOICES[0]:
                    st.warning("Choose a course from the list first.")
                    return frame
                selected=CATALOG.iloc[COURSE_CHOICES.index(course)-1]
                new={"Type":"Course","Code":str(selected["Course Code"]).replace("-"," "),"Course / Requirement":str(selected["Course Title"]),"Credits":credits,"Page":page}
            else:
                types={"Proposed / manual course":"Course","Section heading":"Section","Blank elective / requirement":"Elective / blank","Note":"Note"}
                if not title.strip():
                    st.warning("Enter a title or requirement first.")
                    return frame
                new={"Type":types[kind],"Code":manual_code.strip(),"Course / Requirement":title.strip(),"Credits":credits,"Page":page}
            rows=frame.to_dict("records")
            rows.insert(int(insert_at),new)
            st.session_state[data_key]=pd.DataFrame(rows,columns=FIELDS)
            st.session_state[revision_key]+=1
            st.rerun()
    return frame

creation_mode=st.radio("How would you like to start?",["Upload existing PDF","Create a new program sheet"],horizontal=True)

layout_choice=st.radio(
    "How should your Excel program sheet be organized?",
    ["Automatically choose the layout", "Two columns (side by side)", "Stacked sections (top to bottom)", "Simple data table"],
    captions=[
        "Recommended. Uses the two-column format when recognized, otherwise tries stacked course sections.",
        "Use for program sheets with general education on the left and major requirements on the right.",
        "Use for program sheets with course sections listed vertically, regardless of academic level.",
        "Extract courses into a sortable spreadsheet rather than a formatted program sheet."
    ],
    index=0,
)
st.caption("You can use any layout for undergraduate, graduate, or combined programs.")
files=st.file_uploader("Upload PDF files",type=["pdf"],accept_multiple_files=True) if creation_mode=="Upload existing PDF" else []
totals=st.checkbox("Add automatic totals to simple data tables where appropriate",value=True)
st.caption("Review extracted courses and credits against the original PDF before publishing.")

if creation_mode=="Create a new program sheet":
    name=st.text_input("Program name",value="New Program")
    year=st.text_input("Catalog year",value="2027-2028")
    manual_format=st.radio("New sheet layout",["Two columns (side by side)","Stacked sections (top to bottom)"],horizontal=True)
    uid="new_"+hashlib.sha256((name+year+manual_format).encode()).hexdigest()[:12]
    output_name=f"{name} {year}.pdf"
    if manual_format=="Two columns (side by side)":
        first,second=st.columns(2)
        with first:
            left=edit_requirements(blank_rows(),uid+"_left","University Studies / Left side")
        with second:
            right=edit_requirements(blank_rows(),uid+"_right","Major, support & electives / Right side")
        if not left.empty or not right.empty:
            workbook=build_layout_workbook([(output_name,left,right)])
            st.download_button("Download new program sheet",data=workbook,file_name="new_program_sheet.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    else:
        graduate=edit_requirements(blank_rows(),uid+"_stacked","Program sections")
        if not graduate.empty:
            workbook=build_graduate_workbook([(output_name,graduate)])
            st.download_button("Download new program sheet",data=workbook,file_name="new_program_sheet.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
if files:
    if len(files)>15:
        st.error("Maximum 15 PDFs per conversion in this prototype.")
        st.stop()
    items=[]
    layout_items=[]
    grad_items=[]
    # Stable keys preserve edits across Streamlit reruns.
    for idx,up in enumerate(files):
        raw=up.getvalue()
        with st.expander(f"{up.name}",expanded=(idx==0)):
            if len(raw)>15*1024*1024:
                st.error("File exceeds 15 MB. Skipped.")
                continue
            try:
                signature=hashlib.sha256(raw).hexdigest()
                if layout_choice=='Stacked sections (top to bottom)':
                    key=f'grad_{signature}'
                    if key not in st.session_state:
                        st.session_state[key]=graduate_rows(raw)
                    graduate=edit_requirements(st.session_state[key],f'grad_{signature}','Stacked program requirements')
                    grad_items.append((up.name,graduate))
                    continue
                if layout_choice in ("Automatically choose the layout", "Two columns (side by side)"):
                    layout_key=f"layout_v3_{signature}"
                    if layout_key not in st.session_state:
                        st.session_state[layout_key]=layout_rows(raw)
                    left,right=st.session_state[layout_key]
                    if left.empty and right.empty and layout_choice=="Automatically choose the layout":
                        # Academic level does not determine PDF structure. If the
                        # two-column reader found nothing, try stacked extraction.
                        fallback_key=f"stacked_v3_{signature}"
                        if fallback_key not in st.session_state:
                            st.session_state[fallback_key]=graduate_rows(raw)
                        st.info("A stacked course layout was detected; this file will use a vertical program sheet.")
                        stacked_edit=edit_requirements(st.session_state[fallback_key],f'stacked_{signature}','Stacked program requirements')
                        grad_items.append((up.name,stacked_edit))
                        continue
                    if left.empty and right.empty:
                        st.warning("No two-column requirements were recognized. Try Automatic or Stacked sections.")
                        continue
                    st.caption("Edit either side below. Course/requirement and credit fields are editable, including blank electives. Choose Section in the Type column to create a section heading.")
                    col_left,col_right=st.columns(2)
                    with col_left:
                        st.markdown("**University Studies / Left side**")
                        left_edit=edit_requirements(left,f"left_{signature}","University Studies / Left side")
                    with col_right:
                        st.markdown("**Major, support & electives / Right side**")
                        right_edit=edit_requirements(right,f"right_{signature}","Major, support & electives / Right side")
                    layout_items.append((up.name,left_edit,right_edit))
                    continue
                cache_key=f"parsed_{signature}"
                if cache_key not in st.session_state:
                    st.session_state[cache_key]=extract_pdf(raw,up.name)
                original, meta=st.session_state[cache_key]
                st.write(f"**Detected:** {meta['mode']} · **Pages:** {meta['pages']} · **Rows:** {meta['rows']}")
                st.caption(meta['warning'])
                edited=st.data_editor(original,key=f"editor_{signature}",num_rows="dynamic",hide_index=True,use_container_width=True)
                if meta['mode']=="Program sheet":
                    st.caption("Credits are totaled in Excel when enabled. Review titles and credit values before downloading.")
                items.append((up.name,edited,meta["mode"]))
            except Exception as ex:
                st.error(f"Could not process this file: {ex}")
    if grad_items:
        try:
            workbook=build_graduate_workbook(grad_items)
            st.download_button('Download graduate program sheets',data=workbook,file_name='graduate_program_sheets.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
        except Exception as ex:
            st.error(f'Could not create graduate workbook: {ex}')
    if layout_items:
        try:
            workbook=build_layout_workbook(layout_items)
            st.download_button("Download editable program sheets",data=workbook,file_name="editable_program_sheets.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
            st.caption("To create a PDF after editing, open the workbook in Excel and choose File > Export > Create PDF/XPS. Review the print preview and credits before publishing.")
        except Exception as ex:
            st.error(f"Could not create the editable program sheet: {ex}")
    if items:
        try:
            workbook=build_workbook(items,add_totals=totals)
            st.download_button("Download data table workbook",data=workbook,file_name="pdf_data_tables.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
        except Exception as ex:
            st.error(f"Could not create workbook: {ex}")
