"""Run: streamlit run app.py"""
from __future__ import annotations
import hashlib
import streamlit as st
from converter import extract_pdf, build_workbook, layout_rows, build_layout_workbook, graduate_rows, build_graduate_workbook

st.set_page_config(page_title="PDF to Excel for Jane!",page_icon="📊",layout="wide")
st.title("PDF to Excel for Jane")
st.write("Upload multiple PDFs and get one editable Excel workbook, with one worksheet per PDF.")
st.caption("Prototype: works best with selectable-text PDFs. Scanned/image-only PDFs are not supported yet.")

program_type=st.radio('Program type', ['Undergraduate', 'Graduate'], horizontal=True)
if program_type=='Undergraduate':
    st.caption('Converts existing PDF program sheets to a standardized two-column Excel layout. A future approved University Studies template can replace the left side.')
else:
    st.caption('Converts graduate program sheets into editable Excel with vertically stacked sections and notes. Different degree totals are preserved.')

if program_type=='Undergraduate':
    format_choice=st.radio(
    "How would you like your Excel file organized?",
    ["Keep the document layout", "Make a simple data table"],
    captions=[
        "Best for editing and reprinting a program sheet. Keeps University Studies on the left, major courses and electives on the right.",
        "Best for sorting, filtering, or analyzing information. Places extracted rows in a regular spreadsheet, not a printable program-sheet layout."
    ],
    index=0,
    horizontal=False,
    )
else:
    format_choice='Keep the document layout'
files=st.file_uploader("Upload PDF files",type=["pdf"],accept_multiple_files=True)
totals=st.checkbox("Add automatic totals where appropriate",value=True,help="will change.")
st.info(" spreadsheet styling, filters, frozen headings, and worksheet tabs are always enabled.")

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
                if program_type=='Graduate':
                    key=f'grad_{signature}'
                    if key not in st.session_state:
                        st.session_state[key]=graduate_rows(raw)
                    graduate=st.data_editor(st.session_state[key], key=f'grad_editor_{signature}', num_rows='dynamic', hide_index=True, use_container_width=True)
                    grad_items.append((up.name,graduate))
                    continue
                if format_choice=="Keep the document layout":
                    layout_key=f"layout_{signature}"
                    if layout_key not in st.session_state:
                        st.session_state[layout_key]=layout_rows(raw)
                    left,right=st.session_state[layout_key]
                    st.caption("Edit either side below. Course/requirement and credit fields are editable, including blank electives. Choose Section in the Type column to create a section heading.")
                    col_left,col_right=st.columns(2)
                    with col_left:
                        st.markdown("**University Studies / Left side**")
                        left_edit=st.data_editor(left,key=f"left_{signature}",num_rows="dynamic",hide_index=True,use_container_width=True)
                    with col_right:
                        st.markdown("**Major, support & electives / Right side**")
                        right_edit=st.data_editor(right,key=f"right_{signature}",num_rows="dynamic",hide_index=True,use_container_width=True)
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
