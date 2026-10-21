"""Run: streamlit run app.py"""
from __future__ import annotations
import hashlib
import streamlit as st
from converter import extract_pdf, build_workbook, layout_rows, build_layout_workbook, graduate_rows, build_graduate_workbook

st.set_page_config(page_title="PDF to Excel for Jane!",page_icon="📊",layout="wide")
st.title("PDF to Excel for Jane")
st.write("Upload multiple PDFs and get one editable Excel workbook, with one worksheet per PDF.")
st.caption("Prototype: works best with selectable-text PDFs. Scanned/image-only PDFs are not supported yet.")

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
files=st.file_uploader("Upload PDF files",type=["pdf"],accept_multiple_files=True)
totals=st.checkbox("Add automatic totals to simple data tables where appropriate",value=True)
st.caption("Review extracted courses and credits against the original PDF before publishing.")

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
                    graduate=st.data_editor(st.session_state[key], key=f'grad_editor_{signature}', num_rows='dynamic', hide_index=True, use_container_width=True)
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
                        stacked_edit=st.data_editor(st.session_state[fallback_key],key=f"stacked_editor_v3_{signature}",num_rows="dynamic",hide_index=True,use_container_width=True)
                        grad_items.append((up.name,stacked_edit))
                        continue
                    if left.empty and right.empty:
                        st.warning("No two-column requirements were recognized. Try Automatic or Stacked sections.")
                        continue
                    st.caption("Edit either side below. Course/requirement and credit fields are editable, including blank electives. Choose Section in the Type column to create a section heading.")
                    col_left,col_right=st.columns(2)
                    with col_left:
                        st.markdown("**University Studies / Left side**")
                        left_edit=st.data_editor(left,key=f"left_v3_{signature}",num_rows="dynamic",hide_index=True,use_container_width=True)
                    with col_right:
                        st.markdown("**Major, support & electives / Right side**")
                        right_edit=st.data_editor(right,key=f"right_v3_{signature}",num_rows="dynamic",hide_index=True,use_container_width=True)
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
