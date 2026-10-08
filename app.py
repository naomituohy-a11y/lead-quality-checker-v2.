import io
import hashlib
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from checker import ALIASES, PICK_FIELDS, REL_COLS, detect, read_frame, relationships, process

st.set_page_config(page_title='Lead Quality Checker', page_icon='✅', layout='wide')
st.title('Lead Quality Checker')
st.caption('Flags for review. Does not verify employment, mailbox ownership or deliverability. Original lead fields are preserved.')
st.info('Green: recorded official domain • Amber: connected organisation • Yellow: unverified/review • Red: reviewed conflict • Grey: not checked')
a,b = st.columns(2)
master = a.file_uploader('Lead master (.xlsx)', type=['xlsx'])
pick = b.file_uploader('Picklist (.xlsx), optional', type=['xlsx'])
rel = st.file_uploader('Verified relationships (.csv or .xlsx), optional', type=['csv','xlsx'])
st.caption('Relationship records must be manually verified and include a source URL and date. No automatic web or AI research is performed in this version. Evidence expires after 180 days.')
st.download_button('Download blank relationship template', pd.DataFrame(columns=REL_COLS).to_csv(index=False), 'relationships.csv', 'text/csv')

def select_sheet(upload, name):
    wb = load_workbook(io.BytesIO(upload.getvalue()), read_only=True)
    names = wb.sheetnames
    wb.close()
    return st.selectbox(name, names)

def choose_mapping(df, fields, prefix):
    auto = detect(df)
    result = {}
    with st.expander(prefix + ' column mapping — check before running', expanded=True):
        opts = ['(Not available)'] + list(df.columns)
        cols = st.columns(3)
        for i,f in enumerate(fields):
            selected = cols[i%3].selectbox(f.replace('_',' ').title(), opts, index=opts.index(auto[f]) if auto.get(f) in opts else 0, key=prefix+f)
            result[f] = None if selected == opts[0] else selected
    return result

if master:
    try:
        ms = select_sheet(master, 'Master sheet')
        mapping = choose_mapping(read_frame(master.getvalue(), ms), ALIASES, 'Master')
        ps, pm = 0, None
        if pick:
            ps = select_sheet(pick, 'Picklist sheet')
            pm = choose_mapping(read_frame(pick.getvalue(), ps), PICK_FIELDS, 'Picklist')
        skip = st.checkbox('Skip instruction/template rows (3 or more template markers)', True)
        colors = st.checkbox('Colour-code QA results', True)
        signature = hashlib.sha256(master.getvalue() + (pick.getvalue() if pick else b'') + (rel.getvalue() if rel else b'') + repr((ms,ps,mapping,pm,skip,colors)).encode()).hexdigest()
        if st.session_state.get('signature') != signature:
            st.session_state.pop('output', None)
        if st.button('Run checks', type='primary'):
            st.session_state.pop('output', None)
            with st.spinner('Checking leads…'):
                rels = relationships(rel.getvalue(), rel.name) if rel else {}
                result, preview = process(master.getvalue(), ms, mapping, pick.getvalue() if pick else None, ps, pm, rels, skip, colors)
                st.session_state['output'] = (result, preview)
                st.session_state['signature'] = signature
        if 'output' in st.session_state:
            result, preview = st.session_state['output']
            st.dataframe(preview, use_container_width=True)
            st.download_button('Download checked workbook', result, 'Lead_QA_Results.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as exc:
        st.session_state.pop('output', None)
        st.error(str(exc))
