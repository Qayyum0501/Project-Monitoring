import streamlit as st
import numpy as np
from datetime import datetime
import pandas as pd
import json
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(page_title="Project Monitoring", layout="wide")
st.title("Monitoring Program Prioritas untuk masing-masing Ekosistem")

# =========================
# GOOGLE DRIVE AUTH VIA STREAMLIT SECRETS 
# =========================
sa_json = st.secrets["SERVICE_ACCOUNT_JSON"]

# Buat credentials dari dict
credentials = ServiceAccountCredentials.from_json_keyfile_dict(
    json.loads(sa_json),
    scopes=["https://www.googleapis.com/auth/drive"]
)

# GoogleAuth + Drive
gauth = GoogleAuth()
gauth.credentials = credentials
drive = GoogleDrive(gauth)


# =========================
# FOLDER & FILE SELECTION
# =========================
folder_id = "1ETvoV7t4lZjKLjsNthuCrxwMaD19K-3Y"
file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false"}).GetList()
excel_files = {f['title']: f['id'] for f in file_list if f['title'].endswith('.xlsx')}

if not excel_files:
    st.warning("⚠️ Tidak ada file Excel di folder Google Drive")
    st.stop()

selected_file = st.selectbox("Pilih Ekosistem / File", list(excel_files.keys()))

if selected_file:
    file_id = excel_files[selected_file]
    downloaded = drive.CreateFile({'id': file_id})
    downloaded.GetContentFile(selected_file)
    
    df = pd.read_excel(selected_file, header=8)
    df = df.iloc[:, 1:]
    
    # =========================
    # CLEAN COLUMN
    # =========================
    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={
        'outline number': 'Outline number',
        'task name': 'Name',
        'name': 'Name',
        '% complete': '% complete',
        'start': 'Start',
        'finish': 'Finish',
        'duration': 'Duration',
        'bucket': 'Entitas'
    })
    required_cols = ['Outline number','Name','Start','Finish','Duration','% complete']
    for col in required_cols:
        if col not in df.columns:
            st.error(f"Column '{col}' not found")
            st.stop()
    df['Outline number'] = df['Outline number'].astype(str)
    
    # =========================
    # FORMAT DATE
    # =========================
    df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
    df['Finish'] = pd.to_datetime(df['Finish'], errors='coerce')

    today = today = datetime.today()
    target_date = st.date_input("Pilih Target Tanggal", value=today)
    target_date = datetime.combine(target_date, datetime.min.time())
    
    def networkdays(start, end):
        if pd.isna(start) or start > end:
            return 0
        return np.busday_count(start.date(), end.date())
    
    df['duration_todate'] = df['Start'].apply(lambda x: networkdays(x, target_date))
    
    df['Duration'] = (
        df['Duration'].astype(str)
        .str.replace(' days','', regex=False)
        .str.replace(' day','', regex=False)
        .astype(float).fillna(0).round().astype(int)
    )
    
    # =========================
    # BASELINE
    # =========================
    df['Baseline'] = (df['duration_todate']/df['Duration'].replace(0,1)).clip(upper=1)*100
    
    # =========================
    # WEIGHTED BASELINE
    # =========================
    def weighted_progress(parent):
        children = df[df['Outline number'].str.startswith(parent) & (df['Outline number'] != parent)]
        if children.empty:
            val = df.loc[df['Outline number']==parent,'Baseline']
            return val.values[0] if len(val)>0 else 0
        total_duration = children['Duration'].replace(0,1).sum()
        return (children['Baseline']*children['Duration']).sum()/total_duration
    
    df['Baseline_progress'] = df['Outline number'].apply(weighted_progress)
    
    df['% complete'] = df['% complete'].fillna(0)*100
    
    # =========================
    # DELAY
    # =========================
    def calc_delay(row):
        if pd.isna(row['Finish']):
            return 0
        if row['Finish'] < target_date and row['% complete']<100:
            return np.busday_count(row['Finish'].date(), target_date.date())
        return 0
    
    df['Delay (days)'] = df.apply(calc_delay, axis=1)
    
    # =========================
    # STATUS
    # =========================
    df['Status'] = np.select(
        [
            df['% complete']==0,
            df['% complete']>=100,
            (df['Finish']<target_date) & (df['% complete']<100),
            df['% complete']>=0.9*df['Baseline_progress'],
            df['% complete']>=0.75*df['Baseline_progress']
        ],
        ['Not Started','Complete','Late','On Progress','Concern'],
        default='Late'
    )
    
    # =========================
    # KPI FUNCTION
    # =========================
    def calc_kpi(data):
        total_duration = data['Duration'].replace(0,1).sum()
        progress = (data['% complete']*data['Duration']).sum()/total_duration
        baseline = (data['Baseline']*data['Duration']).sum()/total_duration
        delta = progress - baseline
        return round(progress,1), round(baseline,1), round(delta,1)
    
    def kpi_box(title, progress, baseline, delta):
        color = "green" if delta>=0 else "orange" if delta>=-5 else "red"
        st.markdown(f"""
        <div style="border-radius:15px;padding:20px;text-align:center;background:#f9f9f9;margin-bottom:10px;">
            <div style="font-size:18px;font-weight:bold;">{title}</div>
            <div style="font-size:40px;font-weight:bold;color:{color};">{progress:.1f}%</div>
            <div style="font-size:14px;color:gray;">Baseline: {baseline:.1f}%</div>
            <div style="font-size:18px;color:{color};">Δ {delta:+.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    # =========================
    # OVERALL PROJECT KPI
    # =========================
    total_duration = df['Duration'].replace(0,1).sum()
    overall_progress = (df['% complete']*df['Duration']).sum()/total_duration
    overall_baseline = (df['Baseline_progress']*df['Duration']).sum()/total_duration
    overall_delta = overall_progress - overall_baseline
    color = "green" if overall_delta>=0 else "orange" if overall_delta>=-5 else "red"
    
    st.markdown(f"""
    <div style="border-radius:15px;padding:20px;text-align:center;background:#e0e0e0;margin-bottom:20px;">
        <div style="font-size:20px;font-weight:bold;">Overall Project Progress</div>
        <div style="font-size:50px;font-weight:bold;color:{color};">{overall_progress:.1f}%</div>
        <div style="font-size:16px;color:gray;">Baseline: {overall_baseline:.1f}%</div>
        <div style="font-size:20px;color:{color};">Δ {overall_delta:+.1f}%</div>
    </div>
    """, unsafe_allow_html=True)
    
    # =========================
    # KPI PER SUB EKOSISTEM
    # =========================
    st.markdown("## 📊 Progress per Masing - Masing Objek Pada Ekosistem")
    level2 = df[df['Outline number'].str.count(r'\.')==1]
    if level2.empty:
        st.warning("⚠️ Tidak ada data level 2")
        st.stop()
    cols = st.columns(3)
    for i, (_, row) in enumerate(level2.iterrows()):
        code = row['Outline number']
        sub_data = df[df['Outline number'].str.startswith(code)]
        p,b,d = calc_kpi(sub_data)
        with cols[i%3]:
            kpi_box(row['Name'],p,b,d)
    
    # =========================
    # SELECT SUB EKOSISTEM
    # =========================
    st.divider()
    selected_l2 = st.selectbox("Pilih Objek Ekosistem", level2['Name'])
    selected_row = level2[level2['Name']==selected_l2]
    sub_tasks = pd.DataFrame()
    if not selected_row.empty:
        selected_l2_code = selected_row['Outline number'].values[0]
        sub_tasks = df[df['Outline number'].str.startswith(selected_l2_code)].copy()
    
    # =========================
    # FILTER BUTTON STATUS
    # =========================
    if "status_filter" not in st.session_state:
        st.session_state.status_filter = "ALL"
    
    col1, col2, col3, col4 = st.columns(4)
    total_complete = len(sub_tasks[sub_tasks['Status']=="Complete"])
    total_progress = len(sub_tasks[sub_tasks['Status']=="On Progress"])
    total_concern = len(sub_tasks[sub_tasks['Status']=="Concern"])
    total_late = len(sub_tasks[sub_tasks['Status']=="Late"])
    
    with col1:
        if st.button(f"✅ Complete ({total_complete})"):
            st.session_state.status_filter="Complete"
    with col2:
        if st.button(f"🟢 On Progress ({total_progress})"):
            st.session_state.status_filter="On Progress"
    with col3:
        if st.button(f"🟠 Concern ({total_concern})"):
            st.session_state.status_filter="Concern"
    with col4:
        if st.button(f"🔴 Late ({total_late})"):
            st.session_state.status_filter="Late"
    if st.button("🔄 Show All"):
        st.session_state.status_filter="ALL"
    
    # =========================
    # FILTER DATA
    # =========================
    filtered_tasks = sub_tasks.copy()
    if st.session_state.status_filter!="ALL":
        filtered_tasks = filtered_tasks[filtered_tasks['Status']==st.session_state.status_filter]
    
    # =========================
    # DISPLAY TABLE
    # =========================
    st.subheader("📋 Detail Task Summary")
    if filtered_tasks.empty:
        st.warning("⚠️ Tidak ada task")
    else:
        filtered_tasks['Level'] = filtered_tasks['Outline number'].apply(lambda x: x.count('.'))
        filtered_tasks['Name WBS'] = filtered_tasks.apply(lambda row: "   "*(row['Level']-1)+"▸ "+row['Name'], axis=1)
        filtered_tasks = filtered_tasks.sort_values('Outline number')
        display_df = filtered_tasks.copy()
        display_df['Start'] = display_df['Start'].dt.strftime('%d/%m/%Y')
        display_df['Finish'] = display_df['Finish'].dt.strftime('%d/%m/%Y')
        display_df['Progress (%)'] = display_df['% complete'].map(lambda x: f"{x:.1f}%")
        display_df['Baseline (%)'] = display_df['Baseline'].map(lambda x: f"{x:.1f}%")
        display_df = display_df[['Outline number','Name WBS','Entitas','Start','Finish','Progress (%)','Baseline (%)','Delay (days)','Status']]
        st.dataframe(display_df,use_container_width=True)
    
    # =========================
    # HEAT MAP ENTITAS STATUS
    # =========================
    st.markdown("## 🔥 Heatmap Entitas")
    
    for status,color_hex in zip(['Late','On Progress','Concern'], ['#ff4d4d','#66cc66','#ffa64d']):
        st.markdown(f"### {status}")
        summary = df[df['Status']==status].groupby('Entitas').size().reset_index(name='Total').sort_values('Total',ascending=False)
        cols = st.columns(4)
        for i,row in summary.iterrows():
            entitas = row['Entitas']
            total = row['Total']
            with cols[i%4]:
                if st.button(f"{entitas} ({total})", key=f"{status}_{entitas}"):
                    st.session_state.selected_entitas = entitas
                    st.session_state.selected_status = status
                st.markdown(f"""
                <div style="
                    background-color:{color_hex};
                    padding:15px;
                    border-radius:15px;
                    text-align:center;
                    font-weight:bold;
                    color:white;
                    margin-bottom:5px;
                ">
                    {entitas}<br>{total}
                </div>
                """, unsafe_allow_html=True)
    
    # =========================
    # TABEL DETAIL ENTITAS TERPILIH
    # =========================
    if "selected_entitas" in st.session_state and "selected_status" in st.session_state:
        selected = st.session_state.selected_entitas
        status = st.session_state.selected_status
        st.markdown(f"## 📋 Detail Task - {selected} ({status})")
        detail = df[(df['Entitas']==selected) & (df['Status']==status)].copy()
        if not detail.empty:
            detail['Level'] = detail['Outline number'].apply(lambda x: x.count('.'))
            detail['Name WBS'] = detail.apply(lambda row: "   "*(row['Level']-1)+"▸ "+row['Name'], axis=1)
            detail['Start'] = detail['Start'].dt.strftime('%d/%m/%Y')
            detail['Finish'] = detail['Finish'].dt.strftime('%d/%m/%Y')
            detail['Progress (%)'] = detail['% complete'].map(lambda x: f"{x:.1f}%")
            detail['Baseline (%)'] = detail['Baseline'].map(lambda x: f"{x:.1f}%")
            display_df = detail[['Outline number','Name WBS','Entitas','Start','Finish','Progress (%)','Baseline (%)','Delay (days)','Status']]
            st.dataframe(display_df,use_container_width=True)
        else:
            st.info("Tidak ada task untuk status ini")
