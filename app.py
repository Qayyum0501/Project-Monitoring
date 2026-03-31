import streamlit as st
import numpy as np
from datetime import datetime, date
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
tab1, tab2 = st.tabs(["🌍 Overall Ekosistem", "📊 Detail Monitoring"])

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
# GET FILES
# =========================
folder_id = "1ETvoV7t4lZjKLjsNthuCrxwMaD19K-3Y"
file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false"}).GetList()
excel_files = {f['title']: f['id'] for f in file_list if f['title'].endswith('.xlsx')}

if not excel_files:
    st.warning("⚠️ Tidak ada file Excel di folder Google Drive")
    st.stop()

# =========================
# HELPER FUNCTIONS
# =========================

def networkdays(start, end):
    if pd.isna(start) or start > end:
        return 0
    return np.busday_count(start.date(), end.date())

def get_status(progress, baseline, start_date=None, finish_date=None, target_date=None):

    today = target_date

    # =========================
    # FORCE CLEAN TYPES
    # =========================
    if pd.isna(start_date):
        start_date = None
    if pd.isna(finish_date):
        finish_date = None

    # =========================
    # 1. COMPLETE
    # =========================
    if progress >= 100:
        return "Complete"

    # =========================
    # 2. LATE (TIME-BASED)
    # =========================
    if finish_date is not None and today is not None:
        if today > finish_date and progress < 100:
            return "Late"

    # =========================
    # 3. NOT STARTED (SAFE MODE)
    # =========================
    if progress == 0:

        # kalau start_date valid baru compare
        if start_date is not None and today is not None:
            if today < start_date:
                return "Not Started"
            else:
                return "Late"   # sudah harus mulai tapi belum jalan

        # kalau start_date kosong → anggap risky
        return "Late"

    # =========================
    # 4. PERFORMANCE BASED
    # =========================
    if baseline > 0:
        ratio = progress / baseline

        if ratio < 0.75:
            return "Late"
        elif ratio < 0.9:
            return "Concern"
        else:
            return "On Progress"

    return "On Progress"

def get_color(status):
    return {
        "Not Started": "black",
        "Complete": "blue",
        "On Progress": "green",
        "Concern": "orange",
        "Late": "red"
    }.get(str(status), "black")

def kpi_box(title, progress, baseline):
    delta = progress - baseline
    status = get_status(progress, baseline)
    color = get_color(status)

    st.markdown(f"""
    <div style="border-radius:20px;padding:25px;text-align:center;background:#f5f5f5;margin-bottom:20px;">
        <div style="font-size:20px;font-weight:bold;">{title}</div>
        <div style="font-size:50px;font-weight:bold;color:{color};">{progress:.1f}%</div>
        <div style="font-size:16px;color:gray;">Baseline: {baseline:.1f}%</div>
        <div style="font-size:20px;color:{color};">Δ {delta:+.1f}%</div>
        <div style="font-size:14px;color:{color};">{status}</div>
    </div>
    """, unsafe_allow_html=True)

def load_and_process(file_name, file_id, target_date):
    # LOAD
    downloaded = drive.CreateFile({'id': file_id})
    downloaded.GetContentFile(file_name)

    df = pd.read_excel(file_name, header=8)
    df = df.iloc[:, 1:]

    # CLEANING
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

    df['Outline number'] = df['Outline number'].astype(str)
    df['% complete'] = pd.to_numeric(df['% complete'], errors='coerce').fillna(0) * 100

    df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
    df['Finish'] = pd.to_datetime(df['Finish'], errors='coerce')

    df['Duration'] = (
        df['Duration'].astype(str)
        .str.replace(' days','')
        .str.replace(' day','')
    )
    df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce').fillna(0)

    # BASELINE
    df['duration_todate'] = df['Start'].apply(lambda x: networkdays(x, target_date))
    df['Baseline'] = (
        df['duration_todate'] /
        df['Duration'].replace(0,1)
    ).clip(upper=1) * 100

    # LEAF
    df['is_parent'] = df['Outline number'].apply(
        lambda x: any(df['Outline number'].str.startswith(x + '.'))
    )
    leaf = df[~df['is_parent']].copy()

    return df, leaf

def aggregate(leaf, code):
    subset = leaf[leaf['Outline number'].str.startswith(code)]

    if subset.empty:
        return 0, 0

    total_dur = subset['Duration'].replace(0,1).sum()

    progress = (subset['% complete'] * subset['Duration']).sum() / total_dur
    baseline = (subset['Baseline'] * subset['Duration']).sum() / total_dur

    return progress, baseline

# =========================
# TAB 1
# =========================
with tab1:

    st.subheader("🌍 Overall Ekosistem")

    target_date = st.date_input("Target Date Global", value=date.today())
    target_date = datetime.combine(target_date, datetime.min.time())

    ecosystem_results = []

    for file_name, file_id in excel_files.items():
        df, leaf = load_and_process(file_name, file_id, target_date)

        progress, baseline = aggregate(leaf, "1")

        ecosystem_results.append({
            "name": file_name.replace(".xlsx",""),
            "progress": progress,
            "baseline": baseline,
            "duration": leaf['Duration'].replace(0,1).sum()
        })

    df_eco = pd.DataFrame(ecosystem_results)

    # GLOBAL
    total_weight = df_eco['duration'].sum()

    global_progress = (df_eco['progress'] * df_eco['duration']).sum() / total_weight
    global_baseline = (df_eco['baseline'] * df_eco['duration']).sum() / total_weight

    kpi_box("Overall Project Global", global_progress, global_baseline)

    # PER EKOSISTEM
    st.markdown("## 📊 Progress per Ekosistem")

    cols = st.columns(3)
    for i, row in df_eco.iterrows():
        with cols[i % 3]:
            kpi_box(row['name'], row['progress'], row['baseline'])

# =========================
# TAB 2 (UNCHANGED LOGIC, CLEANED STRUCTURE)
# =========================
with tab2:

    selected_file = st.selectbox("Pilih Ekosistem / File", list(excel_files.keys()))

    if selected_file:

        df, leaf = load_and_process(
            selected_file,
            excel_files[selected_file],
            datetime.combine(
                st.date_input("Pilih Target Tanggal", value=datetime.today()),
                datetime.min.time()
            )
        )

        # KPI
        overall_p, overall_b = aggregate(leaf, "1")
        kpi_box("Overall Project", overall_p, overall_b)

        # LEVEL 2
        st.markdown("## 📊 Progress per Ekosistem")

        level2 = df[df['Outline number'].str.match(r'^\d+\.\d+$')]
        cols = st.columns(3)

        for i, (_, row) in enumerate(level2.iterrows()):
            p, b = aggregate(leaf, row['Outline number'])
            with cols[i % 3]:
                kpi_box(row['Name'], p, b)

        # SELECT
        st.divider()

        selected_l2 = st.selectbox("Pilih Objek Ekosistem", level2['Name'])
        selected_code = level2[level2['Name']==selected_l2]['Outline number'].values[0]

        sub_tasks = leaf[leaf['Outline number'].str.startswith(selected_code)].copy()

        # STATUS
        sub_tasks['Status'] = sub_tasks.apply(
            lambda x: get_status(x['% complete'], x['Baseline'], x['Finish'],target_date), axis=1
        )

        # FILTER
        if "status_filter" not in st.session_state:
            st.session_state.status_filter = "ALL"

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            if st.button(f"✅ Complete ({len(sub_tasks[sub_tasks['Status']=='Complete'])})"):
                st.session_state.status_filter = "Complete"

        with col2:
            if st.button(f"🟢 On Progress ({len(sub_tasks[sub_tasks['Status']=='On Progress'])})"):
                st.session_state.status_filter = "On Progress"

        with col3:
            if st.button(f"🟠 Concern ({len(sub_tasks[sub_tasks['Status']=='Concern'])})"):
                st.session_state.status_filter = "Concern"

        with col4:
            if st.button(f"🔴 Late ({len(sub_tasks[sub_tasks['Status']=='Late'])})"):
                st.session_state.status_filter = "Late"

        if st.button("🔄 Show All"):
            st.session_state.status_filter = "ALL"

        filtered = sub_tasks if st.session_state.status_filter=="ALL" else \
                   sub_tasks[sub_tasks['Status']==st.session_state.status_filter]

        # TABLE
        st.subheader("📋 Detail Task Summary")

        if filtered.empty:
            st.warning("Tidak ada data")
        else:
            filtered['Level'] = filtered['Outline number'].apply(lambda x: x.count('.'))
            filtered['Name WBS'] = filtered.apply(
                lambda r: "   "*(r['Level']-1)+"▸ "+r['Name'], axis=1
            )

            filtered['Start'] = filtered['Start'].dt.strftime('%d/%m/%Y')
            filtered['Finish'] = filtered['Finish'].dt.strftime('%d/%m/%Y')
            filtered['Progress (%)'] = filtered['% complete'].map(lambda x: f"{x:.1f}%")
            filtered['Baseline (%)'] = filtered['Baseline'].map(lambda x: f"{x:.1f}%")

            st.dataframe(filtered[
                ['Outline number','Name WBS','Entitas',
                 'Start','Finish','Progress (%)','Baseline (%)','Status']
            ], use_container_width=True)

        # =========================
        # 🎯 FILTER BERDASARKAN ENTITAS
        # =========================
        st.divider()
        st.markdown("## 🏢 Monitoring Berdasarkan Entitas")

        # ambil semua entitas unik dari leaf
        all_entitas = sorted(leaf['Entitas'].dropna().unique())

        selected_entitas_filter = st.selectbox(
            "Pilih Entitas",
            all_entitas
        )

        # filter semua task berdasarkan entitas
        entitas_tasks = leaf[leaf['Entitas'] == selected_entitas_filter].copy()

        # =========================
        # HITUNG KPI ENTITAS
        # =========================
        total_dur = entitas_tasks['Duration'].replace(0,1).sum()

        entitas_progress = (entitas_tasks['% complete'] * entitas_tasks['Duration']).sum() / total_dur
        entitas_baseline = (entitas_tasks['Baseline'] * entitas_tasks['Duration']).sum() / total_dur

        # KPI BOX
        kpi_box(f"Performance - {selected_entitas_filter}", entitas_progress, entitas_baseline)

        # =========================
        # STATUS PER TASK
        # =========================
        entitas_tasks['Status'] = entitas_tasks.apply(
            lambda x: get_status(x['% complete'], x['Baseline'],x['Finish'], target_date),
            axis=1
        )

        # =========================
        # FILTER STATUS (SAMA PERSIS)
        # =========================
        if "entitas_filter_status" not in st.session_state:
            st.session_state.entitas_filter_status = "ALL"

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            if st.button(f"✅ Complete ({len(entitas_tasks[entitas_tasks['Status']=='Complete'])})", key="ent_c"):
                st.session_state.entitas_filter_status = "Complete"

        with col2:
            if st.button(f"🟢 On Progress ({len(entitas_tasks[entitas_tasks['Status']=='On Progress'])})", key="ent_p"):
                st.session_state.entitas_filter_status = "On Progress"

        with col3:
            if st.button(f"🟠 Concern ({len(entitas_tasks[entitas_tasks['Status']=='Concern'])})", key="ent_con"):
                st.session_state.entitas_filter_status = "Concern"

        with col4:
            if st.button(f"🔴 Late ({len(entitas_tasks[entitas_tasks['Status']=='Late'])})", key="ent_l"):
                st.session_state.entitas_filter_status = "Late"

        if st.button("🔄 Show All (Entitas)", key="ent_all"):
            st.session_state.entitas_filter_status = "ALL"

        # =========================
        # APPLY FILTER
        # =========================
        filtered_entitas = entitas_tasks.copy()

        if st.session_state.entitas_filter_status != "ALL":
            filtered_entitas = filtered_entitas[
                filtered_entitas['Status'] == st.session_state.entitas_filter_status
            ]

        # =========================
        # DISPLAY TABLE (SAMA FORMAT)
        # =========================
        st.subheader("📋 Detail Task per Entitas")

        if filtered_entitas.empty:
            st.warning("Tidak ada data")
        else:
            filtered_entitas['Level'] = filtered_entitas['Outline number'].apply(lambda x: x.count('.'))

            filtered_entitas['Name WBS'] = filtered_entitas.apply(
                lambda r: "   "*(r['Level']-1) + "▸ " + r['Name'],
                axis=1
            )

            filtered_entitas['Start'] = filtered_entitas['Start'].dt.strftime('%d/%m/%Y')
            filtered_entitas['Finish'] = filtered_entitas['Finish'].dt.strftime('%d/%m/%Y')
            filtered_entitas['Progress (%)'] = filtered_entitas['% complete'].map(lambda x: f"{x:.1f}%")
            filtered_entitas['Baseline (%)'] = filtered_entitas['Baseline'].map(lambda x: f"{x:.1f}%")

            display = filtered_entitas[[
                'Outline number','Name WBS','Entitas',
                'Start','Finish','Progress (%)','Baseline (%)','Status'
            ]]

            st.dataframe(display, use_container_width=True)
