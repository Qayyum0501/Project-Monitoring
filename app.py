import streamlit as st
import numpy as np
from datetime import datetime, date
import pandas as pd
import json
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials
import base64
import re

def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

logo_base64 = get_base64_image("logo_pelindo.png")

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Login - Dashboard Pelindo", layout="centered")

# =========================
# LOAD LOGO
# =========================
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

logo_base64 = get_base64_image("logo_pelindo.png")

# =========================
# LOGIN FUNCTION
# =========================
def check_login(username, password):
    return (
        username == st.secrets["auth"]["username"]
        and password == st.secrets["auth"]["password"]
    )

# =========================
# SESSION
# =========================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# =========================
# LOGIN PAGE
# =========================
if not st.session_state.logged_in:

    st.markdown(f"""
    <style>

    /* =========================
       BACKGROUND (PREMIUM DARK)
    ========================= */
    .stApp {{
        background: linear-gradient(135deg, #000000, #020617, #0b2c5a);
    }}

    /* CENTERING */
    .wrapper {{
        display: flex;
        justify-content: center;
        align-items: center;
        height: 95vh;
    }}

    /* =========================
       GLASS CARD
    ========================= */
    .card {{
        width: 420px;
        border-radius: 24px;
        backdrop-filter: blur(18px);
        background: rgba(255,255,255,0.08);
        border: 1px solid rgba(255,255,255,0.15);
        box-shadow: 0 20px 60px rgba(0,0,0,0.6);
        overflow: hidden;
        transition: all 0.3s ease;
    }}

    .card:hover {{
        transform: translateY(-5px);
        box-shadow: 0 30px 80px rgba(0,0,0,0.8);
    }}

    /* HEADER */
    .header {{
        text-align: center;
        padding: 35px 30px 20px 30px;
    }}

    .header img {{
        height: 75px;
        margin-bottom: 15px;
        filter: drop-shadow(0 4px 10px rgba(255,255,255,0.2));
    }}

    .title {{
        color: white;
        font-size: 22px;
        font-weight: 700;
    }}

    .subtitle {{
        color: #94a3b8;
        font-size: 13px;
    }}

    /* FORM AREA (WHITE CLEAN) */
    .form-container {{
        background: white;
        padding: 30px;
        border-top-left-radius: 24px;
        border-top-right-radius: 24px;
    }}

    /* INPUT FIX */
    input {{
        background-color: #f9fafb !important;
        color: black !important;
        border-radius: 10px !important;
        border: 1px solid #e5e7eb !important;
        padding: 10px !important;
        transition: all 0.2s ease;
    }}

    input:focus {{
        border: 1px solid #0b2c5a !important;
        box-shadow: 0 0 0 2px rgba(11,44,90,0.2) !important;
    }}

    /* LABEL */
    label {{
        color: #111827 !important;
        font-weight: 600;
    }}

    /* BUTTON */
    .stButton button {{
        width: 100%;
        height: 48px;
        border-radius: 12px;
        background: linear-gradient(90deg, #0b2c5a, #1e3a8a);
        color: white;
        font-weight: 700;
        border: none;
        transition: all 0.3s ease;
    }}

    .stButton button:hover {{
        transform: scale(1.03);
        box-shadow: 0 10px 25px rgba(11,44,90,0.4);
    }}

    /* ERROR TEXT */
    .error {{
        color: #ef4444;
        font-size: 13px;
        margin-top: 10px;
    }}

    </style>

    <div class="wrapper">
        <div class="card">
            <div class="header">
                <img src="data:image/png;base64,{logo_base64}">
                <div class="title">Dashboard Monitoring</div>
                <div class="subtitle">Program Prioritas Pelindo Group</div>
            </div>
            <div class="form-container">
    """, unsafe_allow_html=True)

    # =========================
    # FORM LOGIN
    # =========================
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        login_btn = st.form_submit_button("Login")

        if login_btn:
            if check_login(username, password):
                st.session_state.logged_in = True
                st.success("Login berhasil")
                st.rerun()
            else:
                st.markdown('<div class="error">❌ Password salah</div>', unsafe_allow_html=True)

    st.markdown("""
    <style>

    /* LABEL Username & Password */
    div[data-testid="stTextInput"] label {
        color: white !important;
        font-weight: 600;
    }

    /* Input biar tetap putih */
    div[data-testid="stTextInput"] input {
        background-color: #f9fafb !important;
        color: black !important;
    }

    /* Password field */
    div[data-testid="stTextInput"] + div label {
        color: white !important;
    }

    </style>
    """, unsafe_allow_html=True)

    st.stop()

# =========================
# GOOGLE DRIVE AUTH VIA STREAMLIT SECRETS 
# =========================
from oauth2client.service_account import ServiceAccountCredentials

service_account_info = st.secrets["gcp_service_account"]

credentials = ServiceAccountCredentials.from_json_keyfile_dict(
    service_account_info,
    scopes=["https://www.googleapis.com/auth/drive"]
)

# GoogleAuth + Drive
gauth = GoogleAuth()
gauth.credentials = credentials
drive = GoogleDrive(gauth)

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Project Monitoring", layout="wide")

col1, col2 = st.columns([10,1])

with col2:
    if st.button("🚪Logout"):
        st.session_state.logged_in = False
        st.rerun()

# =========================
# HEADER (NAVY + LOGO BLENDED)
# =========================
st.markdown(f"""
<style>
.header {{
    background: linear-gradient(90deg, #0b2c5a, #134e96);
    padding: 20px 30px;
    border-radius: 12px;
    margin-bottom: 25px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
}}

.header-text {{
    color: white;
}}

.header-title {{
    font-size: 32px;
    font-weight: 700;
    margin: 0;
}}

.header-subtitle {{
    font-size: 14px;
    color: #cbd5e1;
    margin-top: 5px;
}}

.header-logo img {{
    height: 65px;
}}

.tab-container {{
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
}}

.tab-button {{
    flex: 1;
}}

.tab-button button {{
    width: 100%;
    height: 60px;
    border-radius: 12px;
    border: none;
    background-color: #e5e7eb;
    font-size: 16px;
    font-weight: 600;
    color: #374151;
    transition: all 0.2s ease;
}}

.tab-button button:hover {{
    background-color: #d1d5db;
    color: black;
}}
</style>

<div class="header">
    <div class="header-text">
        <div class="header-title">
            Monitoring Program Prioritas Pelindo Group
        </div>
        <div class="header-subtitle">
            Created by Group MNEV
        </div>
    </div>
    <div class="header-logo">
        <img src="data:image/png;base64,{logo_base64}">
    </div>
</div>
""", unsafe_allow_html=True)

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
    if pd.isna(start) or pd.isna(end):
        return 0
    if start > end:
        return 0
    return np.busday_count(start.date(), end.date())

# =========================
# STATUS ENGINE (FIXED CORE)
# =========================

def get_status(progress, baseline, start_date, finish_date, global_target_date):

    start_date = pd.to_datetime(start_date, errors='coerce')
    finish_date = pd.to_datetime(finish_date, errors='coerce')
    global_target_date = pd.to_datetime(global_target_date, errors='coerce')

    if pd.isna(global_target_date):
        global_target_date = pd.Timestamp.today()

    # COMPLETE
    if progress >= 100:
        return "Complete"

    # NOT STARTED (before start date)
    if progress == 0:
        if pd.notna(start_date) and global_target_date < start_date:
            return "Not Started"
        if pd.notna(start_date) and global_target_date >= start_date:
            return "Late"
        return "Not Started"

    # LATE (OVERDUE FINISH)
    if pd.notna(finish_date) and global_target_date > finish_date:
        return "Late"

    # PERFORMANCE LOGIC
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

# =========================
# FIXED CALL PATTERN IMPORTANT
# =========================

def kpi_box(title, progress, baseline, start=None, finish=None, target=None):
    delta = progress - baseline
    status = get_status(progress, baseline, start, finish, target)
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


# =========================
# LOAD DATA
# =========================

def load_and_process(file_name, file_id, global_target_date):

    downloaded = drive.CreateFile({'id': file_id})
    downloaded.GetContentFile(file_name)

    df = pd.read_excel(file_name, header=8)
    df = df.iloc[:, 1:]

    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={
        'outline number': 'Outline number',
        'task name': 'Name',
        'name': 'Name',
        '% complete': '% complete',
        'start': 'Start',
        'finish': 'Finish',
        'duration': 'Duration',
        'bucket': 'Entitas',
        'progress':'Progress',
        'issues':'Issues'
    })

    for col in ['Progress', 'Issues']:
        if col not in df.columns:
            df[col] = ""

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

    df['duration_todate'] = df['Start'].apply(lambda x: networkdays(x, global_target_date))

    df['Baseline'] = (
        df['duration_todate'] /
        df['Duration'].replace(0,1)
    ).clip(upper=1) * 100

    df['is_parent'] = df['Outline number'].apply(
        lambda x: any(df['Outline number'].str.startswith(x + '.'))
    )

    leaf = df[~df['is_parent']].copy()

    return df, leaf


# =========================
# AGGREGATE
# =========================

def aggregate(leaf, code):
    subset = leaf[leaf['Outline number'].str.startswith(code)]

    if subset.empty:
        return 0, 0

    total_dur = subset['Duration'].replace(0,1).sum()

    progress = (subset['% complete'] * subset['Duration']).sum() / total_dur
    baseline = (subset['Baseline'] * subset['Duration']).sum() / total_dur

    return progress, baseline

# =========================
# GLOBAL TARGET DATE (PALING ATAS)
# =========================
st.markdown("## 📅 Date")

global_target_date = st.date_input(
    "Pilih Target Tanggal",
    value=date.today(),
    key="global_target_date"
)

global_target_date = datetime.combine(global_target_date, datetime.min.time())


# =========================
# TAB STATE
# =========================
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "tab1"

# =========================
# TAB BUTTONS (BOX STYLE)
# =========================
col1, col2 = st.columns(2)

with col1:
    if st.button("🌍 Overall Ekosistem", key="tab1_btn", use_container_width=True):
        st.session_state.active_tab = "tab1"

with col2:
    if st.button("📊 Detail Monitoring", key="tab2_btn", use_container_width=True):
        st.session_state.active_tab = "tab2"


# =========================
# ACTIVE TAB STYLE (IMPORTANT 🔥)
# =========================
active_index = 1 if st.session_state.active_tab == "tab1" else 2

st.markdown(f"""
<style>
.tab-container div[data-testid="stHorizontalBlock"] > div:nth-child({active_index}) button {{
    background-color: #0b2c5a !important;
    color: white !important;
    font-weight: 700;
    border: 2px solid #0b2c5a;
}}
</style>
""", unsafe_allow_html=True)

#tab1, tab2 = st.tabs(["🌍 Overall Ekosistem", "📊 Detail Monitoring"])

# =========================
# TAB 1
# =========================

if st.session_state.active_tab == "tab1":

    st.subheader("🌍 Overall Ekosistem")

    st.info("Halaman ini memberikan gambaran menyeluruh terkait progres program prioritas Pelindo Group, baik secara total maupun per program")

    ecosystem_results = []

    for file_name, file_id in excel_files.items():
        df, leaf = load_and_process(file_name, file_id, global_target_date)

        progress, baseline = aggregate(leaf, "1")

        ecosystem_results.append({
            "name": file_name.replace(".xlsx",""),
            "progress": progress,
            "baseline": baseline,
            "duration": leaf['Duration'].replace(0,1).sum()
        })

    df_eco = pd.DataFrame(ecosystem_results)

    total_weight = df_eco['duration'].sum()

    global_progress = (df_eco['progress'] * df_eco['duration']).sum() / total_weight
    global_baseline = (df_eco['baseline'] * df_eco['duration']).sum() / total_weight

    kpi_box("Overall Project Global", global_progress, global_baseline, None, None, global_target_date)

    st.markdown("## 📊 Progress per Ekosistem")

    cols = st.columns(3)
    for i, row in df_eco.iterrows():
        with cols[i % 3]:
            kpi_box(row['name'], row['progress'], row['baseline'], None, None, global_target_date)



if st.session_state.active_tab == "tab2":

    st.subheader("📊 Detail Monitoring")

    st.info("Detail monitoring untuk masing masing program prioritas")

    selected_file = st.selectbox(
        "Pilih Ekosistem / File",
        list(excel_files.keys()),
        key="file_selector_tab2"
    )

    if selected_file:

        # =========================
        # LOAD DATA
        # =========================
        df, leaf = load_and_process(
            selected_file,
            excel_files[selected_file],
            global_target_date
        )

        # =========================
        # CLEAN WBS FORMAT
        # =========================
        df['Outline number'] = df['Outline number'].astype(str).str.strip()
        leaf['Outline number'] = leaf['Outline number'].astype(str).str.strip()

        # =========================
        # KPI GLOBAL
        # =========================
        overall_p, overall_b = aggregate(leaf, "1")
        kpi_box("Overall Project", overall_p, overall_b, None, None, global_target_date)

        # =========================
        # LEVEL 2 KPI
        # =========================
        st.markdown("## 📊 Progress per Ekosistem")

        level2 = df[df['Outline number'].str.match(r'^\d+\.\d+$')]

        cols = st.columns(3)

        for i, (_, row) in enumerate(level2.iterrows()):
            p, b = aggregate(leaf, row['Outline number'])
            with cols[i % 3]:
                kpi_box(row['Name'], p, b, None, None, global_target_date)

        # =========================
        # SELECT OBJECT
        # =========================
        st.divider()

        st.markdown("### 🔍 Monitoring per Objek Ekosistem")

        selected_l2 = st.selectbox(
            "Pilih Objek Ekosistem",
            level2['Name'],
            key="l2_selector_tab2"
        )

        selected_code = level2.loc[
            level2['Name'] == selected_l2,
            'Outline number'
        ].values[0]

        # =========================
        # FIXED TREE FILTER (ALL DESCENDANTS + PARENT)
        # =========================
        def get_tree(df, code):
            pattern = r'^' + re.escape(code) + r'(\.|$)'
            return df[df['Outline number'].astype(str).str.match(pattern)].copy()

        sub_tasks = get_tree(leaf, selected_code)

        # include parent row from df (if exists)
        parent_row = df[df['Outline number'] == selected_code]
        sub_tasks = pd.concat([parent_row, sub_tasks], ignore_index=True)

        # sort WBS
        sub_tasks = sub_tasks.sort_values('Outline number')

        # =========================
        # STATUS CALCULATION
        # =========================
        sub_tasks['Status'] = sub_tasks.apply(
            lambda x: get_status(
                x['% complete'],
                x['Baseline'],
                x['Start'],
                x['Finish'],
                global_target_date
            ),
            axis=1
        )

        # =========================
        # FILTER STATE
        # =========================
        if "status_filter" not in st.session_state:
            st.session_state.status_filter = "ALL"

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            if st.button(f"✅ Complete ({len(sub_tasks[sub_tasks['Status']=='Complete'])})", key="tab2_c"):
                st.session_state.status_filter = "Complete"

        with col2:
            if st.button(f"🟢 On Progress ({len(sub_tasks[sub_tasks['Status']=='On Progress'])})", key="tab2_p"):
                st.session_state.status_filter = "On Progress"

        with col3:
            if st.button(f"🟠 Concern ({len(sub_tasks[sub_tasks['Status']=='Concern'])})", key="tab2_con"):
                st.session_state.status_filter = "Concern"

        with col4:
            if st.button(f"🔴 Late ({len(sub_tasks[sub_tasks['Status']=='Late'])})", key="tab2_l"):
                st.session_state.status_filter = "Late"

        if st.button("🔄 Show All", key="tab2_all"):
            st.session_state.status_filter = "ALL"

        # =========================
        # APPLY FILTER
        # =========================
        filtered = sub_tasks.copy()

        if st.session_state.status_filter != "ALL":
            filtered = filtered[filtered['Status'] == st.session_state.status_filter]

        # =========================
        # TABLE DISPLAY
        # =========================
        st.subheader("📋 Detail Task Summary")

        if filtered.empty:
            st.warning("Tidak ada data")
        else:

            filtered['Level'] = filtered['Outline number'].astype(str).str.count(r'\.')

            filtered['Name WBS'] = filtered.apply(
                lambda r: "   " * max(r['Level'] - 1, 0) + "▸ " + str(r['Name']),
                axis=1
            )

            filtered['Start'] = pd.to_datetime(filtered['Start'], errors='coerce').dt.strftime('%d/%m/%Y')
            filtered['Finish'] = pd.to_datetime(filtered['Finish'], errors='coerce').dt.strftime('%d/%m/%Y')

            filtered['Progress (%)'] = filtered['% complete'].map(lambda x: f"{x:.1f}%")
            filtered['Baseline (%)'] = filtered['Baseline'].map(lambda x: f"{x:.1f}%")

            st.dataframe(
                filtered[
                    [
                        'Outline number',
                        'Name WBS',
                        'Entitas',
                        'Start',
                        'Finish',
                        'Progress (%)',
                        'Baseline (%)',
                        'Status',
                        'Issues',
                        'Progress'
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )
        # =========================
        # DELAY RANKING (ALL LATE TASK)
        # =========================
        st.markdown("## 🔴 Rank Late Task List per Objek")

        late_tasks = sub_tasks[sub_tasks['Status'] == "Late"].copy()

        if late_tasks.empty:
            st.success("Tidak ada task yang mengalami keterlambatan 👍")
        else:
            today = pd.to_datetime(global_target_date)

            late_tasks['Delay (Days)'] = late_tasks['Finish'].apply(
                lambda x: (today - x).days if pd.notna(x) and today > x else 0
            )

            late_tasks = late_tasks.sort_values(by='Delay (Days)', ascending=False)

            late_tasks['Start'] = pd.to_datetime(late_tasks['Start'], errors='coerce').dt.strftime('%d/%m/%Y')
            late_tasks['Finish'] = pd.to_datetime(late_tasks['Finish'], errors='coerce').dt.strftime('%d/%m/%Y')

            late_tasks = late_tasks.reset_index(drop=True)
            late_tasks['Rank'] = late_tasks.index + 1

            st.dataframe(
                late_tasks[
                    [
                        'Rank',
                        'Outline number',
                        'Name',
                        'Entitas',
                        'Finish',
                        'Delay (Days)',
                        'Issues'
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

        # =========================
        # ENTITAS SECTION (FULL TREE FIX)
        # =========================
        st.divider()
        st.markdown("## 🏢 Monitoring Berdasarkan Entitas")

        
        all_entitas = sorted(leaf['Entitas'].dropna().unique())

        selected_entitas = st.selectbox(
            "Pilih Entitas",
            all_entitas,
            key="entitas_tab2"
        )

        # =========================
        # AMBIL LEAF ENTITAS
        # =========================
        entitas_leaf = leaf[leaf['Entitas'] == selected_entitas].copy()

        # =========================
        # AMBIL SEMUA PARENT + TREE
        # =========================
        def get_full_tree_from_leaf(df_all, leaf_subset):
            all_codes = set()

            for code in leaf_subset['Outline number']:
                parts = str(code).split('.')

                # generate parent chain
                for i in range(1, len(parts) + 1):
                    parent_code = ".".join(parts[:i])
                    all_codes.add(parent_code)

            return df_all[df_all['Outline number'].isin(all_codes)].copy()


        entitas_tasks = get_full_tree_from_leaf(df, entitas_leaf)

        # =========================
        # SORT BIAR RAPI
        # =========================
        entitas_tasks = entitas_tasks.sort_values('Outline number')

        # =========================
        # KPI HITUNG DARI LEAF (BENAR)
        # =========================
        total_dur = entitas_leaf['Duration'].replace(0, 1).sum()

        entitas_progress = (
            (entitas_leaf['% complete'] * entitas_leaf['Duration']).sum()
            / total_dur
        )

        entitas_baseline = (
            (entitas_leaf['Baseline'] * entitas_leaf['Duration']).sum()
            / total_dur
        )

        kpi_box(
            f"Performance - {selected_entitas}",
            entitas_progress,
            entitas_baseline,
            None,
            None,
            global_target_date
        )

        # =========================
        # STATUS (APPLY KE FULL TREE)
        # =========================
        entitas_tasks['Status'] = entitas_tasks.apply(
            lambda x: get_status(
                x['% complete'],
                x['Baseline'],
                x['Start'],
                x['Finish'],
                global_target_date
            ),
            axis=1
        )

        # =========================
        # FILTER STATUS
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
        # DISPLAY TABLE
        # =========================
        st.subheader("📋 Detail Task per Entitas")

        if filtered_entitas.empty:
            st.warning("Tidak ada data")
        else:

            filtered_entitas['Level'] = filtered_entitas['Outline number'].str.count(r'\.')

            filtered_entitas['Name WBS'] = filtered_entitas.apply(
                lambda r: "   " * max(r['Level'] - 1, 0) + "▸ " + str(r['Name']),
                axis=1
            )

            filtered_entitas['Start'] = pd.to_datetime(filtered_entitas['Start'], errors='coerce').dt.strftime('%d/%m/%Y')
            filtered_entitas['Finish'] = pd.to_datetime(filtered_entitas['Finish'], errors='coerce').dt.strftime('%d/%m/%Y')

            filtered_entitas['Progress (%)'] = filtered_entitas['% complete'].map(lambda x: f"{x:.1f}%")
            filtered_entitas['Baseline (%)'] = filtered_entitas['Baseline'].map(lambda x: f"{x:.1f}%")

            st.dataframe(
                filtered_entitas[
                    [
                        'Outline number',
                        'Name WBS',
                        'Entitas',
                        'Start',
                        'Finish',
                        'Progress (%)',
                        'Baseline (%)',
                        'Status',
                        'Issues',
                        'Progress'
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )
        # =========================
        # DELAY RANKING PER ENTITAS
        # =========================
        st.markdown("## 🔴 Ranking Late Task per Entitas")

        late_entitas = entitas_tasks[entitas_tasks['Status'] == "Late"].copy()

        if late_entitas.empty:
            st.success("Tidak ada keterlambatan pada entitas ini 👍")
        else:
            today = pd.to_datetime(global_target_date)

            late_entitas['Delay (Days)'] = late_entitas['Finish'].apply(
                lambda x: (today - x).days if pd.notna(x) and today > x else 0
            )

            late_entitas = late_entitas.sort_values(by='Delay (Days)', ascending=False)

            late_entitas['Start'] = pd.to_datetime(late_entitas['Start'], errors='coerce').dt.strftime('%d/%m/%Y')
            late_entitas['Finish'] = pd.to_datetime(late_entitas['Finish'], errors='coerce').dt.strftime('%d/%m/%Y')

            late_entitas = late_entitas.reset_index(drop=True)
            late_entitas['Rank'] = late_entitas.index + 1

            st.dataframe(
                late_entitas[
                    [
                        'Rank',
                        'Outline number',
                        'Name',
                        'Entitas',
                        'Delay (Days)',
                        'Issues',
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )
