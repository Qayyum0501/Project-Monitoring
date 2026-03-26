import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from streamlit_calendar import calendar

# =========================
# TAB BOOKING
# =========================
st.set_page_config(page_title="Project Monitoring", layout="wide")
st.title("🚧 Monitoring Program Prioritas")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=8)
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

    required_cols = ['Outline number', 'Name', 'Start', 'Finish', 'Duration', '% complete']
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

    # =========================
    # TARGET DATE
    # =========================
    target_date = st.date_input("Pilih Target Tanggal", datetime(2026,12,8))
    target_date = datetime.combine(target_date, datetime.min.time())

    # =========================
    # NETWORK DAYS
    # =========================
    def networkdays(start, end):
        if pd.isna(start):
            return 0
        if start > end:
            return 0
        return np.busday_count(start.date(), end.date())
    
    df['duration_todate'] = df['Start'].apply(lambda x: networkdays(x, target_date))

    # =========================
    # CLEAN DURATION
    # =========================
    df['Duration'] = (
        df['Duration'].astype(str)
        .str.replace(' days','', regex=False)
        .str.replace(' day','', regex=False)
        .astype(float)
        .fillna(0)
        .round()
        .astype(int)
    )

    # =========================
    # BASELINE
    # =========================
    df['Baseline'] = (df['duration_todate'] / df['Duration'].replace(0,1)).clip(upper=1)
    df['Baseline'] = (df['Baseline'] * 100).round(1)

    # =========================
    # WEIGHTED BASELINE
    # =========================
    def weighted_progress(parent):
        children = df[
            df['Outline number'].str.startswith(parent, na=False) &
            (df['Outline number'] != parent)
        ]
        if children.empty:
            val = df.loc[df['Outline number'] == parent, 'Baseline']
            return val.values[0] if len(val) > 0 else 0
        total_duration = children['Duration'].replace(0,1).sum()
        return (children['Baseline'] * children['Duration']).sum() / total_duration

    df['Baseline_progress'] = df['Outline number'].apply(weighted_progress)

    # =========================
    # PROGRESS
    # =========================
    df['% complete'] = (df['% complete'].fillna(0) * 100).round(1)

    # =========================
    # DELAY (PAKAI TARGET DATE)
    # =========================
    def calc_delay(row):
        if pd.isna(row['Finish']):
            return 0
        if row['Finish'] < target_date and row['% complete'] < 100:
            return np.busday_count(row['Finish'].date(), target_date.date())
        else:
            return 0

    df['Delay (days)'] = df.apply(calc_delay, axis=1)

    # =========================
    # STATUS
    # =========================
    df['Status'] = np.select(
        [
            df['% complete'] == 0,
            df['% complete'] >= 100,
            (df['Finish'] < target_date) & (df['% complete'] < 100),
            df['% complete'] >= 0.9 * df['Baseline_progress'],
            df['% complete'] >= 0.75 * df['Baseline_progress']
        ],
        ['Not Started','Complete','Late','On Progress','Concern'],
        default='Late'
    )

    # =========================
    # KPI FUNCTION
    # =========================
    def calc_kpi(data):
        total_duration = data['Duration'].replace(0,1).sum()
        progress = (data['% complete'] * data['Duration']).sum() / total_duration
        baseline = (data['Baseline'] * data['Duration']).sum() / total_duration
        delta = progress - baseline
        return round(progress,1), round(baseline,1), round(delta,1)

    def kpi_box(title, progress, baseline, delta):
        if delta >= 0:
            color = "green"
        elif delta >= -5:
            color = "orange"
        else:
            color = "red"
        st.markdown(f"""
        <div style="border-radius:15px;padding:20px;text-align:center;background:#f9f9f9;">
            <div style="font-size:18px;font-weight:bold;">{title}</div>
            <div style="font-size:40px;font-weight:bold;color:{color};">
                {progress:.1f}%
            </div>
            <div style="font-size:14px;color:gray;">
                Baseline: {baseline:.1f}%
            </div>
            <div style="font-size:18px;color:{color};">
                Δ {delta:+.1f}%
            </div>
        </div>
        """, unsafe_allow_html=True)

    # =========================
    # KPI PER SUB EKOSISTEM
    # =========================
    st.markdown("## 📊 Progress per Sub-Ekosistem")
    level2 = df[df['Outline number'].str.count(r'\.') == 1]

    if level2.empty:
        st.warning("⚠️ Tidak ada data level 2")
        st.stop()

    cols = st.columns(3)
    for i, (_, row) in enumerate(level2.iterrows()):
        code = row['Outline number']
        sub_data = df[df['Outline number'].str.startswith(code, na=False)]
        p, b, d = calc_kpi(sub_data)
        with cols[i % 3]:
            kpi_box(row['Name'], p, b, d)

    # =========================
    # SELECT SUB EKOSISTEM
    # =========================
    st.divider()
    selected_l2 = st.selectbox("Pilih Objek Ekosistem", level2['Name'])
    selected_row = level2[level2['Name'] == selected_l2]
    sub_tasks = pd.DataFrame()
    if not selected_row.empty:
        selected_l2_code = selected_row['Outline number'].values[0]
        sub_tasks = df[df['Outline number'].str.startswith(selected_l2_code, na=False)].copy()

    # =========================
    # HITUNG KPI
    # =========================
    total_complete = len(sub_tasks[sub_tasks['Status']=="Complete"])
    total_progress = len(sub_tasks[sub_tasks['Status']=="On Progress"])
    total_concern = len(sub_tasks[sub_tasks['Status']=="Concern"])
    total_late = len(sub_tasks[sub_tasks['Status']=="Late"])

    # =========================
    # FILTER BUTTON
    # =========================
    if "status_filter" not in st.session_state:
        st.session_state.status_filter = "ALL"

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button(f"✅ Complete ({total_complete})"):
            st.session_state.status_filter = "Complete"
    with col2:
        if st.button(f"🟢 On Progress ({total_progress})"):
            st.session_state.status_filter = "On Progress"
    with col3:
        if st.button(f"🟠 Concern ({total_concern})"):
            st.session_state.status_filter = "Concern"
    with col4:
        if st.button(f"🔴 Late ({total_late})"):
            st.session_state.status_filter = "Late"
    if st.button("🔄 Show All"):
        st.session_state.status_filter = "ALL"

    st.markdown(f"**Filter aktif:** `{st.session_state.status_filter}`")
    if st.session_state.status_filter != "ALL":
        sub_tasks = sub_tasks[sub_tasks['Status'] == st.session_state.status_filter]

    # =========================
    # DISPLAY TABLE
    # =========================
    st.subheader("📋 Detail Task Summary")
    if sub_tasks.empty:
        st.warning("⚠️ Tidak ada task")
    else:
        sub_tasks['Level'] = sub_tasks['Outline number'].apply(lambda x: x.count('.'))
        sub_tasks['Name WBS'] = sub_tasks.apply(
            lambda row: "   " * (row['Level'] - 1) + "▸ " + row['Name'], axis=1
        )
        sub_tasks = sub_tasks.sort_values('Outline number')
        display_df = sub_tasks.copy()
        display_df['Start'] = display_df['Start'].dt.strftime('%d/%m/%Y')
        display_df['Finish'] = display_df['Finish'].dt.strftime('%d/%m/%Y')
        display_df['Progress (%)'] = display_df['% complete'].map(lambda x: f"{x:.1f}%")
        display_df['Baseline (%)'] = display_df['Baseline'].map(lambda x: f"{x:.1f}%")
        display_df = display_df[[
            'Outline number','Name WBS','Entitas','Start','Finish','Progress (%)',
            'Baseline (%)','Delay (days)','Status'
        ]]
        st.dataframe(display_df, use_container_width=True)

    # =========================
    # TOP DELAY
    # =========================
    st.subheader("🏆 Top 10 Task Paling Telat")
    children_tasks = df[df['Outline number'].str.count(r'\.') >= 3]
    top_delay = children_tasks[children_tasks['Delay (days)'] > 0]
    if top_delay.empty:
        st.info("Tidak ada task terlambat 🎉")
    else:
        top_delay = top_delay.sort_values('Delay (days)', ascending=False).head(10)
        top_delay['Finish'] = top_delay['Finish'].dt.strftime('%d/%m/%Y')
        st.dataframe(top_delay[['Outline number','Name','Entitas','Finish','Delay (days)','Status']],
                     use_container_width=True)

    # =========================
    # HEAT MAP
    # =========================
    st.markdown("## 🔥 Heatmap Entitas (Klik untuk Detail)")
    late_summary = (
        df[df['Status']=='Late']
        .groupby('Entitas').size()
        .reset_index(name='Total Late')
        .sort_values('Total Late', ascending=False)
    )
    cols = st.columns(4)
    for i, row in late_summary.iterrows():
        entitas = row['Entitas']
        total = row['Total Late']
        if total > 20:
            color = "#ff4d4d"
        elif total > 10:
            color = "#ffa64d"
        else:
            color = "#66cc66"
        with cols[i % 4]:
            if st.button(f"{entitas}\n{total} Late", key=f"btn_{entitas}"):
                st.session_state.selected_entitas = entitas
            st.markdown(f"""
            <div style="
                background-color:{color};
                padding:20px;
                border-radius:15px;
                text-align:center;
                color:white;
                font-weight:bold;
            ">
                {entitas}<br>{total} Late
            </div>
            """, unsafe_allow_html=True)

    if "selected_entitas" in st.session_state:
        selected = st.session_state.selected_entitas
        st.markdown(f"## 📋 Detail Task Late - {selected}")
        detail = df[(df['Entitas']==selected) & (df['Status']=='Late')].copy()
        if detail.empty:
            st.info("Tidak ada task Late")
        else:
            detail = detail.sort_values('Delay (days)', ascending=False)
            detail['Finish'] = detail['Finish'].dt.strftime('%d/%m/%Y')
            st.dataframe(detail[['Outline number','Name','Finish','Delay (days)','Status']],
                         use_container_width=True)
