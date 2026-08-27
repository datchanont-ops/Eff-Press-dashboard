import pandas as pd
import streamlit as st
import datetime
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import streamlit.components.v1 as components
import os
import io
import json
import base64
import requests

# --- 1. ตั้งค่าหน้าจอ Streamlit (ต้องอยู่บนสุดเสมอ) ---
st.set_page_config(page_title="Production Executive Dashboard", page_icon="🏭", layout="wide")

# ==========================================
# 🔗 GitHub Persistence Layer
# ==========================================
try:
    GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
    GITHUB_REPO = st.secrets["GITHUB_REPO"]
    GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")
    GITHUB_DATA_DIR = st.secrets.get("GITHUB_DATA_DIR", "data")
    GITHUB_ENABLED = True
except Exception:
    GITHUB_ENABLED = False

GITHUB_API = "https://api.github.com"

def gh_headers():
    return {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}

def gh_get_file(remote_path):
    if not GITHUB_ENABLED: return None, None
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}?ref={GITHUB_BRANCH}"
    try:
        r = requests.get(url, headers=gh_headers(), timeout=15)
        if r.status_code == 200:
            data = r.json()
            return base64.b64decode(data["content"]), data["sha"]
    except: pass
    return None, None

def gh_put_file(remote_path, content_bytes, message):
    if not GITHUB_ENABLED: return False
    _, sha = gh_get_file(remote_path)
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}"
    payload = {"message": message, "content": base64.b64encode(content_bytes).decode("utf-8"), "branch": GITHUB_BRANCH}
    if sha: payload["sha"] = sha
    try:
        r = requests.put(url, headers=gh_headers(), json=payload, timeout=15)
        return r.status_code in (200, 201)
    except: return False

# ==========================================
# 🎨 CSS Modern Corporate Styling
# ==========================================
st.markdown("""
<style>
    /* พื้นหลังโดยรวมและฟอนต์ */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 95% !important;
    }
    
    /* สไตล์สำหรับ KPI Cards (Metric) */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 1.2rem 1.5rem;
        border-radius: 0.75rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        border-left: 5px solid #1e293b; 
        transition: transform 0.2s ease;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    
    div[data-testid="column"]:nth-child(1) div[data-testid="metric-container"] { border-left-color: #3b82f6; } 
    div[data-testid="column"]:nth-child(2) div[data-testid="metric-container"] { border-left-color: #8b5cf6; } 
    div[data-testid="column"]:nth-child(3) div[data-testid="metric-container"] { border-left-color: #10b981; } 
    div[data-testid="column"]:nth-child(4) div[data-testid="metric-container"] { border-left-color: #f59e0b; } 
    div[data-testid="column"]:nth-child(5) div[data-testid="metric-container"] { border-left-color: #64748b; } 

    div[data-testid="metric-container"] > div {
        color: #64748b;
        font-weight: 600;
        font-size: 0.95rem;
    }
    div[data-testid="metric-container"] label {
        font-size: 1.8rem !important;
        color: #0f172a;
        font-weight: 700;
    }

    @media print {
        .stPopover, .stExpander, .stDownloadButton, header, [data-testid="stSidebar"] { display: none !important; }
        .block-container { max-width: 100% !important; padding: 0 !important; }
    }
    
    h1, h2, h3 { color: #0f172a; font-weight: 700; }
    hr { margin-top: 1.5rem; margin-bottom: 1.5rem; border-color: #e2e8f0; }
</style>
""", unsafe_allow_html=True)

# --- 2. ฟังก์ชันโหลดไฟล์ ---
@st.cache_data
def load_target_data():
    try:
        df_target = pd.read_excel('target.xlsx')
        if 'Material' in df_target.columns: df_target = df_target.rename(columns={'Material': 'Part', 'cap/day': 'เป้าต่อวัน(3กะ)'})
        else: df_target.columns = ['Part', 'เป้าต่อวัน(3กะ)']
        df_target.columns = df_target.columns.str.strip()
        return df_target, None
    except Exception: return None, "❌ ไม่พบไฟล์ `target.xlsx` ในระบบ"

@st.cache_data
def load_daily_data(file, df_target):
    try:
        xls = pd.ExcelFile(file)
        df_pd = pd.read_excel(xls, 'pd') if 'pd' in xls.sheet_names else pd.read_excel(xls, 0)
        df = df_pd[['Material', 'Document Header Text', 'Qty in Un. of Entry', 'Posting Date', 'Entry Date', 'Time of Entry']].copy()

        def extract_machine(text):
            if pd.isna(text): return None
            parts = str(text).split('/')
            return parts[1] if len(parts) >= 2 and str(text).startswith('1/') else None

        df['Machine'] = df['Document Header Text'].apply(extract_machine)
        df_filtered = df.dropna(subset=['Machine']).copy().rename(columns={'Posting Date': 'วันที่ผลิต', 'Material': 'Part', 'Qty in Un. of Entry': 'actual_qty'})
        df_filtered['วันที่ผลิต'] = pd.to_datetime(df_filtered['วันที่ผลิต']).dt.date
        df_filtered['Entry Date'] = pd.to_datetime(df_filtered['Entry Date']).dt.date

        df_filtered = df_filtered.sort_values(by=['Machine', 'วันที่ผลิต', 'Entry Date', 'Time of Entry'])
        df_filtered['Part_ก่อนหน้า'] = df_filtered.groupby('Machine')['Part'].shift(1)
        df_filtered['Is_Setup'] = (df_filtered['Part'] != df_filtered['Part_ก่อนหน้า']) & (df_filtered['Part_ก่อนหน้า'].notna())

        df_grouped = df_filtered.groupby(['วันที่ผลิต', 'Machine', 'Part'], as_index=False).agg(actual_qty=('actual_qty', 'sum'), Setup_Count=('Is_Setup', 'sum'))
        df_final = pd.merge(df_grouped, df_target, on='Part', how='left')
        df_final['เป้าต่อวัน(3กะ)'] = df_final['เป้าต่อวัน(3กะ)'].fillna(0)
        return df_final, None
    except Exception as e: return None, f"Error: {e}"

# ==========================================
# --- Header / Title ---
# ==========================================
st.markdown("<h1>📊 Production Executive Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<p style='color: #64748b; font-size: 1.1rem; margin-top: -10px;'>ระบบวิเคราะห์และติดตามประสิทธิภาพการผลิตประจำวัน (Daily Efficiency Tracking)</p>", unsafe_allow_html=True)

df_target, target_error = load_target_data()
if target_error: st.error(target_error); st.stop()

# ==========================================
# --- Sidebar (Settings & Upload) ---
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2823/2823512.png", width=60)
    st.markdown("### 📥 1. จัดการข้อมูล (Data)")
    
    template_file = "Template.xlsx"
    if os.path.exists(template_file):
        with open(template_file, "rb") as f:
            st.download_button("🔽 โหลด Template", data=f, file_name=template_file, use_container_width=True)
            
    uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์รายวัน (Excel)", type=["xlsx", "xls"], label_visibility="collapsed")

# --- ระบบซิงค์ข้อมูล Local/GitHub ---
default_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'default_data.xlsx')
settings_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'eff_settings.json')

if GITHUB_ENABLED and not st.session_state.get("eff_github_synced"):
    with st.spinner("🔄 Syncing from GitHub..."):
        c1, _ = gh_get_file(f"{GITHUB_DATA_DIR}/default_data.xlsx")
        if c1: open(default_file_path, "wb").write(c1)
        c2, _ = gh_get_file(f"{GITHUB_DATA_DIR}/eff_settings.json")
        if c2: open(settings_file, "wb").write(c2)
    st.session_state["eff_github_synced"] = True

if uploaded_file:
    if st.sidebar.button("💾 บันทึกไฟล์นี้เป็นค่าเริ่มต้น", use_container_width=True):
        file_bytes = bytes(uploaded_file.getbuffer())
        open(default_file_path, "wb").write(file_bytes)
        if GITHUB_ENABLED: gh_put_file(f"{GITHUB_DATA_DIR}/default_data.xlsx", file_bytes, f"Auto-save: {uploaded_file.name}")
        st.sidebar.success("✅ บันทึกสำเร็จ!")

data_source = uploaded_file if uploaded_file else (default_file_path if os.path.exists(default_file_path) else None)
if not data_source: st.info("👈 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้น"); st.stop()

df, error = load_daily_data(data_source, df_target)
if error: st.error(error); st.stop()

min_date, max_date = df['วันที่ผลิต'].min(), df['วันที่ผลิต'].max()
st.sidebar.caption(f"📅 ช่วงข้อมูล: {min_date.strftime('%d/%m/%y')} - {max_date.strftime('%d/%m/%y')}")

# --- โหลด Settings ---
saved_settings = {}
if os.path.exists(settings_file):
    try: saved_settings = json.load(open(settings_file, 'r', encoding='utf-8'))
    except: pass

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ 2. ตั้งค่าพารามิเตอร์")

init_oee = int(saved_settings.get('oee_val', 90))
init_setup = float(saved_settings.get('setup_hours', 4.0))

oee_val = st.sidebar.number_input("O.E.E. Target (%)", min_value=1, max_value=100, value=init_oee, step=1)
setup_hours = st.sidebar.number_input("เวลา Setup (ชม./ครั้ง)", min_value=0.0, max_value=24.0, value=init_setup, step=0.5)
oee_mult = oee_val / 100.0
setup_deduct = setup_hours / 24.0

st.sidebar.markdown("---")
st.sidebar.markdown("### 🛠️ 3. ปรับกะการทำงาน (Shift)")
# เปลี่ยนชื่อคีย์ให้ตรงกับแบบเดิมที่เซฟไว้ใน eff_settings.json เพื่อแก้ KeyError
shift_map = {"3 กะ (เป้า 100%)": 1.0, "2 กะ (เป้า 67%)": 0.67, "1.5 กะ (เป้า 50%)": 0.5}

# 3.1 Date
with st.sidebar.expander("📅 ปรับตามวัน (By Date)", expanded=False):
    sd_val = st.selectbox("เปลี่ยนทุกวัน:", ["(ใช้ค่าเดิม)"] + list(shift_map.keys()), key='def_d')
    u_dates = sorted(df['วันที่ผลิต'].unique())
    s_d_data = [sd_val]*len(u_dates) if sd_val != "(ใช้ค่าเดิม)" else [saved_settings.get('shift_date', {}).get(str(d), "3 กะ (เป้า 100%)") for d in u_dates]
    df_sd = st.data_editor(pd.DataFrame({'Date': u_dates, 'Shift': s_d_data}), hide_index=True, use_container_width=True)
    df['mult_D'] = df['วันที่ผลิต'].map({row['Date']: shift_map[row['Shift']] for _, row in df_sd.iterrows()})

# 3.2 Machine
with st.sidebar.expander("🚜 ปรับตามเครื่อง (By Machine)", expanded=False):
    sm_val = st.selectbox("เปลี่ยนทุกเครื่อง:", ["(ใช้ค่าเดิม)"] + list(shift_map.keys()), key='def_m')
    u_macs = sorted(df['Machine'].unique())
    s_m_data = [sm_val]*len(u_macs) if sm_val != "(ใช้ค่าเดิม)" else [saved_settings.get('shift_mac', {}).get(m, "3 กะ (เป้า 100%)") for m in u_macs]
    df_sm = st.data_editor(pd.DataFrame({'Machine': u_macs, 'Shift': s_m_data}), hide_index=True, use_container_width=True)
    df['mult_M'] = df['Machine'].map({row['Machine']: shift_map[row['Shift']] for _, row in df_sm.iterrows()})

# 3.3 Spec
with st.sidebar.expander("🎯 ปรับเฉพาะกิจ (Date+Machine)", expanded=False):
    saved_spec = saved_settings.get('shift_spec', [])
    df_sp = pd.DataFrame(saved_spec) if saved_spec else pd.DataFrame(columns=["Date", "Machine", "Shift"])
    if not df_sp.empty and 'Date' in df_sp: df_sp['Date'] = pd.to_datetime(df_sp['Date']).dt.date
    df_spec = st.data_editor(df_sp, num_rows="dynamic", column_config={"Date": st.column_config.DateColumn(format="DD/MM/YYYY"), "Shift": st.column_config.SelectboxColumn(options=list(shift_map.keys()))}, hide_index=True, use_container_width=True)

if st.sidebar.button("💾 บันทึกการตั้งค่าทั้งหมด", use_container_width=True):
    to_save = {
        'oee_val': oee_val, 'setup_hours': setup_hours,
        'shift_date': {str(r['Date']): r['Shift'] for _, r in df_sd.iterrows()},
        'shift_mac': {r['Machine']: r['Shift'] for _, r in df_sm.iterrows()},
        'shift_spec': [{'Date': str(r['Date']), 'Machine': r['Machine'], 'Shift': r['Shift']} for _, r in df_spec.dropna().iterrows()] if not df_spec.empty else []
    }
    j_str = json.dumps(to_save, ensure_ascii=False, indent=4)
    open(settings_file, 'w', encoding='utf-8').write(j_str)
    if GITHUB_ENABLED: gh_put_file(f"{GITHUB_DATA_DIR}/eff_settings.json", j_str.encode('utf-8'), "Save settings")
    st.sidebar.success("✅ บันทึกตั้งค่าแล้ว!")

# --- คำนวณ % Achieve สุทธิ ---
df['mult_Net'] = df[['mult_D', 'mult_M']].min(axis=1)
if not df_spec.empty:
    for _, r in df_spec.dropna().iterrows():
        df.loc[(df['วันที่ผลิต'] == r['Date']) & (df['Machine'] == r['Machine']), 'mult_Net'] = shift_map[r['Shift']]

df['เป้าหมายสุทธิ'] = ((df['เป้าต่อวัน(3กะ)'] * df['mult_Net'] * oee_mult) - ((df['เป้าต่อวัน(3กะ)'] * setup_deduct) * df['Setup_Count'])).clip(lower=0)
df['% Achieve'] = ((df['actual_qty'] / df['เป้าหมายสุทธิ']) * 100).clip(upper=100.0).round(2).fillna(0)

# ==========================================
# 📊 Section 1: Top KPI Metrics (Modern Style)
# ==========================================
total_act = df['actual_qty'].sum()
total_tgt_raw = df['เป้าต่อวัน(3กะ)'].sum()
total_tgt_net = df['เป้าหมายสุทธิ'].sum()
overall_eff = min((total_act / total_tgt_net * 100) if total_tgt_net > 0 else 0, 100.0)

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("📦 ยอดผลิตจริง (Actual)", f"{total_act:,.0f}", "Pcs")
col2.metric("🎯 เป้าดิบ 3 กะ (Raw Target)", f"{total_tgt_raw:,.0f}", "Pcs")
col3.metric("🎯 เป้าสุทธิ (Net Target)", f"{total_tgt_net:,.0f}", "หัก Setup/กะ/OEE")
col4.metric("📈 ประสิทธิภาพรวม (Achieve)", f"{overall_eff:.1f}%", "Overall KPI")
col5.metric("⚙️ O.E.E. Baseline", f"{oee_val}%", "ตั้งค่าระบบ")

st.markdown("<hr>", unsafe_allow_html=True)

# ==========================================
# 📈 Section 2: Main Trend Chart
# ==========================================
st.markdown("### 📉 แนวโน้มยอดผลิตเทียบเป้าหมาย (Daily Trend Analysis)")

daily = df.groupby('วันที่ผลิต').agg({'actual_qty': 'sum', 'เป้าหมายสุทธิ': 'sum'}).reset_index()
daily['% Achieve'] = ((daily['actual_qty'] / daily['เป้าหมายสุทธิ']) * 100).clip(upper=100.0).round(2).fillna(0)

fig_trend = make_subplots(specs=[[{"secondary_y": True}]])

# แท่งเป้าหมาย (สีอ่อน)
fig_trend.add_trace(go.Bar(x=daily['วันที่ผลิต'], y=daily['เป้าหมายสุทธิ'], name="Net Target", marker_color='#cbd5e1', opacity=0.7), secondary_y=False)
# แท่งผลิตจริง (สีน้ำเงิน)
fig_trend.add_trace(go.Bar(x=daily['วันที่ผลิต'], y=daily['actual_qty'], name="Actual Qty", marker_color='#3b82f6'), secondary_y=False)
# เส้นประสิทธิภาพ (สีแดงอมส้ม)
fig_trend.add_trace(go.Scatter(x=daily['วันที่ผลิต'], y=daily['% Achieve'], name="% Achieve", mode='lines+markers', line=dict(color='#ef4444', width=3), marker=dict(size=8, color='#ef4444', line=dict(width=2, color='white'))), secondary_y=True)

fig_trend.update_layout(
    barmode='group', hovermode="x unified",
    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
    margin=dict(l=0, r=0, t=10, b=0),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    font=dict(family="sans-serif", color="#475569")
)
fig_trend.update_yaxes(title_text="Quantity (Pcs)", showgrid=True, gridcolor='#f1f5f9', secondary_y=False)
fig_trend.update_yaxes(title_text="Achievement (%)", showgrid=False, range=[0, 115], secondary_y=True)

st.plotly_chart(fig_trend, use_container_width=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ==========================================
# 🏆 Section 3: Group Analysis & Top 5 Parts
# ==========================================
col_grp, col_top = st.columns([1.2, 1])

with col_grp:
    st.markdown("### 📊 ประสิทธิภาพแยกกลุ่มเครื่องจักร (By Group)")
    
    def ext_grp(mc):
        for g in ['INJ', 'INM', '510', 'VAC', '400T', '300T', '350T', 'PRESS']:
            if g in str(mc).upper(): return g
        return "OTHER"
        
    df['Group'] = df['Machine'].apply(ext_grp)
    grp = df.groupby('Group').agg({'actual_qty': 'sum', 'เป้าหมายสุทธิ': 'sum'}).reset_index()
    grp['% Achieve'] = ((grp['actual_qty'] / grp['เป้าหมายสุทธิ']) * 100).clip(upper=100.0).round(2).fillna(0)
    
    def prio(g):
        g = str(g).upper()
        if "INJ" in g: return 1
        if "PRESS" in g: return 2
        if "VAC" in g: return 3
        return 4
        
    grp['sort'] = grp['Group'].apply(prio)
    grp = grp.sort_values(by=['sort', 'Group'])
    
    # ใช้สีระดับ Corporate (Slate -> Emerald)
    fig_grp = px.bar(grp, x='Group', y='% Achieve', text=grp['% Achieve'].apply(lambda x: f"{x:.1f}%"),
                     color='% Achieve', color_continuous_scale=['#f43f5e', '#f59e0b', '#10b981'])
    fig_grp.update_traces(textposition='outside', textfont=dict(weight='bold'))
    fig_grp.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
        margin=dict(t=20, b=0, l=0, r=0), yaxis_range=[0, 115],
        coloraxis_showscale=False # ซ่อนแถบสีด้านข้างให้ดูสะอาด
    )
    fig_grp.update_yaxes(showgrid=True, gridcolor='#f1f5f9')
    st.plotly_chart(fig_grp, use_container_width=True)

with col_top:
    st.markdown("### 🏆 Top 5 ชิ้นงาน (Best & Worst Performers)")
    
    pt = df.groupby('Part').agg({'actual_qty': 'sum', 'เป้าหมายสุทธิ': 'sum'}).reset_index()
    pt = pt[pt['เป้าหมายสุทธิ'] > 0]
    pt['% Achieve'] = ((pt['actual_qty'] / pt['เป้าหมายสุทธิ']) * 100).clip(upper=100.0).round(2).fillna(0)
    
    t_best = pt.sort_values(by='% Achieve', ascending=False).head(5)
    t_worst = pt.sort_values(by='% Achieve', ascending=True).head(5)
    
    tb1, tb2 = st.tabs(["⭐ 5 อันดับแรก (Top Performers)", "⚠️ 5 อันดับรั้งท้าย (Needs Attention)"])
    
    fmt = {'actual_qty': '{:,.0f}', 'เป้าหมายสุทธิ': '{:,.0f}', '% Achieve': '{:.1f}%'}
    conf = {"% Achieve": st.column_config.ProgressColumn("Achieve (%)", format="%.1f%%", min_value=0, max_value=100)}
    
    with tb1:
        st.dataframe(t_best[['Part', 'actual_qty', 'เป้าหมายสุทธิ', '% Achieve']].style.format(fmt), use_container_width=True, hide_index=True, column_config=conf)
    with tb2:
        st.dataframe(t_worst[['Part', 'actual_qty', 'เป้าหมายสุทธิ', '% Achieve']].style.format(fmt), use_container_width=True, hide_index=True, column_config=conf)

st.markdown("<hr>", unsafe_allow_html=True)

# ==========================================
# 📋 Section 4: Data Table & Export
# ==========================================
col_hdr, col_btn = st.columns([4, 1])
with col_hdr:
    st.markdown("### 📋 ข้อมูลการผลิตเชิงลึก (Detailed Production Data)")
with col_btn:
    with st.popover("📤 Export Report", use_container_width=True):
        st.write("เลือกรูปแบบที่ต้องการ:")
        
        out_df = df[['วันที่ผลิต', 'Machine', 'Part', 'actual_qty', 'เป้าต่อวัน(3กะ)', 'mult_Net', 'Setup_Count', 'เป้าหมายสุทธิ', '% Achieve']].sort_values(by=['วันที่ผลิต', 'Machine'], ascending=[False, True])
        out_df.rename(columns={'mult_Net': 'อัตราส่วนกะ'}, inplace=True)
        
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as wrt: out_df.to_excel(wrt, index=False, sheet_name='Data')
        
        st.download_button("💾 ดาวน์โหลด Excel", data=buf.getvalue(), file_name=f"Production_Report_{min_date.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.divider()
        components.html("""
            <button onclick="window.parent.print()" style="
                background-color: #0f172a; color: white; border: none; padding: 10px;
                border-radius: 6px; cursor: pointer; width: 100%; font-weight: bold;
            ">🖨️ พิมพ์เป็น PDF</button>
            <div style="font-size:11px; color:#64748b; text-align:center; margin-top:8px;">* กรุณาเปิด 'Background graphics' ตอนพิมพ์</div>
        """, height=90)

st.dataframe(
    out_df, use_container_width=True, hide_index=True,
    column_config={
        "actual_qty": st.column_config.NumberColumn("ยอดผลิตจริง (Pcs)"),
        "เป้าต่อวัน(3กะ)": st.column_config.NumberColumn("เป้าดิบ 3 กะ"),
        "อัตราส่วนกะ": st.column_config.NumberColumn("อัตราส่วนกะ", format="%.2f"),
        "Setup_Count": st.column_config.NumberColumn("ครั้งเปลี่ยน Part"),
        "เป้าหมายสุทธิ": st.column_config.NumberColumn("เป้าสุทธิ (Pcs)"),
        "% Achieve": st.column_config.ProgressColumn("% Achieve", format="%.1f%%", min_value=0, max_value=100)
    }
)