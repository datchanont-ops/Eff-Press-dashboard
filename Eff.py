import pandas as pd
import streamlit as st
import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit.components.v1 as components
import os
import io
import json
import base64
import requests

# --- 1. ตั้งค่าหน้าจอ Streamlit (ต้องอยู่บนสุดเสมอ) ---
st.set_page_config(page_title="Press Daily Production Dashboard", layout="wide")

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
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }

def gh_get_file(remote_path):
    """คืนค่า (content_bytes, sha) หรือ (None, None) ถ้าไม่พบไฟล์"""
    if not GITHUB_ENABLED:
        return None, None
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}?ref={GITHUB_BRANCH}"
    try:
        r = requests.get(url, headers=gh_headers(), timeout=15)
        if r.status_code == 200:
            data = r.json()
            content = base64.b64decode(data["content"])
            return content, data["sha"]
    except Exception:
        pass
    return None, None

def gh_put_file(remote_path, content_bytes, message):
    """เขียน/อัปเดตไฟล์ใน GitHub repo คืนค่า True/False"""
    if not GITHUB_ENABLED:
        return False
    _, sha = gh_get_file(remote_path)
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    try:
        r = requests.put(url, headers=gh_headers(), json=payload, timeout=15)
        return r.status_code in (200, 201)
    except Exception:
        return False

# --- CSS สำหรับจัดระเบียบตอนสั่ง Print เป็น PDF ---
st.markdown("""
<style>
    @media print {
        .stPopover { display: none !important; }
        .stExpander { display: none !important; }
        .stDownloadButton { display: none !important; }
        header { display: none !important; }
        [data-testid="stSidebar"] { display: none !important; }
    }
</style>
""", unsafe_allow_html=True)

# --- 2. ฟังก์ชันโหลดไฟล์เป้าหมาย (หลังบ้าน) ---
@st.cache_data
def load_target_data():
    try:
        df_target = pd.read_excel('target.xlsx')
        if 'Material' in df_target.columns:
            df_target = df_target.rename(columns={'Material': 'Part', 'cap/day': 'เป้าต่อวัน(3กะ)'})
        else:
            df_target.columns = ['Part', 'เป้าต่อวัน(3กะ)']
            
        df_target.columns = df_target.columns.str.strip()
        return df_target, None
    except Exception as e:
        return None, "❌ **ไม่พบไฟล์เป้าหมาย:** กรุณาสร้างไฟล์เป้าหมายการผลิต ตั้งชื่อว่า `target.xlsx` แล้วนำมาวางไว้ในโฟลเดอร์เดียวกับโปรแกรมครับ"

# --- 3. ฟังก์ชันโหลดไฟล์ผลิตรายวัน ---
@st.cache_data
def load_daily_data(file, df_target):
    try:
        xls = pd.ExcelFile(file)
        if 'pd' in xls.sheet_names:
            df_pd = pd.read_excel(xls, 'pd')
        else:
            df_pd = pd.read_excel(xls, 0)

        df = df_pd[['Material', 'Document Header Text', 'Qty in Un. of Entry', 'Posting Date', 'Entry Date', 'Time of Entry']].copy()

        def extract_machine(text):
            if pd.isna(text): return None
            parts = str(text).split('/')
            if len(parts) >= 2 and str(text).startswith('1/'):
                return parts[1]
            return None

        df['Machine'] = df['Document Header Text'].apply(extract_machine)

        df_filtered = df.dropna(subset=['Machine']).copy()
        df_filtered = df_filtered.rename(columns={
            'Posting Date': 'วันที่ผลิต',
            'Material': 'Part',
            'Qty in Un. of Entry': 'actual_qty'
        })

        df_filtered['วันที่ผลิต'] = pd.to_datetime(df_filtered['วันที่ผลิต']).dt.date
        df_filtered['Entry Date'] = pd.to_datetime(df_filtered['Entry Date']).dt.date

        # --- ตรวจจับการเปลี่ยน Part (Setup) ---
        df_filtered = df_filtered.sort_values(by=['Machine', 'วันที่ผลิต', 'Entry Date', 'Time of Entry'])
        df_filtered['Part_ก่อนหน้า'] = df_filtered.groupby('Machine')['Part'].shift(1)
        df_filtered['Is_Setup'] = (df_filtered['Part'] != df_filtered['Part_ก่อนหน้า']) & (df_filtered['Part_ก่อนหน้า'].notna())

        df_grouped = df_filtered.groupby(['วันที่ผลิต', 'Machine', 'Part'], as_index=False).agg(
            actual_qty=('actual_qty', 'sum'),
            Setup_Count=('Is_Setup', 'sum')
        )

        df_final = pd.merge(df_grouped, df_target, on='Part', how='left')
        df_final['เป้าต่อวัน(3กะ)'] = df_final['เป้าต่อวัน(3กะ)'].fillna(0)

        return df_final, None
    except Exception as e:
        return None, f"เกิดข้อผิดพลาดในการอ่านไฟล์รายวัน: {e}"

# ==========================================
# --- UI หลักของ Dashboard ---
# ==========================================
st.title("🏭 Press Daily Production Dashboard")

# โหลดข้อมูลเป้าหมายจากหลังบ้าน
df_target, target_error = load_target_data()
if target_error:
    st.error(target_error)
    st.stop()

# --- ส่วนดาวน์โหลด Template ---
st.sidebar.header("📥 ดาวน์โหลดแบบฟอร์ม")
template_file_name = "Template.xlsx"

if os.path.exists(template_file_name):
    with open(template_file_name, "rb") as file:
        st.sidebar.download_button(
            label=f"คลิกดาวน์โหลด {template_file_name}",
            data=file,
            file_name=template_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.sidebar.markdown("---")

# ==========================================
# --- ระบบซิงค์ข้อมูลลง Local (ทำครั้งเดียวต่อ session) ---
# ==========================================
default_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'default_data.xlsx')
settings_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'eff_settings.json')

if GITHUB_ENABLED and not st.session_state.get("eff_github_synced"):
    with st.spinner("🔄 กำลังซิงค์ข้อมูลล่าสุดจาก GitHub..."):
        content, _ = gh_get_file(f"{GITHUB_DATA_DIR}/default_data.xlsx")
        if content:
            with open(default_file_path, "wb") as f:
                f.write(content)

        content, _ = gh_get_file(f"{GITHUB_DATA_DIR}/eff_settings.json")
        if content:
            with open(settings_file, "wb") as f:
                f.write(content)
    st.session_state["eff_github_synced"] = True
elif not GITHUB_ENABLED:
    st.sidebar.warning("⚠️ ยังไม่ได้ตั้งค่า GitHub Secrets ข้อมูลจะไม่ถูกบันทึกข้ามการเปิดเว็บใหม่", icon="⚠️")

# ==========================================
# --- ระบบอัปโหลดและเซฟไฟล์เริ่มต้น ---
# ==========================================
st.sidebar.header("📂 อัปโหลดยอดผลิตรายวัน")
uploaded_file = st.sidebar.file_uploader("เลือกไฟล์ Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    if st.sidebar.button("💾 บันทึกไฟล์ข้อมูลนี้ไว้ใช้รอบหน้า", use_container_width=True):
        file_bytes = bytes(uploaded_file.getbuffer())
        with open(default_file_path, "wb") as f:
            f.write(file_bytes)
            
        if GITHUB_ENABLED:
            with st.spinner("☁️ กำลังบันทึกไฟล์ขึ้น GitHub..."):
                ok = gh_put_file(f"{GITHUB_DATA_DIR}/default_data.xlsx", file_bytes, f"Auto-save eff data: {uploaded_file.name}")
            if ok:
                st.sidebar.success("✅ บันทึกไฟล์ขึ้น GitHub เรียบร้อยแล้ว!")
            else:
                st.sidebar.warning("⚠️ บันทึกลง local ได้ แต่ push ไป GitHub ไม่สำเร็จ")
        else:
            st.sidebar.success("✅ บันทึกเป็นไฟล์เริ่มต้น (Local) เรียบร้อย!")

data_source = None
if uploaded_file is not None:
    data_source = uploaded_file
    st.sidebar.caption("🟢 กำลังแสดงผลจาก: **ไฟล์ที่เพิ่งอัปโหลด**")
elif os.path.exists(default_file_path):
    data_source = default_file_path
    st.sidebar.caption("📌 กำลังแสดงผลจาก: **ไฟล์ที่บันทึกไว้ในระบบ (Default)**")

# ตรวจสอบว่ามีแหล่งข้อมูลหรือไม่
if data_source is None:
    st.info("👈 กรุณาอัปโหลดไฟล์ Excel ที่เมนูด้านซ้ายเพื่อเริ่มการคำนวณครับ")
    st.stop()

# ==========================================
# --- โหลดข้อมูลและเริ่มประมวลผล ---
# ==========================================
df, error = load_daily_data(data_source, df_target)

if error:
    st.error(error)
else:
    min_date_db = df['วันที่ผลิต'].min()
    max_date_db = df['วันที่ผลิต'].max()
    st.caption(f"📂 ฐานข้อมูลภาพรวมทั้งหมดในไฟล์: {min_date_db.strftime('%d/%m/%Y')} ถึง {max_date_db.strftime('%d/%m/%Y')}")
    
    # 📌 โหลดไฟล์ตั้งค่าเงื่อนไข (eff_settings.json)
    saved_settings = {}
    if os.path.exists(settings_file):
        try:
            with open(settings_file, 'r', encoding='utf-8') as f:
                saved_settings = json.load(f)
        except:
            pass

    # --- ตั้งค่าประสิทธิภาพ ---
    st.sidebar.markdown("---")
    st.sidebar.header("🎯 ตั้งค่าประสิทธิภาพ")
    
    init_oee = int(saved_settings.get('oee_val', 90)) # เริ่มต้นที่ 90%
    init_setup = float(saved_settings.get('setup_hours', 4.0))
    
    oee_val = st.sidebar.number_input("1. ค่า O.E.E. (1-100%)", min_value=1, max_value=100, value=init_oee, step=1)
    oee_multiplier = oee_val / 100.0
    
    setup_hours = st.sidebar.number_input("2. เวลา Setup เปลี่ยน Part (ชั่วโมง)", min_value=0.0, max_value=24.0, value=init_setup, step=0.5)
    setup_deduct_ratio = setup_hours / 24.0
    
    # --- ตั้งค่ากะการผลิต ---
    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ ตั้งค่ากะการผลิต (แบบผสม)")
    st.sidebar.caption("ระบบจะใช้เป้าที่น้อยที่สุด หากมีการตั้งค่าซ้อนทับกัน")

    shift_mapping = {
        "3 กะ (เป้า 100%)": 1.0,
        "2 กะ (เป้า 67%)": 0.67,
        "1.5 กะ (เป้า 50%)": 0.5
    }
    
    saved_shift_date = saved_settings.get('shift_date', {})
    saved_shift_mac = saved_settings.get('shift_mac', {})
    saved_spec = saved_settings.get('shift_spec', [])

    with st.sidebar.expander("📅 1. ตั้งค่ากะรายวัน (By Date)", expanded=False):
        st.caption("ตั้งค่าเริ่มต้นถูกดึงมาจากที่เคยบันทึกไว้")
        default_shift_date = st.selectbox("🔄 เปลี่ยนพร้อมกันทุกวัน:", ["(ใช้ค่าเดิม)"] + list(shift_mapping.keys()), index=0, key='def_date')
        unique_dates = sorted(df['วันที่ผลิต'].unique())
        
        if default_shift_date != "(ใช้ค่าเดิม)":
            shift_date_data = [default_shift_date] * len(unique_dates)
        else:
            shift_date_data = [saved_shift_date.get(str(d), "3 กะ (เป้า 100%)") for d in unique_dates]
            
        shift_date_df = pd.DataFrame({'วันที่': unique_dates, 'กะการทำงาน': shift_date_data})

        edited_shift_date_df = st.data_editor(
            shift_date_df,
            column_config={
                "วันที่": st.column_config.DateColumn("วันที่", disabled=True, format="DD/MM/YYYY"),
                "กะการทำงาน": st.column_config.SelectboxColumn("กะการทำงาน", options=list(shift_mapping.keys()), required=True)
            },
            hide_index=True, use_container_width=True, key='edit_date'
        )
        shift_multiplier_date = {row['วันที่']: shift_mapping[row['กะการทำงาน']] for _, row in edited_shift_date_df.iterrows()}
        df['ตัวคูณกะ_Date'] = df['วันที่ผลิต'].map(shift_multiplier_date)

    with st.sidebar.expander("🚜 2. ตั้งค่ากะรายเครื่องจักร (By Machine)", expanded=False):
        st.caption("ตั้งค่าเริ่มต้นถูกดึงมาจากที่เคยบันทึกไว้")
        default_shift_machine = st.selectbox("🔄 เปลี่ยนพร้อมกันทุกเครื่อง:", ["(ใช้ค่าเดิม)"] + list(shift_mapping.keys()), index=0, key='def_mac')
        unique_machines = sorted(df['Machine'].unique())
        
        if default_shift_machine != "(ใช้ค่าเดิม)":
            shift_mac_data = [default_shift_machine] * len(unique_machines)
        else:
            shift_mac_data = [saved_shift_mac.get(m, "3 กะ (เป้า 100%)") for m in unique_machines]
            
        shift_machine_df = pd.DataFrame({'Machine': unique_machines, 'กะการทำงาน': shift_mac_data})

        edited_shift_machine_df = st.data_editor(
            shift_machine_df,
            column_config={
                "Machine": st.column_config.TextColumn("ชื่อเครื่องจักร", disabled=True),
                "กะการทำงาน": st.column_config.SelectboxColumn("กะการทำงาน", options=list(shift_mapping.keys()), required=True)
            },
            hide_index=True, use_container_width=True, key='edit_mac'
        )
        shift_multiplier_machine = {row['Machine']: shift_mapping[row['กะการทำงาน']] for _, row in edited_shift_machine_df.iterrows()}
        df['ตัวคูณกะ_Machine'] = df['Machine'].map(shift_multiplier_machine)

    with st.sidebar.expander("🎯 3. ตั้งค่ากะเฉพาะกิจ (เครื่อง + วันที่)", expanded=False):
        if saved_spec:
            empty_spec_df = pd.DataFrame(saved_spec)
            if 'วันที่' in empty_spec_df.columns:
                empty_spec_df['วันที่'] = pd.to_datetime(empty_spec_df['วันที่']).dt.date
        else:
            empty_spec_df = pd.DataFrame(columns=["วันที่", "เครื่องจักร", "กะการทำงาน"])
        
        edited_spec_df = st.data_editor(
            empty_spec_df,
            num_rows="dynamic",
            column_config={
                "วันที่": st.column_config.DateColumn("วันที่", required=True, format="DD/MM/YYYY"),
                "เครื่องจักร": st.column_config.SelectboxColumn("เครื่องจักร", options=unique_machines, required=True),
                "กะการทำงาน": st.column_config.SelectboxColumn("กะการทำงาน", options=list(shift_mapping.keys()), required=True)
            },
            hide_index=True, use_container_width=True, key='edit_spec'
        )

    # --- 📌 ปุ่มบันทึกเงื่อนไข (Save Conditions) ---
    st.sidebar.markdown("---")
    if st.sidebar.button("💾 บันทึกเงื่อนไข (Save Conditions)", use_container_width=True):
        sd_dict = {str(row['วันที่']): row['กะการทำงาน'] for _, row in edited_shift_date_df.iterrows()}
        sm_dict = {row['Machine']: row['กะการทำงาน'] for _, row in edited_shift_machine_df.iterrows()}
        
        spec_list = []
        if not edited_spec_df.empty:
            for _, row in edited_spec_df.dropna().iterrows():
                spec_list.append({
                    'วันที่': str(row['วันที่']),
                    'เครื่องจักร': row['เครื่องจักร'],
                    'กะการทำงาน': row['กะการทำงาน']
                })
        
        settings_to_save = {
            'oee_val': oee_val,
            'setup_hours': setup_hours,
            'shift_date': sd_dict,
            'shift_mac': sm_dict,
            'shift_spec': spec_list
        }
        
        try:
            json_str = json.dumps(settings_to_save, ensure_ascii=False, indent=4)
            with open(settings_file, 'w', encoding='utf-8') as f:
                f.write(json_str)
                
            if GITHUB_ENABLED:
                with st.spinner("☁️ กำลังบันทึกเงื่อนไขขึ้น GitHub..."):
                    ok = gh_put_file(f"{GITHUB_DATA_DIR}/eff_settings.json", json_str.encode('utf-8'), "Auto-save eff settings")
                if ok:
                    st.sidebar.success("✅ บันทึกเงื่อนไขขึ้น GitHub เรียบร้อยแล้ว!")
                else:
                    st.sidebar.warning("⚠️ บันทึกเงื่อนไขลง local ได้ แต่ push ไป GitHub ไม่สำเร็จ")
            else:
                st.sidebar.success("✅ บันทึกเงื่อนไขลง Local เรียบร้อยแล้ว!")
        except Exception as e:
            st.sidebar.error(f"❌ บันทึกไม่สำเร็จ: {e}")

    # --- คำนวณตัวคูณกะสุทธิ ---
    df['ตัวคูณกะสุทธิ'] = df[['ตัวคูณกะ_Date', 'ตัวคูณกะ_Machine']].min(axis=1)

    if not edited_spec_df.empty:
        for _, row in edited_spec_df.dropna().iterrows():
            try:
                spec_date = row['วันที่']
                spec_mac = row['เครื่องจักร']
                spec_mult = shift_mapping[row['กะการทำงาน']]
                mask = (df['วันที่ผลิต'] == spec_date) & (df['Machine'] == spec_mac)
                df.loc[mask, 'ตัวคูณกะสุทธิ'] = spec_mult
            except Exception:
                pass

    adjusted_dates = edited_shift_date_df[edited_shift_date_df['กะการทำงาน'] != "3 กะ (เป้า 100%)"]
    adjusted_macs = edited_shift_machine_df[edited_shift_machine_df['กะการทำงาน'] != "3 กะ (เป้า 100%)"]
    adjusted_specs = edited_spec_df[edited_spec_df['กะการทำงาน'] != "3 กะ (เป้า 100%)"].dropna()
    
    if not adjusted_dates.empty or not adjusted_macs.empty or not adjusted_specs.empty:
        alert_text = "**⚠️ สรุปรายการปรับลดกะ:**\n"
        if not adjusted_dates.empty:
            for _, row in adjusted_dates.iterrows():
                alert_text += f"- {row['วันที่'].strftime('%d/%m/%Y')} ➔ {row['กะการทำงาน']}\n"
        if not adjusted_macs.empty:
            for _, row in adjusted_macs.iterrows():
                alert_text += f"- {row['Machine']} ➔ {row['กะการทำงาน']}\n"
        if not adjusted_specs.empty:
            for _, row in adjusted_specs.iterrows():
                alert_text += f"- {row['วันที่'].strftime('%d/%m/%Y')} | {row['เครื่องจักร']} ➔ {row['กะการทำงาน']}\n"
        st.sidebar.warning(alert_text)

    # --- คำนวณเป้าหมายที่ปรับแล้ว ---
    reverse_shift_mapping = {1.0: "3 กะ", 0.67: "2 กะ", 0.5: "1.5 กะ"}
    df['จำนวนกะ'] = df['ตัวคูณกะสุทธิ'].map(reverse_shift_mapping)

    df['เป้าหมายก่อนหักSetup'] = df['เป้าต่อวัน(3กะ)'] * df['ตัวคูณกะสุทธิ'] * oee_multiplier
    df['ยอดลดเป้าSetup'] = (df['เป้าต่อวัน(3กะ)'] * setup_deduct_ratio) * df['Setup_Count']
    
    df['เป้าหมายที่ปรับแล้ว'] = df['เป้าหมายก่อนหักSetup'] - df['ยอดลดเป้าSetup']
    df['เป้าหมายที่ปรับแล้ว'] = df['เป้าหมายที่ปรับแล้ว'].clip(lower=0)
    
    # 📌 จำกัด % Achieve สูงสุดที่ 100%
    df['% Achieve'] = (df['actual_qty'] / df['เป้าหมายที่ปรับแล้ว']) * 100
    df['% Achieve'] = df['% Achieve'].clip(upper=100.0).round(2).fillna(0)

    # --- 4. ตัวกรองข้อมูล (Filters) ---
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 ตัวกรองข้อมูล (Filters)")
    
    date_range = st.sidebar.date_input("📅 เลือกช่วงวันที่แสดงผล", value=(min_date_db, max_date_db), min_value=min_date_db, max_value=max_date_db)
    
    start_disp_date = min_date_db
    end_disp_date = max_date_db
    
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_disp_date, end_disp_date = date_range
        df = df[(df['วันที่ผลิต'] >= start_disp_date) & (df['วันที่ผลิต'] <= end_disp_date)]
    elif isinstance(date_range, tuple) and len(date_range) == 1:
        start_disp_date = date_range[0]
        end_disp_date = date_range[0]
        df = df[df['วันที่ผลิต'] == start_disp_date]

    st.sidebar.markdown("🧪 **การกรองงานทดลองผลิต (Trial)**")
    trial_option = st.sidebar.selectbox("เลือกเงื่อนไขการตัดงานทดลองผลิต:", ["แสดงทั้งหมด (ไม่ตัด)", "ตัด Part ที่ผลิตเพียง 1 วัน (<= 1 วัน)", "ตัด Part ที่ผลิต 1 - 2 วัน (<= 2 วัน)", "กำหนดจำนวนวันเอง (Custom)"], index=0)

    part_day_counts = df.groupby('Part')['วันที่ผลิต'].nunique().to_dict()
    df['จำนวนวันผลิตของPart'] = df['Part'].map(part_day_counts)

    cut_days = 0
    if trial_option == "ตัด Part ที่ผลิตเพียง 1 วัน (<= 1 วัน)": cut_days = 1
    elif trial_option == "ตัด Part ที่ผลิต 1 - 2 วัน (<= 2 วัน)": cut_days = 2
    elif trial_option == "กำหนดจำนวนวันเอง (Custom)": cut_days = st.sidebar.number_input("ตัด Part ที่ผลิตน้อยกว่าหรือเท่ากับ (วัน):", min_value=1, max_value=30, value=2, step=1)

    if cut_days > 0:
        removed_parts = df[df['จำนวนวันผลิตของPart'] <= cut_days]['Part'].unique()
        df = df[df['จำนวนวันผลิตของPart'] > cut_days]
        st.info(f"🧪 **เปิดใช้งานการตัดงานทดลองผลิต (<= {cut_days} วัน):** ตัดออกทั้งหมด `{len(removed_parts)}` Part")

    st.sidebar.markdown("⚙️ **ตัวกรองเครื่องจักร**")
    machine_groups = ['INJ', 'INM', '510', 'VAC', '400T', '300T', '350T']
    selected_groups = st.sidebar.multiselect("1. เลือกกลุ่มเครื่องจักร (Machine Group)", options=machine_groups, default=[])
    
    if selected_groups:
        pattern = '|'.join(selected_groups)
        df = df[df['Machine'].str.contains(pattern, case=False, na=False)]

    available_machines = sorted(df['Machine'].unique())
    selected_machines = st.sidebar.multiselect("2. เลือกเครื่องจักร (ระบุรายเครื่อง)", options=available_machines, default=[])
    if selected_machines:
        df = df[df['Machine'].isin(selected_machines)]
        
    selected_parts = st.sidebar.multiselect("เลือกชิ้นงาน (Part)", options=sorted(df['Part'].unique()), default=[])
    if selected_parts:
        df = df[df['Part'].isin(selected_parts)]

    # --- 📌 แสดงแถบสถานะช่วงวันที่ ---
    days_count = (end_disp_date - start_disp_date).days + 1
    st.success(f"📅 **ช่วงวันที่เลือกแสดงผล:** {start_disp_date.strftime('%d/%m/%Y')} ถึง {end_disp_date.strftime('%d/%m/%Y')} (รวม {days_count:,} วัน)")

    # --- 5. แสดงผลตัวชี้วัด (Metrics) ---
    st.markdown("---")
    total_actual = df['actual_qty'].sum()
    total_target_original = df['เป้าต่อวัน(3กะ)'].sum()
    total_target = df['เป้าหมายที่ปรับแล้ว'].sum()
    overall_achieve = (total_actual / total_target * 100) if total_target > 0 else 0
    overall_achieve = min(overall_achieve, 100.0)

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: st.metric("ยอดผลิตจริง (Actual)", f"{total_actual:,.0f} Pcs")
    with col2: st.metric("เป้า 100% (Original Target)", f"{total_target_original:,.0f} Pcs")
    with col3: st.metric("เป้าสุทธิ (Target)", f"{total_target:,.0f} Pcs")
    with col4: st.metric("ประสิทธิภาพรวม (% Achieve)", f"{overall_achieve:.2f}%")
    with col5: st.metric("⚙️ ค่า O.E.E. (ตั้งค่า)", f"{oee_val}%")

    # --- 6. กราฟ 2 แกน (Plotly Dual-Axis Chart) ---
    st.subheader("📊 เปรียบเทียบยอดผลิตจริง กับ เป้าหมาย (พร้อม % Achieve)")
    daily_summary = df.groupby('วันที่ผลิต').agg({'actual_qty': 'sum', 'เป้าหมายที่ปรับแล้ว': 'sum'}).reset_index()
    daily_summary['% Achieve'] = (daily_summary['actual_qty'] / daily_summary['เป้าหมายที่ปรับแล้ว'] * 100).fillna(0).round(2)
    daily_summary['% Achieve'] = daily_summary['% Achieve'].clip(upper=100.0)
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=daily_summary['วันที่ผลิต'], y=daily_summary['actual_qty'], name="ยอดผลิตจริง (Actual)", marker_color='#1f77b4'), secondary_y=False)
    fig.add_trace(go.Bar(x=daily_summary['วันที่ผลิต'], y=daily_summary['เป้าหมายที่ปรับแล้ว'], name="เป้าหมาย (Target)", marker_color='#ff7f0e'), secondary_y=False)
    fig.add_trace(go.Scatter(x=daily_summary['วันที่ผลิต'], y=daily_summary['% Achieve'], name="% Achieve", mode='lines+markers', line=dict(color='red', width=3), marker=dict(size=8)), secondary_y=True)
    
    fig.update_layout(barmode='group', hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), margin=dict(l=0, r=0, t=30, b=0))
    fig.update_yaxes(title_text="จำนวนชิ้นงาน (Pcs)", secondary_y=False)
    fig.update_yaxes(title_text="ประสิทธิภาพ (% Achieve)", secondary_y=True, ticksuffix="%")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # --- 7. ตารางข้อมูลดิบ และปุ่ม Export เมนูใหม่ ---
    col_table_header, col_export_menu = st.columns([3.5, 1])
    
    with col_table_header:
        st.subheader("📋 รายละเอียดข้อมูลการผลิต (Data Table)")
        
    with col_export_menu:
        st.write("") 
        with st.popover("📥 Export Report"):
            st.markdown("**1. ส่งออกข้อมูลเป็น Excel**")
            
            display_df = df[['วันที่ผลิต', 'Machine', 'Part', 'actual_qty', 'เป้าต่อวัน(3กะ)', 'จำนวนกะ', 'ตัวคูณกะสุทธิ', 'Setup_Count', 'เป้าหมายที่ปรับแล้ว', '% Achieve']]
            display_df = display_df.sort_values(by=['วันที่ผลิต', 'Machine'], ascending=[False, True])
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                display_df.to_excel(writer, index=False, sheet_name='Production_Data')
            excel_data = output.getvalue()
            
            st.download_button(
                label="💾 ดาวน์โหลด Data (.xlsx)",
                data=excel_data,
                file_name=f"Production_Report_{start_disp_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.divider()
            st.markdown("**2. ส่งออกหน้าเว็บพร้อมกราฟเป็น PDF**")
            
            components.html(
                """
                <button onclick="window.parent.print()" style="
                    background-color: #EF553B; 
                    border: none;
                    color: white;
                    padding: 10px 20px;
                    text-align: center;
                    border-radius: 5px;
                    cursor: pointer;
                    width: 100%;
                    font-family: sans-serif;
                    font-weight: bold;
                    font-size: 14px;
                ">🖨️ Print / Save as PDF</button>
                <p style="font-size:12px; color:gray; font-family:sans-serif; text-align:center; margin-top:10px;">
                * แนะนำให้เปิดตัวเลือก <b>'Background graphics'</b> ในตั้งค่า Print
                </p>
                """,
                height=110
            )

    # แสดงตารางข้อมูล
    st.dataframe(
        display_df, 
        use_container_width=True,
        column_config={
            "actual_qty": st.column_config.NumberColumn("ยอดผลิตจริง"),
            "เป้าต่อวัน(3กะ)": st.column_config.NumberColumn("เป้า 3 กะ"),
            "จำนวนกะ": st.column_config.TextColumn("จำนวนกะ"),
            "ตัวคูณกะสุทธิ": st.column_config.NumberColumn("อัตราส่วนกะสุทธิ"),
            "Setup_Count": st.column_config.NumberColumn("จำนวนครั้งเปลี่ยน Part"),
            "เป้าหมายที่ปรับแล้ว": st.column_config.NumberColumn("เป้าสุทธิ"),
            "% Achieve": st.column_config.ProgressColumn("% เทียบเป้า", format="%.2f%%", min_value=0, max_value=100)
        }
    )