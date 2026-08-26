import streamlit as st
import pandas as pd
import plotly.express as px
import os

# ตั้งค่าหน้าเพจ Dashboard
st.set_page_config(page_title="Production Efficiency Dashboard", layout="wide")
st.title("📊 Dashboard ประสิทธิภาพการผลิตเทียบเป้าหมาย")

# Sidebar สำหรับอัปโหลด
st.sidebar.header("📁 อัปโหลดไฟล์ข้อมูลรายวัน")
uploaded_file = st.sidebar.file_uploader("อัปโหลดไฟล์ data.xlsx หรือ data.csv", type=["xlsx", "csv"])

# --- 📌 ระบบจำไฟล์อัตโนมัติ (Fallback System) ---
saved_xlsx = "saved_data_eff.xlsx"
saved_csv = "saved_data_eff.csv"
active_file = None

if uploaded_file is not None:
    active_file = uploaded_file
    if st.sidebar.button("💾 บันทึกไฟล์ข้อมูลนี้ไว้ใช้รอบหน้า", use_container_width=True):
        # ล้างไฟล์เก่าออกก่อน
        if os.path.exists(saved_xlsx): os.remove(saved_xlsx)
        if os.path.exists(saved_csv): os.remove(saved_csv)
        
        # เช็คว่าเป็น csv หรือ xlsx
        ext = ".csv" if uploaded_file.name.endswith('.csv') else ".xlsx"
        with open(f"saved_data_eff{ext}", "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ บันทึกไฟล์เรียบร้อย! คราวหน้าไม่ต้องอัปโหลดซ้ำแล้วครับ")
    st.sidebar.caption("🟢 กำลังแสดงผลจาก: **ไฟล์ที่เพิ่งอัปโหลด**")
else:
    # เช็คว่ามีไฟล์เก่าเซฟไว้หรือไม่
    if os.path.exists(saved_xlsx):
        active_file = saved_xlsx
    elif os.path.exists(saved_csv):
        active_file = saved_csv
        
    if active_file:
        st.sidebar.caption("📌 กำลังแสดงผลจาก: **ไฟล์ที่บันทึกไว้ล่าสุด**")
        if st.sidebar.button("🗑️ ล้างข้อมูลไฟล์ที่บันทึกไว้", use_container_width=True):
            os.remove(active_file)
            st.rerun()

# ฟังก์ชันดึง Master Data เป้าหมาย (Capacity per day) จากไฟล์ในโฟลเดอร์โดยอัตโนมัติ
@st.cache_data
def load_target_data():
    # ระบบจะพยายามค้นหาไฟล์ capday ในโฟลเดอร์ทำงานให้อัตโนมัติ
    possible_target_files = ['rou capday.csv', 'capday.csv', 'เทียบผลิตจริง JULY26 .xlsx -  rou capday.csv']
    for fname in possible_target_files:
        if os.path.exists(fname):
            try:
                df_target = pd.read_csv(fname)
                df_target.columns = [c.strip() for c in df_target.columns]
                return df_target
            except Exception:
                pass
    return pd.DataFrame(columns=['Material', 'cap/day'])

# ฟังก์ชันสกัดหาชื่อเครื่องจักร (mc) จาก Batch หรือ Cost Center
def get_machine_name(row):
    batch = str(row.get('Batch', '')).strip()
    if batch and batch.lower() != 'nan' and '/' in batch:
        parts = batch.split('/')
        if len(parts) > 1 and parts[1].strip():
            return parts[1].strip()
            
    cost_center = str(row.get('Cost Center', '')).strip()
    if cost_center and cost_center.lower() != 'nan':
        return cost_center
        
    return 'Unknown_MC'

if active_file:
    try:
        # เช็คชนิดของไฟล์เพื่อใช้วิธีอ่านที่ถูกต้อง
        is_csv = False
        if hasattr(active_file, 'name'):
            is_csv = active_file.name.endswith('.csv')
        elif isinstance(active_file, str):
            is_csv = active_file.endswith('.csv')

        # 1. อ่านข้อมูลจากไฟล์ data
        if is_csv:
            df_pd = pd.read_csv(active_file)
        else:
            xls = pd.ExcelFile(active_file)
            sheet_target = 'pd' if 'pd' in xls.sheet_names else xls.sheet_names[0]
            df_pd = pd.read_excel(active_file, sheet_name=sheet_target)

        # 2. ทำความสะอาดข้อมูลวันที่
        if 'Posting Date' in df_pd.columns:
            df_pd['Posting Date'] = pd.to_datetime(df_pd['Posting Date']).dt.date
        else:
            st.error("ไม่พบคอลัมน์ 'Posting Date' ในไฟล์ที่อัปโหลด")
            st.stop()

        # 3. สกัดหาชื่อเครื่องจักร (mc)
        df_pd['mc'] = df_pd.apply(get_machine_name, axis=1)

        # 4. แปลงคอลัมน์จำนวนผลิตจริง (Qty in Un. of Entry)
        if 'Qty in Un. of Entry' in df_pd.columns:
            df_pd['Qty in Un. of Entry'] = pd.to_numeric(df_pd['Qty in Un. of Entry'], errors='coerce').fillna(0)
        else:
            st.error("ไม่พบคอลัมน์ 'Qty in Un. of Entry' ในไฟล์ที่อัปโหลด")
            st.stop()

        # 5. รวมยอดผลิตจริงแยกตาม วันที่, เครื่องจักร (mc), และ Material
        df_actual = df_pd.groupby(['Posting Date', 'mc', 'Material'])['Qty in Un. of Entry'].sum().reset_index()
        df_actual.rename(columns={'Qty in Un. of Entry': 'Actual_Qty'}, inplace=True)

        # 6. ดึงข้อมูล Target (Cap/Day)
        df_target = load_target_data()
        
        if not df_target.empty:
            target_col = [c for c in df_target.columns if 'cap' in c.lower() or 'target' in c.lower()]
            if target_col:
                df_target.rename(columns={target_col[0]: 'cap_per_day'}, inplace=True)
            else:
                df_target['cap_per_day'] = 0
        else:
            df_target = pd.DataFrame(columns=['Material', 'cap_per_day'])

        # 7. รวมข้อมูลผลิตจริงเข้ากับเป้าหมาย
        df_merged = pd.merge(df_actual, df_target, on='Material', how='left')
        
        # กรณี Material ตัวไหนไม่มีข้อมูล cap_per_day ให้ใช้ค่า Actual_Qty แทนชั่วคราว
        df_merged['cap_per_day'] = df_merged['cap_per_day'].fillna(df_merged['Actual_Qty'])
        
        # คำนวณ % ประสิทธิภาพ
        df_merged['Efficiency (%)'] = df_merged.apply(
            lambda row: (row['Actual_Qty'] / row['cap_per_day'] * 100) if row['cap_per_day'] > 0 else 0, axis=1
        )

        # ==========================================
        # 📈 ส่วนแสดงผล DASHBOARD
        # ==========================================
        
        # --- 1. ประสิทธิภาพการผลิตเทียบเป้าหมายโดยรวม ---
        st.subheader("📌 1. ประสิทธิภาพการผลิตโดยรวม (Overall Efficiency)")
        total_actual = df_merged['Actual_Qty'].sum()
        total_target = df_merged['cap_per_day'].sum()
        overall_eff = (total_actual / total_target * 100) if total_target > 0 else 0

        col1, col2, col3 = st.columns(3)
        col1.metric("ยอดผลิตจริงรวม (ชิ้น)", f"{total_actual:,.0f}")
        col2.metric("เป้าหมายรวม (ชิ้น)", f"{total_target:,.0f}")
        col3.metric("ประสิทธิภาพภาพรวม", f"{overall_eff:.2f}%")

        st.divider()

        # --- กราฟแยกเครื่องจักร และ กราฟรายวัน ---
        col_left, col_right = st.columns(2)

        # --- 2. ประสิทธิภาพการผลิตเทียบเป้าหมายแต่ละเครื่อง ---
        with col_left:
            st.subheader("⚙️ 2. ประสิทธิภาพเทียบเป้าหมาย (รายเครื่องจักร)")
            df_mc = df_merged.groupby('mc')[['Actual_Qty', 'cap_per_day']].sum().reset_index()
            df_mc['Efficiency (%)'] = df_mc.apply(
                lambda row: (row['Actual_Qty'] / row['cap_per_day'] * 100) if row['cap_per_day'] > 0 else 0, axis=1
            )
            
            fig_mc = px.bar(
                df_mc, 
                x='mc', 
                y='Efficiency (%)',
                text=df_mc['Efficiency (%)'].apply(lambda x: f"{x:.1f}%"),
                color='Efficiency (%)',
                color_continuous_scale=['#FF4B4B', '#FFE800', '#00CC96'],
                labels={'mc': 'เครื่องจักร', 'Efficiency (%)': 'ประสิทธิภาพ (%)'}
            )
            fig_mc.update_traces(textposition='outside')
            fig_mc.update_layout(yaxis_range=[0, max(df_mc['Efficiency (%)'].max() * 1.2, 110)])
            st.plotly_chart(fig_mc, use_container_width=True)

        # --- 3. ประสิทธิภาพการผลิตเทียบเป้าหมายแต่ละวัน ---
        with col_right:
            st.subheader("📅 3. ประสิทธิภาพเทียบเป้าหมาย (รายวัน)")
            df_day = df_merged.groupby('Posting Date')[['Actual_Qty', 'cap_per_day']].sum().reset_index()
            df_day['Efficiency (%)'] = df_day.apply(
                lambda row: (row['Actual_Qty'] / row['cap_per_day'] * 100) if row['cap_per_day'] > 0 else 0, axis=1
            )
            df_day['Posting Date'] = df_day['Posting Date'].astype(str)

            fig_day = px.line(
                df_day, 
                x='Posting Date', 
                y='Efficiency (%)',
                markers=True,
                text=df_day['Efficiency (%)'].apply(lambda x: f"{x:.1f}%"),
                labels={'Posting Date': 'วันที่', 'Efficiency (%)': 'ประสิทธิภาพ (%)'}
            )
            fig_day.update_traces(textposition="top center")
            fig_day.add_hline(y=100, line_dash="dash", line_color="green", annotation_text="Target 100%")
            st.plotly_chart(fig_day, use_container_width=True)

        # --- ตารางรายละเอียดข้อมูล ---
        st.subheader("📋 ตารางสรุปข้อมูลรายละเอียด (Summary Data Table)")
        st.dataframe(
            df_merged[['Posting Date', 'mc', 'Material', 'Actual_Qty', 'cap_per_day', 'Efficiency (%)']]
            .sort_values(by=['Posting Date', 'mc'])
            .style.format({
                'Actual_Qty': '{:,.0f}',
                'cap_per_day': '{:,.0f}',
                'Efficiency (%)': '{:.2f}%'
            }),
            use_container_width=True
        )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลข้อมูล: {e}")
        st.info("💡 คำแนะนำ: ตรวจสอบว่าไฟล์ที่อัปโหลดมีคอลัมน์ 'Posting Date', 'Batch' (หรือ Cost Center) และ 'Qty in Un. of Entry' ครบถ้วน")
else:
    st.info("👈 กรุณาอัปโหลดไฟล์ data.xlsx หรือ data.csv ที่เมนูด้านซ้ายเพื่อเริ่มการทำงาน")