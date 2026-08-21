import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Production Dashboard", layout="wide")
st.title("📊 Dashboard ประสิทธิภาพการผลิตเทียบเป้าหมาย")

# Sidebar
st.sidebar.header("📁 1. อัปโหลดไฟล์ข้อมูล")
uploaded_file = st.sidebar.file_uploader("อัปโหลดไฟล์ data.xlsx ที่อัปเดตทุกวัน", type=["xlsx", "xls", "csv"])

# ระบบค้นหาไฟล์เป้าหมายอ้างอิงเบื้องหลัง (ไม่ต้องอัปโหลด)
target_file_name = None
for f in os.listdir():
    if "เทียบผลิตจริง" in f and ".xlsx" in f:
        target_file_name = f
        break

if target_file_name:
    st.sidebar.success(f"✅ พบไฟล์เป้าหมายอัตโนมัติ: {target_file_name}")
else:
    st.sidebar.error("❌ ไม่พบไฟล์ 'เทียบผลิตจริง JULY26 .xlsx' ในโฟลเดอร์เดียวกัน กรุณาตรวจสอบ")

# ฟังก์ชันดึงชื่อเครื่องจักร (แยกคนละเครื่อง)
def extract_machine(doc_text, cost_center):
    text = str(doc_text).strip()
    if "/" in text:
        parts = text.split("/")
        if len(parts) > 1 and parts[1].strip() != "":
            return parts[1].strip()
            
    cc = str(cost_center).strip()
    if cc and cc.lower() != 'nan':
        return cc
    return "Unknown_MC"

if uploaded_file:
    try:
        # อ่านไฟล์ data.xlsx
        if uploaded_file.name.endswith(".csv"):
            df_pd = pd.read_csv(uploaded_file)
        else:
            df_pd = pd.read_excel(uploaded_file, sheet_name='pd')
            
        df_pd.columns = [str(c).strip() for c in df_pd.columns]

        # เช็คคอลัมน์ที่จำเป็น (A, G, K)
        required_cols = ['Material', 'Qty in Un. of Entry', 'Posting Date']
        missing = [c for c in required_cols if c not in df_pd.columns]
        if missing:
            st.error(f"❌ โครงสร้างไฟล์ผิด ไม่พบคอลัมน์: {', '.join(missing)}")
            st.stop()

        # ==========================================
        # 📌 1. ทำความสะอาดและดึงข้อมูลตามเงื่อนไข
        # ==========================================
        
        # 1.1 Qty in Un. of Entry (G) -> Actual Qty
        df_pd['Actual_Qty'] = pd.to_numeric(df_pd['Qty in Un. of Entry'], errors='coerce').fillna(0)
        
        # 1.2 Posting Date (K) -> Date
        df_pd['Posting Date'] = pd.to_datetime(df_pd['Posting Date'], errors='coerce').dt.date
        df_pd = df_pd.dropna(subset=['Posting Date']) # ตัดบรรทัดที่ไม่มีวันที่ทิ้ง
        
        # 1.3 Material (A) -> Part
        df_pd['Material'] = df_pd['Material'].astype(str).str.strip()
        
        # แกะชื่อเครื่องจักร (ใช้สำหรับแยกคนละเครื่อง)
        df_pd['mc'] = df_pd.apply(
            lambda r: extract_machine(r.get('Document Header Text', ''), r.get('Cost Center', '')), axis=1
        )

        # ==========================================
        # 📌 2. จับกลุ่มรวมยอด (วันเดียวกันรวมกัน, คนละเครื่องแยกกัน)
        # ==========================================
        # Group by [วันที่, เครื่องจักร, Part] แล้ว Sum ยอด G
        df_actual = df_pd.groupby(['Posting Date', 'mc', 'Material'], as_index=False)['Actual_Qty'].sum()

        # ==========================================
        # 📌 3. ดึงเป้าหมายจาก "เทียบผลิตจริง JULY26 .xlsx"
        # ==========================================
        df_target = pd.DataFrame(columns=['Material', 'cap_per_day'])
        if target_file_name:
            try:
                df_cap = pd.read_excel(target_file_name, sheet_name='rou capday')
                df_cap.columns = [str(c).strip() for c in df_cap.columns]
                
                cap_col = [c for c in df_cap.columns if 'cap' in c.lower() or 'target' in c.lower()]
                if cap_col:
                    df_target = df_cap[['Material', cap_col[0]]].rename(columns={cap_col[0]: 'cap_per_day'})
                    df_target['Material'] = df_target['Material'].astype(str).str.strip()
                    df_target['cap_per_day'] = pd.to_numeric(df_target['cap_per_day'], errors='coerce').fillna(0)
                    df_target = df_target.drop_duplicates(subset=['Material'])
            except Exception as e:
                st.error(f"อ่านชีต 'rou capday' ไม่สำเร็จ: {e}")

        # นำยอด Actual ไปชนกับ Target ตาม Material
        df_merged = pd.merge(df_actual, df_target, on='Material', how='left')
        
        # ถ้าระบบหา Target ไม่เจอ "จะปรับให้เป้าเป็น 0" (ไม่แอบใช้ยอด Actual เป็น Target เหมือนเดิมแล้ว เพื่อให้ข้อมูลตรงความจริง)
        df_merged['cap_per_day'] = df_merged['cap_per_day'].fillna(0)

        # ==========================================
        # 📌 4. คำนวณประสิทธิภาพ
        # ==========================================
        df_merged['Efficiency (%)'] = df_merged.apply(
            lambda r: (r['Actual_Qty'] / r['cap_per_day'] * 100) if r['cap_per_day'] > 0 else 0, axis=1
        )

        # ---------------------------------------------------------
        # 📊 ส่วนแสดงผล Dashboard (เป้าหมาย 3 ข้อ)
        # ---------------------------------------------------------
        
        # 🎯 3.1 ประสิทธิภาพโดยรวม
        st.subheader("1. ประสิทธิภาพการผลิตโดยรวม (Overall Efficiency)")
        total_act = df_merged['Actual_Qty'].sum()
        total_tgt = df_merged['cap_per_day'].sum()
        overall_eff = (total_act / total_tgt * 100) if total_tgt > 0 else 0

        col1, col2, col3 = st.columns(3)
        col1.metric("ยอดผลิตจริง (ชิ้น)", f"{total_act:,.0f}")
        col2.metric("เป้าหมายรวม (ชิ้น)", f"{total_tgt:,.0f}")
        col3.metric("ประสิทธิภาพโดยรวม", f"{overall_eff:.2f}%")

        st.divider()
        col_left, col_right = st.columns(2)

        # 🎯 3.2 ประสิทธิภาพแต่ละเครื่อง (คนละเครื่องแยกแท่งกันชัดเจน)
        with col_left:
            st.subheader("⚙️ 2. ประสิทธิภาพแยกตามแต่ละเครื่อง")
            df_mc = df_merged.groupby('mc')[['Actual_Qty', 'cap_per_day']].sum().reset_index()
            df_mc['Efficiency (%)'] = df_mc.apply(
                lambda r: (r['Actual_Qty'] / r['cap_per_day'] * 100) if r['cap_per_day'] > 0 else 0, axis=1
            )
            
            fig_mc = px.bar(
                df_mc, x='mc', y='Efficiency (%)',
                text=df_mc['Efficiency (%)'].apply(lambda x: f"{x:.1f}%"),
                color='Efficiency (%)',
                color_continuous_scale=['#FF4B4B', '#FFE800', '#00CC96'],
                labels={'mc': 'เครื่องจักร'}
            )
            fig_mc.update_traces(textposition='outside')
            fig_mc.update_layout(yaxis_range=[0, max(df_mc['Efficiency (%)'].max() * 1.2, 110)])
            st.plotly_chart(fig_mc, use_container_width=True)

        # 🎯 3.3 ประสิทธิภาพแต่ละวัน (วันเดียวกันรวมเป็น 1 จุด)
        with col_right:
            st.subheader("📅 3. ประสิทธิภาพแยกตามแต่ละวัน")
            df_day = df_merged.groupby('Posting Date')[['Actual_Qty', 'cap_per_day']].sum().reset_index()
            df_day['Efficiency (%)'] = df_day.apply(
                lambda r: (r['Actual_Qty'] / r['cap_per_day'] * 100) if r['cap_per_day'] > 0 else 0, axis=1
            )
            df_day['Posting Date'] = df_day['Posting Date'].astype(str)

            fig_day = px.line(
                df_day, x='Posting Date', y='Efficiency (%)', markers=True,
                text=df_day['Efficiency (%)'].apply(lambda x: f"{x:.1f}%")
            )
            fig_day.update_traces(textposition="top center")
            fig_day.add_hline(y=100, line_dash="dash", line_color="green", annotation_text="Target 100%")
            st.plotly_chart(fig_day, use_container_width=True)

        # เช็คข้อมูลดิบแบบละเอียด
        st.subheader("📋 ตารางข้อมูลสรุป (ตรวจสอบการดึงคอลัมน์ A, G, K)")
        st.dataframe(
            df_merged[['Posting Date', 'mc', 'Material', 'Actual_Qty', 'cap_per_day', 'Efficiency (%)']]
            .rename(columns={
                'Posting Date': 'วันที่ (K)',
                'Material': 'Part (A)',
                'Actual_Qty': 'ผลิตจริง (G)',
                'cap_per_day': 'เป้าหมาย/วัน'
            })
            .sort_values(by=['วันที่ (K)', 'mc', 'Part (A)'])
            .style.format({'ผลิตจริง (G)': '{:,.0f}', 'เป้าหมาย/วัน': '{:,.0f}', 'Efficiency (%)': '{:.2f}%'}),
            use_container_width=True
        )

        # แสดงรายการ Part ที่ดึง Target ไม่เจอ (เพื่อเช็คหาข้อผิดพลาด)
        missing_parts = df_merged[df_merged['เป้าหมาย/วัน'] == 0]['Part (A)'].unique()
        if len(missing_parts) > 0:
            with st.expander("⚠️ พบรายการ Part ที่ไม่พบข้อมูลเป้าหมายในชีต 'rou capday'"):
                st.write(missing_parts)

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")