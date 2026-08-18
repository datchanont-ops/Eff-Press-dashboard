import pandas as pd
import streamlit as st
import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --- ตั้งค่าหน้าจอ Streamlit ---
st.set_page_config(page_title="Press Daily Production Dashboard", layout="wide")

# --- ฟังก์ชันโหลดและเตรียมข้อมูล ---
@st.cache_data
def load_and_prep_data(file):
    try:
        xls = pd.ExcelFile(file)
        if 'pd' not in xls.sheet_names or ' rou capday' not in xls.sheet_names:
            return None, "Error: ไม่พบชีต 'pd' หรือ ' rou capday' ในไฟล์ที่อัปโหลด โปรดตรวจสอบไฟล์อีกครั้ง"
        
        df_pd = pd.read_excel(xls, 'pd')
        df_target = pd.read_excel(xls, ' rou capday')

        # ดึงข้อมูล (เพิ่ม Entry Date และ Time of Entry เพื่อไว้เรียงลำดับหาการเปลี่ยน Part)
        df = df_pd[['Material', 'Document Header Text', 'Qty in Un. of Entry', 'Posting Date', 'Entry Date', 'Time of Entry']].copy()

        def extract_machine(text):
            if pd.isna(text):
                return None
            parts = str(text).split('/')
            if len(parts) >= 2 and str(text).startswith('1/'):
                return parts[1]
            return None

        df['Machine'] = df['Document Header Text'].apply(extract_machine)

        # กรองและเปลี่ยนชื่อคอลัมน์
        df_filtered = df.dropna(subset=['Machine']).copy()
        df_filtered = df_filtered.rename(columns={
            'Posting Date': 'วันที่ผลิต',
            'Material': 'Part',
            'Qty in Un. of Entry': 'actual_qty'
        })

        # จัดการวันที่
        df_filtered['วันที่ผลิต'] = pd.to_datetime(df_filtered['วันที่ผลิต']).dt.date
        df_filtered['Entry Date'] = pd.to_datetime(df_filtered['Entry Date']).dt.date

        # --- ลอจิกตรวจจับการเปลี่ยน Part (Setup) ---
        df_filtered = df_filtered.sort_values(by=['Machine', 'วันที่ผลิต', 'Entry Date', 'Time of Entry'])
        df_filtered['Part_ก่อนหน้า'] = df_filtered.groupby('Machine')['Part'].shift(1)
        df_filtered['Is_Setup'] = (df_filtered['Part'] != df_filtered['Part_ก่อนหน้า']) & (df_filtered['Part_ก่อนหน้า'].notna())

        # Group By รวมยอด (วัน, เครื่อง, พาร์ท) และนับจำนวนครั้ง Setup
        df_grouped = df_filtered.groupby(['วันที่ผลิต', 'Machine', 'Part'], as_index=False).agg(
            actual_qty=('actual_qty', 'sum'),
            Setup_Count=('Is_Setup', 'sum')
        )

        # นำเป้าการผลิตมาเชื่อม
        df_target = df_target.rename(columns={'Material': 'Part', 'cap/day': 'เป้าต่อวัน(3กะ)'})
        df_target.columns = df_target.columns.str.strip()
        
        df_final = pd.merge(df_grouped, df_target, on='Part', how='left')
        df_final['เป้าต่อวัน(3กะ)'] = df_final['เป้าต่อวัน(3กะ)'].fillna(0)

        return df_final, None
    except Exception as e:
        return None, f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}"

# --- UI หลักของ Dashboard ---
st.title("🏭 Press Daily Production Dashboard")

# 1. แถบเมนูด้านซ้าย: อัปโหลดไฟล์
st.sidebar.header("📂 อัปโหลดไฟล์ข้อมูล")
uploaded_file = st.sidebar.file_uploader("เลือกไฟล์ Excel (เฉพาะไฟล์ data.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df, error = load_and_prep_data(uploaded_file)
    
    if error:
        st.error(error)
    else:
        st.sidebar.success("✅ โหลดข้อมูลสำเร็จ!")
        
        # --- แสดงช่วงวันที่ของ Data Base ---
        min_date = df['วันที่ผลิต'].min().strftime('%d/%m/%Y')
        max_date = df['วันที่ผลิต'].max().strftime('%d/%m/%Y')
        st.markdown(f"**📅 ข้อมูลตั้งแต่วันที่ {min_date} ถึง {max_date} (ตาม Data Base)**")
        
        # --- 2. ตั้งค่า O.E.E. และ เวลา Setup ---
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 ตั้งค่าประสิทธิภาพ")
        
        oee_val = st.sidebar.number_input("1. ค่า O.E.E. (1-100%)", min_value=1, max_value=100, value=100, step=1)
        oee_multiplier = oee_val / 100.0
        
        setup_hours = st.sidebar.number_input("2. เวลา Setup เปลี่ยน Part (ชั่วโมง)", min_value=0.0, max_value=24.0, value=4.0, step=0.5)
        setup_deduct_ratio = setup_hours / 24.0
        
        # --- 3. จัดการตั้งค่า กะการทำงาน (ผสมกันได้) ---
        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ ตั้งค่ากะการผลิต (แบบผสม)")
        st.sidebar.caption("ระบบจะใช้เป้าที่น้อยที่สุด หากมีการตั้งค่าซ้อนทับกัน")

        shift_mapping = {
            "3 กะ (เป้า 100%)": 1.0,
            "2 กะ (เป้า 67%)": 0.67,
            "1.5 กะ (เป้า 50%)": 0.5
        }

        # [ตั้งค่าแบบที่ 1: รายวัน]
        with st.sidebar.expander("📅 1. ตั้งค่ากะรายวัน (By Date)", expanded=False):
            default_shift_date = st.selectbox("กะมาตรฐาน (สำหรับทุกวัน):", list(shift_mapping.keys()), index=0, key='def_date')
            unique_dates = sorted(df['วันที่ผลิต'].unique())
            shift_date_df = pd.DataFrame({'วันที่': unique_dates, 'กะการทำงาน': [default_shift_date] * len(unique_dates)})

            st.write("แก้ไขกะเฉพาะบางวัน:")
            edited_shift_date_df = st.data_editor(
                shift_date_df,
                column_config={
                    "วันที่": st.column_config.DateColumn("วันที่", disabled=True, format="DD/MM/YYYY"),
                    "กะการทำงาน": st.column_config.SelectboxColumn("กะการทำงาน", options=list(shift_mapping.keys()), required=True)
                },
                hide_index=True, use_container_width=True, key='edit_date'
            )

            shift_multiplier_date = {}
            for index, row in edited_shift_date_df.iterrows():
                shift_multiplier_date[row['วันที่']] = shift_mapping[row['กะการทำงาน']]
            
            df['ตัวคูณกะ_Date'] = df['วันที่ผลิต'].map(shift_multiplier_date)

        # [ตั้งค่าแบบที่ 2: รายเครื่องจักร]
        with st.sidebar.expander("🚜 2. ตั้งค่ากะรายเครื่องจักร (By Machine)", expanded=False):
            default_shift_machine = st.selectbox("กะมาตรฐาน (สำหรับทุกเครื่อง):", list(shift_mapping.keys()), index=0, key='def_mac')
            unique_machines = sorted(df['Machine'].unique())
            shift_machine_df = pd.DataFrame({'Machine': unique_machines, 'กะการทำงาน': [default_shift_machine] * len(unique_machines)})

            st.write("แก้ไขกะเฉพาะบางเครื่อง:")
            edited_shift_machine_df = st.data_editor(
                shift_machine_df,
                column_config={
                    "Machine": st.column_config.TextColumn("ชื่อเครื่องจักร", disabled=True),
                    "กะการทำงาน": st.column_config.SelectboxColumn("กะการทำงาน", options=list(shift_mapping.keys()), required=True)
                },
                hide_index=True, use_container_width=True, key='edit_mac'
            )

            shift_multiplier_machine = {}
            for index, row in edited_shift_machine_df.iterrows():
                shift_multiplier_machine[row['Machine']] = shift_mapping[row['กะการทำงาน']]
            
            df['ตัวคูณกะ_Machine'] = df['Machine'].map(shift_multiplier_machine)

        # --- คำนวณเป้าหมายที่ปรับแล้ว ---
        df['ตัวคูณกะสุทธิ'] = df[['ตัวคูณกะ_Date', 'ตัวคูณกะ_Machine']].min(axis=1)

        # 1. เป้าหมายตั้งต้น * ตัวคูณกะสุทธิ * OEE
        df['เป้าหมายก่อนหักSetup'] = df['เป้าต่อวัน(3กะ)'] * df['ตัวคูณกะสุทธิ'] * oee_multiplier
        
        # 2. คำนวณยอดลดเป้าจาก Setup
        df['ยอดลดเป้าSetup'] = (df['เป้าต่อวัน(3กะ)'] * setup_deduct_ratio) * df['Setup_Count']
        
        # 3. เป้าหมายสุดท้าย
        df['เป้าหมายที่ปรับแล้ว'] = df['เป้าหมายก่อนหักSetup'] - df['ยอดลดเป้าSetup']
        df['เป้าหมายที่ปรับแล้ว'] = df['เป้าหมายที่ปรับแล้ว'].clip(lower=0)
        
        # คำนวณ % Achieve
        df['% Achieve'] = (df['actual_qty'] / df['เป้าหมายที่ปรับแล้ว']) * 100
        df['% Achieve'] = df['% Achieve'].round(2).fillna(0) 

        # --- 4. ตัวกรองข้อมูล (Filters) ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 ตัวกรองข้อมูล (Filters)")
        
        selected_machines = st.sidebar.multiselect("เลือกเครื่องจักร (Machine)", options=sorted(df['Machine'].unique()), default=[])
        if selected_machines:
            df = df[df['Machine'].isin(selected_machines)]
            
        selected_parts = st.sidebar.multiselect("เลือกชิ้นงาน (Part)", options=sorted(df['Part'].unique()), default=[])
        if selected_parts:
            df = df[df['Part'].isin(selected_parts)]

        # --- 5. แสดงผลตัวชี้วัด (Metrics) แบบดั้งเดิม (ตามที่ต้องการ) ---
        st.markdown("---")
        total_actual = df['actual_qty'].sum()
        total_target_original = df['เป้าต่อวัน(3กะ)'].sum()
        total_target = df['เป้าหมายที่ปรับแล้ว'].sum()
        
        overall_achieve = (total_actual / total_target * 100) if total_target > 0 else 0

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("ยอดผลิตจริง (Actual)", f"{total_actual:,.0f} Pcs")
        with col2:
            st.metric("เป้า 100% (Original Target)", f"{total_target_original:,.0f} Pcs")
        with col3:
            st.metric("เป้าหลังหัก กะ/OEE/Setup", f"{total_target:,.0f} Pcs")
        with col4:
            st.metric("ประสิทธิภาพรวม (% Achieve)", f"{overall_achieve:.2f}%")

        # --- 6. กราฟ 2 แกน (Plotly Dual-Axis Chart) ---
        st.subheader("📊 เปรียบเทียบยอดผลิตจริง กับ เป้าหมาย (พร้อม % Achieve)")
        
        daily_summary = df.groupby('วันที่ผลิต').agg({
            'actual_qty': 'sum',
            'เป้าหมายที่ปรับแล้ว': 'sum'
        }).reset_index()
        
        daily_summary['% Achieve'] = (daily_summary['actual_qty'] / daily_summary['เป้าหมายที่ปรับแล้ว'] * 100).fillna(0).round(2)
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig.add_trace(
            go.Bar(x=daily_summary['วันที่ผลิต'], y=daily_summary['actual_qty'], name="ยอดผลิตจริง (Actual)", marker_color='#1f77b4'),
            secondary_y=False,
        )
        fig.add_trace(
            go.Bar(x=daily_summary['วันที่ผลิต'], y=daily_summary['เป้าหมายที่ปรับแล้ว'], name="เป้าหมาย (Target)", marker_color='#ff7f0e'),
            secondary_y=False,
        )
        fig.add_trace(
            go.Scatter(x=daily_summary['วันที่ผลิต'], y=daily_summary['% Achieve'], name="% Achieve", mode='lines+markers', line=dict(color='red', width=3), marker=dict(size=8)),
            secondary_y=True,
        )
        
        fig.update_layout(
            barmode='group',
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=0, r=0, t=30, b=0)
        )
        fig.update_yaxes(title_text="จำนวนชิ้นงาน (Pcs)", secondary_y=False)
        fig.update_yaxes(title_text="ประสิทธิภาพ (% Achieve)", secondary_y=True, ticksuffix="%")
        
        st.plotly_chart(fig, use_container_width=True)

        # --- 7. ตารางข้อมูลดิบ และปุ่ม Export ---
        st.subheader("📋 รายละเอียดข้อมูลการผลิต (Data Table)")
        
        display_df = df[['วันที่ผลิต', 'Machine', 'Part', 'actual_qty', 'เป้าต่อวัน(3กะ)', 'ตัวคูณกะสุทธิ', 'Setup_Count', 'เป้าหมายที่ปรับแล้ว', '% Achieve']]
        display_df = display_df.sort_values(by=['วันที่ผลิต', 'Machine'], ascending=[False, True])
        
        st.dataframe(
            display_df, 
            use_container_width=True,
            column_config={
                "actual_qty": st.column_config.NumberColumn("ยอดผลิตจริง"),
                "เป้าต่อวัน(3กะ)": st.column_config.NumberColumn("เป้า 3 กะ"),
                "ตัวคูณกะสุทธิ": st.column_config.NumberColumn("อัตราส่วนกะสุทธิ"),
                "Setup_Count": st.column_config.NumberColumn("จำนวนครั้งเปลี่ยน Part"),
                "เป้าหมายที่ปรับแล้ว": st.column_config.NumberColumn("เป้าสุทธิ"),
                "% Achieve": st.column_config.ProgressColumn("% เทียบเป้า", format="%.2f%%", min_value=0, max_value=150)
            }
        )
        
        csv_data = display_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 ดาวน์โหลดข้อมูล (Export to CSV)",
            data=csv_data,
            file_name=f"Production_Data_Export.csv",
            mime="text/csv"
        )

else:
    st.info("👈 กรุณาใช้แถบเมนูด้านซ้ายเพื่ออัปโหลดไฟล์ **data.xlsx** ระบบจะทำการสร้าง Dashboard ให้ทันทีครับ")