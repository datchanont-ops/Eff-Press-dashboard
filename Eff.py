import os

# --- ส่วนอัปโหลดไฟล์ ---
st.sidebar.header("📂 อัปโหลดข้อมูล")
uploaded_file = st.sidebar.file_uploader("เลือกไฟล์ Excel", type=["xlsx", "xls"])

# กำหนดชื่อไฟล์ที่จะใช้เป็นไฟล์ Default ในระบบ
default_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'default_data.xlsx')

# --- 📌 ปุ่มบันทึกไฟล์เป็นค่าเริ่มต้น ---
if uploaded_file is not None:
    if st.sidebar.button("💾 บันทึกไฟล์นี้เป็นค่าเริ่มต้น (Save Default)", use_container_width=True):
        # บันทึกไฟล์ที่อัปโหลดลงไปในเครื่อง
        with open(default_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ บันทึกเป็นไฟล์เริ่มต้นเรียบร้อย! คราวหน้าไม่ต้องอัปโหลดซ้ำแล้วครับ")
        st.rerun() # รีเฟรชหน้าเว็บ 1 ครั้งเพื่ออัปเดตสถานะ

# --- ระบบเลือกไฟล์อัตโนมัติ (Fallback) ---
data_source = None
if uploaded_file is not None:
    data_source = uploaded_file
    st.sidebar.caption("🟢 กำลังแสดงผลจาก: **ไฟล์ที่เพิ่งอัปโหลด**")
elif os.path.exists(default_file_path):
    data_source = default_file_path
    st.sidebar.caption("📌 กำลังแสดงผลจาก: **ไฟล์ที่บันทึกไว้ในระบบล่าสุด**")

# ตรวจสอบว่ามีไฟล์ให้ทำงานต่อหรือไม่
if data_source is None:
    st.info("👋 กรุณาอัปโหลดไฟล์เพื่อเริ่มต้นใช้งานครับ")
    st.stop()

# (จากนั้นนำ data_source ไปให้ Pandas อ่านข้อมูลต่อได้เลย)
# df = pd.read_excel(data_source)