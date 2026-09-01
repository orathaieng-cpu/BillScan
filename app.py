import streamlit as st
from google import genai
import pandas as pd
from PIL import Image
import io
import json
import gspread
from google.oauth2.service_account import Credentials

# ตั้งค่าหน้าตาของ Web App
st.set_page_config(
    page_title="ระบบอ่านบิล & สกัดข้อมูล",
    page_icon="🧾",
    layout="centered"
)

# Style ตกแต่งเพิ่มเติม
st.markdown("""
    <style>
    .main-title { text-align: center; color: #1F497D; font-weight: bold; }
    .sub-title { text-align: center; color: #555555; margin-bottom: 30px; }
    .stButton>button { width: 100%; background-color: #1F497D; color: white; border-radius: 8px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-title'>🧾 Receipt OCR to Excel / Google Sheets</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>ถ่ายรูปหรืออัปโหลดรูปบิลเพื่อสกัดข้อมูล ส่งเข้า Google Sheets และดาวน์โหลด Excel</p>", unsafe_allow_html=True)

# แถบข้างสำหรับใส่การตั้งค่า (Sidebar)
with st.sidebar:
    st.header("⚙️ ตั้งค่าระบบ")
    api_key = st.text_input("🔑 Gemini API Key", type="password", help="รับ API Key ฟรีได้จาก Google AI Studio")
    sheet_url = st.text_input("📊 Google Sheet URL (ถ้าต้องการเซฟลง Sheets)", help="วาง Link ของ Google Sheet ที่ต้องการให้บันทึกข้อมูล")
    st.info("💡 หากไม่ใส่ Google Sheet URL ระบบจะยังคงสกัดข้อมูลและให้ดาวน์โหลดไฟล์ Excel ได้ตามปกติครับ")

# ฟังก์ชันบันทึกลง Google Sheets
def save_to_google_sheet(df, target_url):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    client = gspread.authorize(creds)
    
    sheet = client.open_by_url(target_url).sheet1
    
    # ถ้าชีทยังไม่มีข้อมูล ให้สร้าง Header ก่อน
    if len(sheet.get_all_values()) == 0:
        sheet.append_row(df.columns.tolist())
        
    sheet.append_rows(df.values.tolist())

# ส่วนอัปโหลดรูปภาพ
uploaded_file = st.file_uploader("📷 เลือกรูปภาพ หรือ ถ่ายรูปบิล", type=["jpg", "jpeg", "png"])

if uploaded_file:
    image = Image.open(uploaded_file)
    st.image(image, caption="รูปภาพที่เลือก", use_container_width=True)
    
    if st.button("🚀 สกัดข้อมูลจากบิล"):
        if not api_key:
            st.error("⚠️ กรุณากรอก Gemini API Key ที่แถบเมนูด้านซ้ายก่อนครับ")
        else:
            with st.spinner("🧠 AI กำลังอ่านข้อมูลในบิล กรุณารอสักครู่..."):
                try:
                    client = genai.Client(api_key=api_key)
                    
                    prompt = """
                    จงอ่านข้อมูลบิลนี้อย่างละเอียด แล้วส่งผลลัพธ์กลับมาเป็น JSON array ของวัตถุ โดยมีโครงสร้างดังนี้:
                    [
                      {
                        "quantity": จำนวน (ตัวเลขหรือข้อความ),
                        "unit": "หน่วยนับ (เช่น เส้น, แผ่น)",
                        "description": "รายการสินค้า/บริการ",
                        "unit_price": ราคาต่อหน่วย (ตัวเลข),
                        "amount": จำนวนเงินรวม (ตัวเลข)
                      }
                    ]
                    ส่งคืนเฉพาะ JSON เพียวๆ เท่านั้น ห้ามมี Markdown หรือข้อความอื่น
                    """
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash-lite',
                        contents=[image, prompt]
                    )
                    
                    clean_text = response.text.replace("```json", "").replace("```", "").strip()
                    data = json.loads(clean_text)
                    
                    df = pd.DataFrame(data)
                    df.columns = ["จำนวน", "หน่วย", "รายการ", "หน่วยละ (บาท)", "จำนวนเงิน (บาท)"]
                    
                    st.success("✨ อ่านข้อมูลสำเร็จ!")
                    st.subheader("📋 ตารางข้อมูลที่อ่านได้")
                    st.dataframe(df, use_container_width=True)
                    
                    # 1. ปุ่ม Export เป็น Excel
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='บิลสินค้า')
                    excel_data = excel_buffer.getvalue()
                    
                    st.download_button(
                        label="📥 Export เป็นไฟล์ Excel (.xlsx)",
                        data=excel_data,
                        file_name="receipt_extracted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # 2. บันทึกลง Google Sheets (ถ้ามีการระบุ URL ไว้)
                    if sheet_url:
                        try:
                            save_to_google_sheet(df, sheet_url)
                            st.success("✅ บันทึกข้อมูลลง Google Sheets เรียบร้อยแล้ว!")
                        except Exception as s_err:
                            st.warning(f"⚠️ สกัดข้อมูลได้แล้ว แต่บันทึกลง Google Sheets ไม่สำเร็จ: {s_err}")
                            
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")
