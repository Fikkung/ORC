from PIL import Image
import pytesseract
import cv2
import pandas as pd
import re
import os

# ตั้งค่า Tesseract OCR
TESSERACT_PATH = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
tessdata_dir_config = r'--tessdata-dir "C:\\Program Files\\Tesseract-OCR\\tessdata"'

def preprocess_image(image_path):
    """ ปรับแต่งภาพให้ OCR อ่านง่ายขึ้น """
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    img = cv2.GaussianBlur(img, (1, 1), 0)
    return img

def extract_important_data(text):
    """ ดึงเฉพาะข้อมูลที่ต้องการ """
    data = {
        "เลขใบกำกับภาษี": re.findall(r'เลขที่ใบกำกับภาษี[:\s]*([A-Za-z0-9-]+)', text),
        "เลขประจำตัวผู้เสียภาษี": re.findall(r'เลขประจำตัวผู้เสียภาษี[:\s]*([\d]+)', text),
        "สินค้าที่ยกเว้นภาษีมูลค่าเพิ่ม": re.findall(r'สินค้าที่ยกเว้นภาษีมูลค่าเพิ่ม[:\s]*([\d,.]+)', text),
        "สินค้ารวมภาษีมูลค่าเพิ่ม": re.findall(r'สินค้ารวมภาษีมูลค่าเพิ่ม[:\s]*([\d,.]+)', text),
        "ภาษีมูลค่าเพิ่ม": re.findall(r'ภาษีมูลค่าเพิ่ม[:\s]*([\d,.]+)', text),
        "สินค้าที่เสียภาษีมูลค่าเพิ่ม": re.findall(r'สินค้าที่เสียภาษีมูลค่าเพิ่ม[:\s]*([\d,.]+)', text),
        "จำนวนเงินรวมทั้งสิ้น": re.findall(r'จำนวนเงินรวมทั้งสิ้น[:\s]*([\d,.]+)', text),
    }

    # แปลงเป็นค่าเดียว
    for key in data.keys():
        data[key] = data[key][0] if data[key] else "ไม่พบข้อมูล"

    # ดึงรายการสินค้า
    items = re.findall(r'(\d+)\s+([\wก-๙\s]+)\s+([\d,.]+)\s+([\d,.]+)', text)
    data["รายการสินค้า"] = [{"ลำดับ": i[0], "ชื่อสินค้า": i[1], "หน่วยละ": i[2], "จำนวนเงิน": i[3]} for i in items]

    return data

def save_to_excel(data, output_file="invoice_data.xlsx"):
    """ บันทึกข้อมูลลงไฟล์ Excel """
    df_main = pd.DataFrame([data])
    df_items = pd.DataFrame(data["รายการสินค้า"])

    with pd.ExcelWriter(output_file) as writer:
        df_main.to_excel(writer, sheet_name="ข้อมูลหลัก", index=False)
        df_items.to_excel(writer, sheet_name="รายการสินค้า", index=False)

def process_invoice(image_path):
    """ ประมวลผล OCR และบันทึกข้อมูล """
    processed_img = preprocess_image(image_path)
    text = pytesseract.image_to_string(processed_img, config=tessdata_dir_config, lang="tha+eng")
    
    extracted_data = extract_important_data(text)
    save_to_excel(extracted_data)
    print(f"บันทึกข้อมูลลงไฟล์ Excel เรียบร้อย: invoice_data.xlsx")

# ใช้งาน OCR กับไฟล์ภาพ
image_path = "S__9265180.jpg"
process_invoice(image_path)
