import deepseek
import pandas as pd
from PIL import Image

def extract_text_from_image(image_path):
    """ใช้ DeepSeek OCR ดึงข้อความจากภาพ"""
    ocr_model = deepseek.OCR()  # โหลดโมเดล
    image = Image.open(image_path)
    results = ocr_model.recognize(image)
    
    extracted_text = "\n".join([res[1] for res in results])
    return extracted_text

def parse_invoice_data(text):
    """แยกข้อมูลที่ต้องการจากข้อความ OCR"""
    data = {
        "เลขใบกำกับภาษี": "",
        "เลขประจำตัวผู้เสียภาษี": "",
        "รายการ": [],
        "สินค้าที่ยกเว้นภาษีมูลค่าเพิ่ม": "",
        "สินค้ารวมภาษีมูลค่าเพิ่ม": "",
        "ภาษีมูลค่าเพิ่ม": "",
        "สินค้าที่เสียภาษีมูลค่าเพิ่ม": "",
        "จำนวนเงินรวมทั้งสิ้น": ""
    }
    
    lines = text.split("\n")
    for line in lines:
        if "เลขใบกำกับภาษี" in line:
            data["เลขใบกำกับภาษี"] = line.split(":")[-1].strip()
        elif "เลขประจำตัวผู้เสียภาษี" in line:
            data["เลขประจำตัวผู้เสียภาษี"] = line.split(":")[-1].strip()
        elif "จำนวนเงินรวมทั้งสิ้น" in line:
            data["จำนวนเงินรวมทั้งสิ้น"] = line.split(":")[-1].strip()
        # เพิ่มเงื่อนไขสำหรับข้อมูลอื่น ๆ ตามต้องการ
    
    return data

def save_to_excel(data, output_path="invoice_data.xlsx"):
    """บันทึกข้อมูลลงไฟล์ Excel"""
    df = pd.DataFrame([data])
    df.to_excel(output_path, index=False)
    print(f"บันทึกข้อมูลสำเร็จ: {output_path}")

if __name__ == "__main__":
    image_path = "ttt.jpg"  # ใส่พาธของไฟล์ภาพ
    text = extract_text_from_image(image_path)
    invoice_data = parse_invoice_data(text)
    save_to_excel(invoice_data)
