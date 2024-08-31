import easyocr
import cv2
import numpy as np
from transformers import pipeline
import openai
import os
from dotenv import load_dotenv

# โหลดตัวแปรสภาพแวดล้อมจากไฟล์ .env
load_dotenv()

# ฟังก์ชันสำหรับการทำ Thresholding
def get_grayscale(image):
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

def thresholding(image):
    _, thresh = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return thresh

def opening(image):
    kernel = np.ones((5,5),np.uint8)
    return cv2.morphologyEx(image, cv2.MORPH_OPEN, kernel)

def canny(image):
    return cv2.Canny(image, 100, 200)

# ฟังก์ชันสำหรับการทำ Preprocessing ของภาพ
def preprocess_image(image_path):
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"The file {image_path} does not exist.")
    
    # โหลดภาพ
    img = cv2.imread(image_path)
    
    if img is None:
        raise ValueError(f"Could not read the image file {image_path}. Check the file format and path.")
    
    # ทำการแปลงภาพเป็น grayscale
    grey = get_grayscale(img)
    
    # ทำการ Thresholding
    thresh = thresholding(grey)
    cv2.imwrite('thresh.png', thresh)
    
    # ทำการ Opening
    opening_img = opening(grey)
    cv2.imwrite('opening.png', opening_img)
    
    # ทำการ Canny Edge Detection
    canny_img = canny(grey)
    cv2.imwrite('canny.png', canny_img)
    
    return 'thresh.png', 'opening.png', 'canny.png'

# ฟังก์ชันสำหรับการดึงข้อความจากภาพ
def extract_text_from_image(image_path, languages=['en', 'th']):
    # สร้าง Reader สำหรับ EasyOCR
    reader = easyocr.Reader(languages)
    
    # อ่านข้อความจากภาพ
    result = reader.readtext(image_path, detail=0, paragraph=True, contrast_ths=0.1, adjust_contrast=0.5)
    
    # รวมข้อความทั้งหมดเป็นสตริงเดียว
    text = ' '.join(result)
    
    return text

# ฟังก์ชันสำหรับการปรับปรุงข้อความด้วย AI (ใช้ Hugging Face's Transformers)
def improve_text_with_ai(text):
    corrector = pipeline("summarization", model="facebook/bart-large-cnn")
    
    # ตัวอย่างการใช้ pipeline เพื่อปรับปรุงข้อความ
    improved_text = corrector(text, max_length=512, min_length=10, do_sample=False)
    
    return improved_text[0]['summary_text']

# ฟังก์ชันสำหรับการปรับปรุงข้อความด้วย GPT ของ OpenAI
def improve_text_with_gpt(text):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",  # ใช้โมเดลใหม่
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"Please correct the following OCR text:\n\n{text}"}
        ],
        max_tokens=512,
        temperature=0.3,
    )
    
    improved_text = response.choices[0].message['content'].strip()
    return improved_text

# ฟังก์ชันสำหรับการบันทึกข้อความลงในไฟล์
def save_text_to_file(text, filename):
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(text)

# ตั้งค่า API Key ของ OpenAI จากไฟล์ .env
openai.api_key = os.getenv('OPENAI_API_KEY')

# เส้นทางภาพต้นฉบับ
image_path = r'D:\test.png'  # เส้นทางของภาพที่บันทึกไว้ในเครื่อง

# ขั้นตอนการทำ Preprocessing
try:
    thresh_path, opening_path, canny_path = preprocess_image(image_path)
    
    # ดึงข้อความจากภาพที่ถูก Preprocessing แล้ว
    extracted_text = extract_text_from_image(thresh_path, languages=['en', 'th'])  # ปรับตามภาษาที่ต้องการ

    print("ข้อความที่ดึงได้จากภาพก่อนปรับปรุง:")
    print(extracted_text)

    # ปรับปรุงข้อความด้วย AI (ใช้ Hugging Face's Transformers)
    improved_text_ai = improve_text_with_ai(extracted_text)

    print("\nข้อความที่ปรับปรุงด้วย AI:")
    print(improved_text_ai)

    # ปรับปรุงข้อความด้วย GPT ของ OpenAI
    improved_text_gpt = improve_text_with_gpt(extracted_text)

    print("\nข้อความที่ปรับปรุงด้วย GPT:")
    print(improved_text_gpt)

    # บันทึกข้อความที่ปรับปรุงแล้วลงในไฟล์
    save_text_to_file(improved_text_gpt, 'improved_text_gpt.txt')

except Exception as e:
    print(f"An error occurred: {e}")
