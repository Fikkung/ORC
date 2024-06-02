import os
import pytesseract
import cv2
import pandas as pd
from tkinter import Tk, Button, Label, filedialog, messagebox, Entry, Frame
from tkinter.font import Font
from PIL import Image, ImageTk
import logging
import re

# ตั้งค่า logging
logging.basicConfig(filename='ocr_log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Set the TESSDATA_PREFIX environment variable to point to the tessdata directory
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
tessdata_dir_config = r'--tessdata-dir "C:\Program Files\Tesseract-OCR\tessdata"'

class OCRApp:
    
    def __init__(self, master):
        self.master = master
        master.title("OCR Application")
        master.geometry("600x600")
        master.configure(bg="#f0f0f0")
        master.state('zoomed')  # Set window to maximized state

        self.title_font = Font(family="Helvetica", size=16, weight="bold")
        self.label_font = Font(family="Helvetica", size=12)
        self.button_font = Font(family="Helvetica", size=10)

        self.title_label = Label(master, text="OCR Application", font=self.title_font, bg="#f0f0f0")
        self.title_label.pack(pady=10)

        self.label = Label(master, text="Select an image file for OCR processing:", font=self.label_font, bg="#f0f0f0")
        self.label.pack(pady=5)

        self.button_frame = Frame(master, bg="#f0f0f0")
        self.button_frame.pack(pady=10)

        self.select_button = Button(self.button_frame, text="Select File", font=self.button_font, command=self.select_file, bg="#4CAF50", fg="white", padx=10, pady=5)
        self.select_button.pack(side="left", padx=5)

        self.start_button = Button(self.button_frame, text="Start OCR", font=self.button_font, command=self.start_ocr, bg="#2196F3", fg="white", padx=10, pady=5)
        self.start_button.pack(side="left", padx=5)

        self.clear_button = Button(self.button_frame, text="Clear Data", font=self.button_font, command=self.clear_data, bg="#F44336", fg="white", padx=10, pady=5)
        self.clear_button.pack(side="right", padx=5)

        self.image_label = Label(master, bg="#f0f0f0")
        self.image_label.pack(pady=10)

        self.file_path = None
        self.is_fullscreen = True

        master.bind("<Escape>", self.exit_fullscreen)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff")])
        if self.file_path:
            self.display_image(self.file_path)

    def display_image(self, file_path):
        img = Image.open(file_path)

        # กำหนดขนาดภาพที่ต้องการ
        max_width, max_height = 400, 400
        img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)

        img = ImageTk.PhotoImage(img)
        self.image_label.config(image=img)
        self.image_label.image = img

    def validate_filename(self, filename):
        # ใช้ regex เพื่อตรวจสอบว่า filename ไม่มีตัวอักษรที่ไม่อนุญาต
        if re.match(r'^[^\\/*?:"<>|]+$', filename):
            return True
        return False

    def start_ocr(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select an image file first.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            messagebox.showerror("Error", "Please specify a file path to save the Excel file.")
            return

        # แยกชื่อไฟล์จากเส้นทางและตรวจสอบชื่อไฟล์
        filename = os.path.basename(save_path)
        if not self.validate_filename(filename):
            messagebox.showerror("Error", "The file name contains invalid characters. Please use a different name.")
            return

        try:
            # อ่านภาพโดยใช้ OpenCV
            img = cv2.imread(self.file_path)
            if img is None:
                raise FileNotFoundError('The image file was not found.')
            logging.info('Image loaded successfully.')

            # แปลงภาพเป็นรูปแบบ RGB
            img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            logging.info('Image converted to RGB format.')

            # ดึงข้อความจากภาพโดยใช้ Tesseract พร้อมกำหนดให้ใช้ภาษาไทยและภาษาอังกฤษ
            text = pytesseract.image_to_string(img_rgb, config=tessdata_dir_config, lang='tha+eng')
            logging.info('Text extracted from image successfully.')
            print(text)

            # แยกข้อความแต่ละบรรทัดและสร้าง DataFrame
            lines = text.split('\n')
            df = pd.DataFrame({'Extracted Text': lines, 'New Text': ''})

            # บันทึกข้อความลงในไฟล์ Excel
            df.to_excel(save_path, index=False)
            logging.info(f'Text saved to {save_path}')
            messagebox.showinfo("Success", f"Text extracted and saved to {save_path}")

        except FileNotFoundError as fnf_error:
            logging.error(f"FileNotFoundError: {fnf_error}")
            messagebox.showerror("Error", f"An error occurred: {fnf_error}")
            print(f"An error occurred: {fnf_error}")

        except Exception as e:
            logging.error(f"Exception: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
            print(f"An error occurred: {e}")

    def clear_data(self):
        self.file_path = None
        self.image_label.config(image='')
        self.image_label.image = None
        messagebox.showinfo("Clear Data", "All data has been cleared.")

    def exit_fullscreen(self, event=None):
        self.is_fullscreen = False
        self.master.attributes("-fullscreen", False)

if __name__ == "__main__":
    root = Tk()
    app = OCRApp(root)
    root.mainloop()
