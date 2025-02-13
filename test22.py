from PIL import Image, ImageEnhance, ImageTk
import pytesseract
import cv2
from tkinter import Tk, Button, Label, filedialog, messagebox, Canvas, Frame, Text
import os
import logging
import re
import pyperclip

# ตั้งค่า logging
logging.basicConfig(filename='ocr_log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# ตรวจสอบการติดตั้ง Tesseract
TESSERACT_PATH = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
if not os.path.exists(TESSERACT_PATH):
    messagebox.showerror("Error", "Tesseract OCR not found! Please install Tesseract-OCR.")
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

tessdata_dir_config = r'--tessdata-dir "C:\\Program Files\\Tesseract-OCR\\tessdata"'

class OCRApp:
    def __init__(self, master):
        self.master = master
        master.title("OCR Application")
        master.geometry("900x700")
        master.configure(bg="#121212")  # Dark Mode

        self.label = Label(master, text="Select an image file for OCR processing:", fg="white", bg="#121212")
        self.label.pack(pady=10)

        self.button_frame = Frame(master, bg="#121212")
        self.button_frame.pack(pady=10)

        self.select_button = Button(self.button_frame, text="Select File", command=self.select_file, bg="#4CAF50", fg="white")
        self.select_button.pack(side="left", padx=5)

        self.start_button = Button(self.button_frame, text="Start OCR", command=self.start_ocr, bg="#2196F3", fg="white")
        self.start_button.pack(side="left", padx=5)
        
        self.copy_button = Button(self.button_frame, text="Copy Text", command=self.copy_to_clipboard, bg="#FF9800", fg="white")
        self.copy_button.pack(side="left", padx=5)

        self.clear_button = Button(self.button_frame, text="Clear", command=self.clear_data, bg="#F44336", fg="white")
        self.clear_button.pack(side="right", padx=5)

        self.canvas = Canvas(master, bg="#1E1E1E")
        self.canvas.pack(fill="both", expand=True)

        self.text_output = Text(master, height=10, bg="#1E1E1E", fg="white")
        self.text_output.pack(fill="both", expand=True, padx=10, pady=10)

        self.file_path = None
        self.image = None
        self.ocr_result = ""

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff")])
        if self.file_path:
            self.display_image(self.file_path)

    def display_image(self, file_path):
        self.image = Image.open(file_path)
        self.image.thumbnail((700, 500), Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(self.image)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=img_tk)
        self.canvas.image = img_tk

    def start_ocr(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select an image file first.")
            return
        try:
            img = cv2.imread(self.file_path)
            img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            img_thresh = cv2.adaptiveThreshold(img_gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2)
            image_pil = Image.fromarray(img_thresh)

            self.ocr_result = pytesseract.image_to_string(image_pil, config=tessdata_dir_config, lang='tha+eng')
            self.text_output.delete("1.0", "end")
            self.text_output.insert("1.0", self.ocr_result)

            messagebox.showinfo("Success", "OCR completed!")
        except Exception as e:
            logging.error(f"OCR Error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")

    def copy_to_clipboard(self):
        pyperclip.copy(self.ocr_result)
        messagebox.showinfo("Copied", "Text copied to clipboard!")

    def clear_data(self):
        self.file_path = None
        self.image = None
        self.ocr_result = ""
        self.canvas.delete("all")
        self.text_output.delete("1.0", "end")
        messagebox.showinfo("Cleared", "All data has been cleared.")

if __name__ == "__main__":
    root = Tk()
    app = OCRApp(root)
    root.mainloop()
