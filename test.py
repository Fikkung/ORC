from PIL import Image, ImageEnhance, ImageTk
import pytesseract
import cv2
from tkinter import Tk, Button, Label, filedialog, messagebox, Canvas, Frame
from tkinter.font import Font
import os
import logging
import re

# ตั้งค่า logging
logging.basicConfig(filename='ocr_log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# ตั้งค่า TESSDATA_PREFIX เพื่อชี้ไปยังไดเรกทอรี tessdata
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

# ตั้งเส้นทางไปยังโปรแกรม Tesseract
tessdata_dir_config = r'--tessdata-dir "C:\Program Files\Tesseract-OCR\tessdata"'

class OCRApp:
    
    def __init__(self, master):
        self.master = master
        master.title("OCR Application")
        master.geometry("800x800")
        master.configure(bg="#f0f0f0")
        master.state('zoomed')  # ตั้งหน้าต่างให้ขยายเต็มจอ

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

        self.canvas = Canvas(master, bg="#f0f0f0")
        self.canvas.pack(fill="both", expand=True)

        self.file_path = None
        self.image = None
        self.crop_box = None
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        self.img_size = None
        self.display_size = None
        self.image_id = None

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

        master.bind("<Escape>", self.exit_fullscreen)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff")])
        if self.file_path:
            self.display_image(self.file_path)

    def display_image(self, file_path):
        self.image = Image.open(file_path)
        self.img_size = self.image.size

        # ขนาดของ Canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = self.img_size

        # คำนวณขนาดที่ต้องการให้พอดีกับ Canvas
        if img_width > canvas_width or img_height > canvas_height:
            ratio = min(canvas_width / img_width, canvas_height / img_height)
            self.display_size = (int(img_width * ratio), int(img_height * ratio))
        else:
            self.display_size = (img_width, img_height)

        # ปรับขนาดภาพตามขนาดที่แสดง
        self.image = self.image.resize(self.display_size, Image.Resampling.LANCZOS)
        img_tk = ImageTk.PhotoImage(self.image)

        # เคลียร์ Canvas และแสดงภาพ
        self.canvas.delete("all")
        self.image_id = self.canvas.create_image(0, 0, anchor="nw", image=img_tk)
        self.canvas.image = img_tk

        # คำนวณตำแหน่งที่ต้องการให้ภาพอยู่กลาง Canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        image_width, image_height = self.display_size
        x_center = (canvas_width - image_width) // 2
        y_center = (canvas_height - image_height) // 2

        # เคลื่อนย้ายภาพไปยังตำแหน่งกลาง
        self.canvas.coords(self.image_id, x_center, y_center)

    def on_mouse_down(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def on_mouse_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def on_mouse_up(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        
        self.crop_box = (min(self.start_x, end_x), min(self.start_y, end_y), max(self.start_x, end_x), max(self.start_y, end_y))

    def validate_filename(self, filename):
        return re.match(r'^[^\\/*?:"<>|]+$', filename) is not None

    def start_ocr(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select an image file first.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if not save_path:
            messagebox.showerror("Error", "Please specify a file path to save the text file.")
            return

        filename = os.path.basename(save_path)
        if not self.validate_filename(filename):
            messagebox.showerror("Error", "The file name contains invalid characters. Please use a different name.")
            return

        try:
            # โหลดภาพด้วย OpenCV และแปลงเป็น RGB
            img = cv2.imread(self.file_path)
            if img is None:
                raise FileNotFoundError('The image file was not found.')
            logging.info('Image loaded successfully.')

            img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            logging.info('Image converted to RGB format.')

            # ถ้ามีการเลือกพื้นที่
            if self.crop_box:
                left, top, right, bottom = map(int, self.crop_box)
                # แปลงพิกัดให้สอดคล้องกับขนาดภาพที่แสดงบน Canvas
                canvas_width = self.canvas.winfo_width()
                canvas_height = self.canvas.winfo_height()
                image_width, image_height = self.display_size
                x_center = (canvas_width - image_width) // 2
                y_center = (canvas_height - image_height) // 2
                ratio_x = self.img_size[0] / image_width
                ratio_y = self.img_size[1] / image_height
                left = int((left - x_center) * ratio_x)
                top = int((top - y_center) * ratio_y)
                right = int((right - x_center) * ratio_x)
                bottom = int((bottom - y_center) * ratio_y)




                img_rgb = img_rgb[top:bottom, left:right]
                logging.info('Image cropped based on selection.')

            # ปรับปรุงภาพก่อนทำ OCR
            image_pil = Image.fromarray(img_rgb)
            enhancer = ImageEnhance.Contrast(image_pil)
            image_pil = enhancer.enhance(2)  # ปรับคอนทราสต์
            enhancer = ImageEnhance.Sharpness(image_pil)
            image_pil = enhancer.enhance(2)  # ปรับความคมชัด
            image_pil = image_pil.convert('L')  # แปลงภาพเป็นขาวดำ

            # แสดงภาพหลังการปรับปรุง
            image_pil.show(title="Processed Image")

            # ทำ OCR
            text = pytesseract.image_to_string(image_pil, config=tessdata_dir_config, lang='tha+eng')
            logging.info('Text extracted from image successfully.')
            print(text)

            with open(save_path, 'w', encoding='utf-8') as file:
                file.write(text)
            logging.info(f'Text saved to {save_path}')
            messagebox.showinfo("Success", f"Text extracted and saved to {save_path}")

        except FileNotFoundError as fnf_error:
            logging.error(f"FileNotFoundError: {fnf_error}")
            messagebox.showerror("Error", f"An error occurred: {fnf_error}")
            print(f"An error occurred: {fnf_error}")

        except pytesseract.TesseractError as tess_error:
            logging.error(f"TesseractError: {tess_error}")
            messagebox.showerror("Error", f"An OCR error occurred: {tess_error}")
            print(f"An OCR error occurred: {tess_error}")

        except Exception as e:
            logging.error(f"Exception: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
            print(f"An error occurred: {e}")

    def clear_data(self):
        self.file_path = None
        self.image = None
        self.crop_box = None
        self.canvas.delete("all")
        messagebox.showinfo("Clear Data", "All data has been cleared.")

    def exit_fullscreen(self, event=None):
        self.master.attributes("-fullscreen", False)

if __name__ == "__main__":
    root = Tk()
    app = OCRApp(root)
    root.mainloop()
