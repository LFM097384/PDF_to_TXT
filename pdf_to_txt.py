import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import sv_ttk
import darkdetect
from pathlib import Path
import os
from PIL import Image, ImageEnhance
import sys
import pytesseract
import uuid
import logging
import shutil

class PDFConverterGUI:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("PDF 转换工具")
        self.root.geometry("600x400")
        
        # 设置主题
        ctk.set_appearance_mode("system")  # 自动跟随系统主题
        ctk.set_default_color_theme("blue")
        sv_ttk.use_light_theme() if darkdetect.isLight() else sv_ttk.use_dark_theme()
        
        # 设置图标
        if sys.platform == "win32":
            self.root.iconbitmap("icon.ico")  # 需要准备一个icon.ico文件
            
        # 创建主框架
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="PDF 转 TXT 转换工具",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.pack(pady=(20, 30))

        # 选择文件按钮
        self.select_btn = ctk.CTkButton(
            self.main_frame,
            text="选择 PDF 文件 (可多选)",
            command=self.select_files,
            font=ctk.CTkFont(size=14),
            height=40,
            compound="left"  # 图标在左侧
        )
        self.select_btn.pack(pady=20)

        # 文件名标签
        self.file_label = ctk.CTkLabel(
            self.main_frame,
            text="未选择文件",
            font=ctk.CTkFont(size=12),
            wraplength=500
        )
        self.file_label.pack(pady=10)

        # 进度条框架
        self.progress_frame = ctk.CTkFrame(self.main_frame)
        self.progress_frame.pack(fill="x", padx=30, pady=20)
        
        # 进度条
        self.progress = ctk.CTkProgressBar(
            self.progress_frame,
            width=400,
            height=10
        )
        self.progress.pack(pady=10)
        self.progress.set(0)
        
        # 状态标签
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="准备就绪",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(pady=10)

        # 添加OCR选项
        self.ocr_var = tk.BooleanVar(value=False)
        self.ocr_checkbox = ctk.CTkCheckBox(
            self.main_frame,
            text="使用OCR识别(用于扫描版PDF)",
            variable=self.ocr_var,
            font=ctk.CTkFont(size=12)
        )
        self.ocr_checkbox.pack(pady=10)

        # 底部信息
        self.info_label = ctk.CTkLabel(
            self.main_frame,
            text="© 2024 PDF转换工具",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        self.info_label.pack(side="bottom", pady=10)

        # 创建临时目录
        self.temp_dir = os.path.join(os.path.expanduser('~'), '.pdf_converter_temp')
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)

        # 检查Tesseract是否已安装
        self.check_tesseract()
        
        # 设置日志
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger('PDFConverter')

    def check_tesseract(self):
        """检查Tesseract是否可用"""
        if not shutil.which('tesseract'):
            # 尝试设置Windows下的默认路径
            if sys.platform == "win32":
                if os.path.exists('C:\\Program Files\\Tesseract-OCR\\tesseract.exe'):
                    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
                elif os.path.exists('C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'):
                    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
                else:
                    messagebox.showwarning("警告", "未检测到Tesseract-OCR，OCR功能可能无法使用。\n请安装Tesseract-OCR后重试。")
                    self.logger.warning("Tesseract not found")
                    return False
        return True

    def select_files(self):
        pdf_paths = filedialog.askopenfilenames(
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if pdf_paths:
            self.file_label.configure(text=f"已选择: {len(pdf_paths)} 个文件")
            self.status_label.configure(text="正在转换...")
            self.convert_pdfs(pdf_paths)

    def preprocess_image(self, image):
        """预处理图片以提高OCR识别率"""
        # 转换为灰度图
        image = image.convert('L')
        
        # 增加对比度
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        
        # 调整亮度
        enhancer = ImageEnhance.Brightness(image)
        image = enhancer.enhance(1.5)
        
        return image

    def ocr_image(self, image):
        """对图片进行OCR识别"""
        try:
            # 预处理图片
            processed_image = self.preprocess_image(image)
            
            # 进行OCR识别
            text = pytesseract.image_to_string(processed_image, lang='chi_sim+eng')
            
            if not text.strip():
                self.logger.warning("OCR结果为空")
                return None
                
            return text
            
        except Exception as e:
            self.logger.error(f"OCR识别错误: {str(e)}")
            messagebox.showwarning("OCR警告", f"OCR识别出现错误：{str(e)}\n如果是tesseract错误，请确认已正确安装Tesseract-OCR")
            return None

    def cleanup_temp_files(self):
        """清理临时文件"""
        try:
            for file in os.listdir(self.temp_dir):
                try:
                    os.remove(os.path.join(self.temp_dir, file))
                except:
                    pass
        except:
            pass

    def convert_pdfs(self, pdf_paths):
        total_files = len(pdf_paths)
        use_ocr = self.ocr_var.get()
        
        if use_ocr and not self.check_tesseract():
            self.ocr_var.set(False)
            use_ocr = False
        
        try:
            for file_idx, pdf_path in enumerate(pdf_paths, 1):
                output_path = Path(pdf_path).with_suffix('.txt')
                
                with pdfplumber.open(pdf_path) as pdf:
                    total_pages = len(pdf.pages)
                    with open(output_path, 'w', encoding='utf-8') as f:
                        for page_idx, page in enumerate(pdf.pages, 1):
                            text = page.extract_text()
                            
                            # 如果启用OCR且无法提取文本,尝试OCR识别
                            if not text and use_ocr:
                                self.status_label.configure(text=f"正在OCR识别第 {page_idx} 页...")
                                self.logger.info(f"开始OCR识别 {pdf_path} 第 {page_idx} 页")
                                
                                try:
                                    # 生成唯一的临时文件名
                                    temp_image_path = os.path.join(
                                        self.temp_dir, 
                                        f"temp_{uuid.uuid4().hex}.png"
                                    )
                                    
                                    # 转换为图片时使用更高的DPI
                                    img = page.to_image(resolution=300)
                                    img.save(temp_image_path)
                                    
                                    # OCR识别
                                    with Image.open(temp_image_path) as img:
                                        text = self.ocr_image(img) or ''
                                        if text:
                                            self.logger.info(f"OCR识别成功，文本长度: {len(text)}")
                                        else:
                                            self.logger.warning("OCR未能识别出文本")
                                            
                                finally:
                                    # 删除临时图片文件
                                    try:
                                        if os.path.exists(temp_image_path):
                                            os.remove(temp_image_path)
                                    except:
                                        pass
                            
                            if text:
                                f.write(text)
                            
                            # 更新进度
                            progress = (page_idx / total_pages + file_idx - 1) / total_files
                            self.progress.set(progress)
                            self.status_label.configure(
                                text=f"正在转换第 {file_idx}/{total_files} 个文件... ({page_idx}/{total_pages}页)"
                            )
                            self.root.update_idletasks()
            
            self.status_label.configure(text=f"转换完成！已成功转换 {total_files} 个文件")
        except Exception as e:
            self.status_label.configure(text=f"转换失败: {str(e)}")
            messagebox.showerror("错误", f"转换过程中出现错误:\n{str(e)}")
        finally:
            self.progress.set(0)
            self.cleanup_temp_files()

    def run(self):
        try:
            self.root.mainloop()
        finally:
            self.cleanup_temp_files()

if __name__ == "__main__":
    app = PDFConverterGUI()
    app.run()
