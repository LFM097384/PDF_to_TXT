import pdfplumber
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
import sv_ttk
import darkdetect
from pathlib import Path
import os
from PIL import Image
import sys

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

        # 底部信息
        self.info_label = ctk.CTkLabel(
            self.main_frame,
            text="© 2024 PDF转换工具",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        self.info_label.pack(side="bottom", pady=10)

    def select_files(self):
        pdf_paths = filedialog.askopenfilenames(
            filetypes=[("PDF 文件", "*.pdf")]
        )
        if pdf_paths:
            self.file_label.configure(text=f"已选择: {len(pdf_paths)} 个文件")
            self.status_label.configure(text="正在转换...")
            self.convert_pdfs(pdf_paths)

    def convert_pdfs(self, pdf_paths):
        total_files = len(pdf_paths)
        
        try:
            for file_idx, pdf_path in enumerate(pdf_paths, 1):
                output_path = Path(pdf_path).with_suffix('.txt')
                
                with pdfplumber.open(pdf_path) as pdf:
                    total_pages = len(pdf.pages)
                    with open(output_path, 'w', encoding='utf-8') as f:
                        for page_idx, page in enumerate(pdf.pages, 1):
                            text = page.extract_text()
                            if text:
                                f.write(text)
                            # 计算总体进度: 当前文件进度 + 已完成文件进度
                            progress = (page_idx / total_pages + file_idx - 1) / total_files
                            self.progress.set(progress)
                            self.status_label.configure(
                                text=f"正在转换第 {file_idx}/{total_files} 个文件... ({page_idx}/{total_pages}页)"
                            )
                            self.root.update_idletasks()
            
            self.status_label.configure(text=f"转换完成！已成功转换 {total_files} 个文件")
        except Exception as e:
            self.status_label.configure(text=f"转换失败: {str(e)}")
        finally:
            self.progress.set(0)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = PDFConverterGUI()
    app.run()
