import sys
import os
from pathlib import Path
import logging
import uuid
import shutil
from typing import List, Dict
from PIL import Image, ImageEnhance
import pdfplumber
import pytesseract

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QCheckBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QProgressBar, QFileDialog, QMessageBox, QStatusBar,
    QStyle, QStyleFactory
)
from PyQt6.QtCore import Qt, pyqtSignal, QSize, QThread
from PyQt6.QtGui import QIcon, QFont
import subprocess

class ConvertWorker(QThread):
    """PDF转换工作线程"""
    progress_updated = pyqtSignal(str, int, int)  # 文件名, 当前页, 总页数
    file_completed = pyqtSignal(dict, bool)  # file_data, success
    conversion_completed = pyqtSignal()
    status_updated = pyqtSignal(dict, str)  # file_data, status_text

    def __init__(self, files, use_ocr, temp_dir, logger):
        super().__init__()
        self.files = files
        self.use_ocr = use_ocr
        self.temp_dir = temp_dir
        self.running = True
        self.logger = logger

    def preprocess_image(self, image):
        """预处理图片以提高OCR识别率"""
        image = image.convert('L')
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)
        enhancer = ImageEnhance.Brightness(image)
        image = enhancer.enhance(1.5)
        return image

    def ocr_image(self, image):
        """对图片进行OCR识别"""
        try:
            processed_image = self.preprocess_image(image)
            text = pytesseract.image_to_string(processed_image, lang='chi_sim+eng')
            if not text.strip():
                self.logger.warning("OCR结果为空")
                return None
            return text
        except Exception as e:
            self.logger.error(f"OCR识别错误: {str(e)}")
            return None

    def convert_single_file(self, file_data):
        """转换单个文件"""
        pdf_path = file_data["path"]
        output_path = Path(pdf_path).with_suffix('.txt')
        file_data["output_path"] = str(output_path)  # 保存输出路径
        
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            with open(output_path, 'w', encoding='utf-8') as f:
                for page_idx, page in enumerate(pdf.pages, 1):
                    if not self.running:
                        return False
                        
                    text = page.extract_text()
                    
                    if not text and self.use_ocr:
                        # 更新状态
                        self.progress_updated.emit(file_data["name"], page_idx, total_pages)
                        self.status_updated.emit(file_data, f"OCR识别中 {page_idx}/{total_pages}")
                        self.logger.info(f"开始OCR识别 {pdf_path} 第 {page_idx} 页")
                        
                        try:
                            temp_image_path = os.path.join(
                                self.temp_dir, 
                                f"temp_{uuid.uuid4().hex}.png"
                            )
                            
                            img = page.to_image(resolution=300)
                            img.save(temp_image_path)
                            
                            with Image.open(temp_image_path) as img:
                                text = self.ocr_image(img) or ''
                        finally:
                            try:
                                if os.path.exists(temp_image_path):
                                    os.remove(temp_image_path)
                            except:
                                pass
                    
                    if text:
                        f.write(text)
                    
                    # 更新状态
                    self.progress_updated.emit(file_data["name"], page_idx, total_pages)
                    self.status_updated.emit(file_data, f"转换中 {page_idx}/{total_pages}")
        
        return True

    def run(self):
        try:
            for idx, file_data in enumerate(self.files):
                if not self.running:
                    break
                
                try:
                    success = self.convert_single_file(file_data)
                    if success:
                        self.file_completed.emit(file_data, True)
                    else:
                        self.file_completed.emit(file_data, False)
                except Exception as e:
                    self.logger.error(f"转换失败: {str(e)}")
                    self.file_completed.emit(file_data, False)
        finally:
            self.conversion_completed.emit()

    def stop(self):
        self.running = False

class PDFConverterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 转换工具")
        self.setMinimumSize(900, 600)
        
        # 设置应用样式
        self.setStyle(QStyleFactory.create("Fusion"))
        self.apply_material_style()
        
        # 文件记录
        self.files_data: List[Dict] = []
        
        # 创建临时目录
        self.temp_dir = os.path.join(os.path.expanduser('~'), '.pdf_converter_temp')
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
            
        # 设置日志
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger('PDFConverter')
        
        # 检查Tesseract
        self.check_tesseract()
        
        # 创建界面
        self.init_ui()
        
        self.worker = None
        
    def apply_material_style(self):
        """应用Material Design样式"""
        self.setStyleSheet("""
            QMainWindow {
                background: #f5f5f5;
            }
            QWidget {
                font-family: 'Segoe UI', 'Microsoft YaHei';
            }
            QPushButton {
                background: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #1976D2;
            }
            QPushButton:pressed {
                background: #1565C0;
            }
            QPushButton:disabled {
                background: #BDBDBD;
            }
            QTableWidget {
                background: white;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                gridline-color: #F5F5F5;
            }
            QTableWidget::item {
                padding: 8px;
            }
            QHeaderView::section {
                background: #FAFAFA;
                color: #424242;
                padding: 8px;
                border: none;
                border-bottom: 1px solid #E0E0E0;
                font-weight: bold;
            }
            QStatusBar {
                background: white;
                color: #424242;
            }
            QProgressBar {
                border: none;
                background: #E0E0E0;
                border-radius: 2px;
            }
            QProgressBar::chunk {
                background: #2196F3;
                border-radius: 2px;
            }
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        
    def init_ui(self):
        """初始化界面"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 工具栏 - 使用卡片式设计
        toolbar = QWidget()
        toolbar.setObjectName("toolbarCard")
        toolbar.setStyleSheet("""
            #toolbarCard {
                background: white;
                border-radius: 8px;
                padding: 16px;
            }
        """)
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(16, 16, 16, 16)
        
        # 标题
        title_label = QLabel("PDF 批量转换工具")
        title_label.setFont(QFont('', 20, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #212121;")
        toolbar_layout.addWidget(title_label)
        
        # 按钮组
        button_widget = QWidget()
        button_layout = QHBoxLayout(button_widget)
        button_layout.setSpacing(12)
        
        self.select_btn = QPushButton("添加文件")
        self.select_btn.setMinimumWidth(120)
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setMinimumWidth(120)
        
        self.ocr_checkbox = QCheckBox("启用OCR")
        self.ocr_checkbox.setStyleSheet("""
            QCheckBox {
                color: #424242;
                font-size: 14px;
            }
        """)
        
        button_layout.addWidget(self.select_btn)
        button_layout.addWidget(self.convert_btn)
        button_layout.addWidget(self.ocr_checkbox)
        
        # 添加取消按钮（初始禁用）
        self.cancel_btn = QPushButton("取消转换")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background: #FF5722;
            }
            QPushButton:hover {
                background: #F4511E;
            }
        """)
        self.cancel_btn.clicked.connect(self.cancel_conversion)
        button_layout.addWidget(self.cancel_btn)
        
        toolbar_layout.addWidget(button_widget)
        toolbar_layout.addStretch()
        layout.addWidget(toolbar)
        
        # 文件表格
        self.table = QTableWidget()
        self.table.setColumnCount(6)  # 增加一列用于打开文件位置
        self.table.setHorizontalHeaderLabels(["文件名", "大小", "页数", "状态", "打开位置", "操作"])
        
        # 设置表格样式
        self.table.setShowGrid(False)  # 隐藏网格线
        self.table.setAlternatingRowColors(True)  # 交替行颜色
        self.table.setStyleSheet(self.table.styleSheet() + """
            QTableWidget {
                alternate-background-color: #FAFAFA;
            }
            QTableWidget::item:selected {
                background: #E3F2FD;
                color: #212121;
            }
        """)
        
        # 设置列宽
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(50)  # 增加行高
        column_widths = [0, 100, 80, 100, 120, 120]  # 第一列自适应(Stretch)
        for i, width in enumerate(column_widths[1:], 1):
            self.table.setColumnWidth(i, width)
            
        layout.addWidget(self.table)
        
        # 状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.setMaximumHeight(32)
        
        # 进度条
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(200)
        self.progress.setMaximumHeight(10)
        self.progress.setTextVisible(False)
        self.statusBar.addPermanentWidget(self.progress)
        
        self.statusBar.showMessage("就绪")

        # 绑定事件
        self.select_btn.clicked.connect(self.select_files)
        self.convert_btn.clicked.connect(self.start_conversion)
        
    def add_file_to_table(self, file_data: Dict):
        """添加文件到表格"""
        row = self.table.rowCount()
        self.table.insertRow(row)
        
        # 文件名
        name_item = QTableWidgetItem(file_data["name"])
        name_item.setToolTip(file_data["path"])
        self.table.setItem(row, 0, name_item)
        
        # 大小
        self.table.setItem(row, 1, QTableWidgetItem(file_data["size"]))
        # 页数
        self.table.setItem(row, 2, QTableWidgetItem(file_data["pages"]))
        # 状态
        status_item = QTableWidgetItem(file_data["status"])
        self.table.setItem(row, 3, status_item)
        file_data["status_item"] = status_item
        
        # 打开位置按钮(初始禁用)
        open_location_btn = QPushButton("打开位置")
        open_location_btn.setEnabled(False)
        open_location_btn.setMinimumHeight(40)
        open_location_btn.setStyleSheet("""
            QPushButton {
                font-size: 12px;
                padding: 5px 10px;
            }
        """)
        # 修改这里，使用output_path而不是path
        open_location_btn.clicked.connect(lambda: self.open_file_location(file_data))
        self.table.setCellWidget(row, 4, open_location_btn)
        file_data["location_btn"] = open_location_btn
        
        # 删除按钮
        delete_btn = QPushButton("删除文件")
        delete_btn.setMinimumHeight(40)
        delete_btn.setStyleSheet("""
            QPushButton {
                background: #F44336;
                font-size: 12px;
                padding: 5px 10px;
            }
            QPushButton:hover {
                background: #D32F2F;
            }
        """)
        delete_btn.clicked.connect(lambda: self.remove_file(row, file_data))
        self.table.setCellWidget(row, 5, delete_btn)
        
        file_data["row"] = row
        
    def remove_file(self, row: int, file_data: Dict):
        """从表格中移除文件"""
        self.table.removeRow(row)
        self.files_data.remove(file_data)
        # 更新其他文件的行索引
        for f in self.files_data:
            if f["row"] > row:
                f["row"] -= 1
                
    def select_files(self):
        """选择PDF文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf)"
        )
        
        for path in files:
            if any(f["path"] == path for f in self.files_data):
                continue
                
            try:
                with pdfplumber.open(path) as pdf:
                    file_data = {
                        "path": path,
                        "name": os.path.basename(path),
                        "size": f"{os.path.getsize(path) / 1024 / 1024:.1f} MB",
                        "pages": str(len(pdf.pages)),
                        "status": "待转换"
                    }
                    self.files_data.append(file_data)
                    self.add_file_to_table(file_data)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法读取文件 {path}:\n{str(e)}")
                
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
                    QMessageBox.warning(self, "警告", "未检测到Tesseract-OCR，OCR功能可能无法使用。\n请安装Tesseract-OCR后重试。")
                    self.logger.warning("Tesseract not found")
                    return False
        return True

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
            QMessageBox.warning(self, "OCR警告", f"OCR识别出现错误：{str(e)}\n如果是tesseract错误，请确认已正确安装Tesseract-OCR")
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

    def start_conversion(self):
        """开始转换所有文件"""
        if not self.files_data:
            QMessageBox.information(self, "提示", "请先添加需要转换的文件")
            return
            
        unconverted_files = [f for f in self.files_data if f["status"] == "待转换"]
        if not unconverted_files:
            QMessageBox.information(self, "提示", "没有需要转换的文件")
            return
        
        # 禁用相关按钮
        self.select_btn.setEnabled(False)
        self.convert_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        
        # 创建并启动工作线程
        self.worker = ConvertWorker(
            unconverted_files,
            self.ocr_checkbox.isChecked(),
            self.temp_dir,
            self.logger
        )
        self.worker.progress_updated.connect(self.update_conversion_progress)
        self.worker.status_updated.connect(self.update_file_status)
        self.worker.file_completed.connect(self.handle_file_completed)
        self.worker.conversion_completed.connect(self.handle_conversion_completed)
        self.worker.start()

    def cancel_conversion(self):
        """取消转换"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
            self.statusBar.showMessage("转换已取消")
            self.handle_conversion_completed()

    def update_conversion_progress(self, filename, current_page, total_pages):
        """更新转换进度"""
        self.statusBar.showMessage(f"正在转换: {filename} ({current_page}/{total_pages})")

    def handle_file_completed(self, file_data, success):
        """处理单个文件转换完成"""
        file_data["status_item"].setText("已完成" if success else "失败")
        if success:
            file_data["location_btn"].setEnabled(True)

    def handle_conversion_completed(self):
        """处理所有文件转换完成"""
        # 恢复按钮状态
        self.select_btn.setEnabled(True)
        self.convert_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.progress.setValue(0)
        self.statusBar.showMessage("转换完成")
        self.worker = None

    def update_file_status(self, file_data, status_text):
        """更新文件状态"""
        file_data["status_item"].setText(status_text)

    def open_file_location(self, file_data):
        """打开文件所在位置"""
        output_path = file_data.get("output_path")
        if not output_path or not os.path.exists(output_path):
            QMessageBox.warning(self, "警告", "转换后的文件不存在")
            return
            
        if sys.platform == "win32":
            subprocess.run(['explorer', '/select,', output_path], check=False)
        elif sys.platform == "darwin":  # macOS
            subprocess.run(['open', '-R', output_path], check=False)
        else:  # Linux
            subprocess.run(['xdg-open', os.path.dirname(output_path)], check=False)

if __name__ == "__main__":
    # 设置环境变量
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_SCALE_FACTOR"] = "1"

    app = QApplication(sys.argv)
    
    # 在 PyQt6 中，高DPI缩放已经默认启用
    # 我们只需要设置一些额外的选项
    if hasattr(Qt.ApplicationAttribute, 'UseHighDpiPixmaps'):
        app.setAttribute(Qt.ApplicationAttribute.UseHighDpiPixmaps)
        
    # 设置字体DPI
    font = app.font()
    font.setPointSize(10)  # 基础字体大小
    app.setFont(font)
    
    # 设置样式
    app.setStyle('Fusion')
    
    window = PDFConverterGUI()
    window.show()
    sys.exit(app.exec())
