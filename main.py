import sys
import os
import shutil
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                           QVBoxLayout, QWidget, QFileDialog, QProgressBar, QTextEdit,
                           QMessageBox, QHBoxLayout)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QClipboard
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import traceback

def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        base_path = sys._MEIPASS
    else:
        # 如果是开发环境
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def get_tesseract_path():
    # 首先检查打包后的路径
    tesseract_path = get_resource_path(os.path.join('tesseract', 'tesseract.exe'))
    if os.path.exists(tesseract_path):
        return tesseract_path
    
    # 然后检查系统环境变量中的路径
    tesseract_path = shutil.which('tesseract')
    if tesseract_path:
        return tesseract_path
    
    # 最后检查常见安装路径
    common_paths = [
        r'D:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
    
    return None

def get_poppler_path():
    # 首先检查打包后的路径
    poppler_path = get_resource_path('poppler')
    if os.path.exists(poppler_path):
        return poppler_path
    
    # 检查系统环境变量中的路径
    poppler_path = os.environ.get('POPPLER_HOME')
    if poppler_path and os.path.exists(poppler_path):
        return poppler_path
    
    # 检查常见安装路径
    common_paths = [
        r'D:\Program Files\poppler-24.08.0\Library\bin',
        r'C:\Program Files\poppler-24.08.0\Library\bin',
        r'C:\Program Files (x86)\poppler-24.08.0\Library\bin',
        r'D:\Program Files\poppler\bin',
        r'C:\Program Files\poppler\bin',
        r'C:\Program Files (x86)\poppler\bin'
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
    
    return None

def check_dependencies():
    errors = []
    
    # 检查Tesseract
    tesseract_path = get_tesseract_path()
    if not tesseract_path:
        errors.append("找不到Tesseract-OCR")
    else:
        # 检查tessdata目录
        tessdata_dir = os.path.join(os.path.dirname(tesseract_path), 'tessdata')
        if not os.path.exists(tessdata_dir):
            errors.append("找不到tessdata目录")
        else:
            # 检查语言文件
            for lang in ['chi_sim.traineddata', 'eng.traineddata', 'equ.traineddata']:  # 添加公式识别支持
                lang_path = os.path.join(tessdata_dir, lang)
                if not os.path.exists(lang_path):
                    errors.append(f"找不到语言文件: {lang}")
        
        # 检查DLL文件
        tesseract_dir = os.path.dirname(tesseract_path)
        required_dlls = [
            'libtesseract-5.dll',
            'libgcc_s_seh-1.dll',
            'libstdc++-6.dll',
            'libwinpthread-1.dll',
            'zlib1.dll',
            'libpng16-16.dll',
            'libjpeg-8.dll',
            'libtiff-6.dll',
            'libwebp-7.dll'
        ]
        for dll in required_dlls:
            if not os.path.exists(os.path.join(tesseract_dir, dll)):
                errors.append(f"找不到Tesseract组件: {dll}")
    
    # 检查Poppler
    poppler_path = get_poppler_path()
    if not poppler_path:
        errors.append("找不到Poppler")
    else:
        required_dlls = ['pdfinfo.exe', 'pdftoppm.exe']
        for dll in required_dlls:
            if not os.path.exists(os.path.join(poppler_path, dll)):
                errors.append(f"找不到Poppler组件: {dll}")
    
    return errors

class OCRThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    log = pyqtSignal(str)
    
    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path
        
    def run(self):
        try:
            # 设置Tesseract路径
            tesseract_path = get_tesseract_path()
            if tesseract_path:
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                self.log.emit(f"Tesseract路径: {tesseract_path}")
            else:
                raise Exception("找不到Tesseract-OCR")
            
            # 设置Poppler路径
            poppler_path = get_poppler_path()
            if not poppler_path:
                raise Exception("找不到Poppler，请确保程序完整性")
            self.log.emit(f"Poppler路径: {poppler_path}")
            
            # 将PDF转换为图像
            try:
                self.log.emit("开始转换PDF...")
                images = convert_from_path(
                    self.pdf_path,
                    poppler_path=poppler_path,
                    dpi=300  # 提高DPI以获得更好的图像质量
                )
                self.log.emit(f"PDF转换完成，共{len(images)}页")
            except Exception as e:
                self.log.emit(f"PDF转换错误: {str(e)}")
                self.log.emit(traceback.format_exc())
                raise Exception(f"PDF转换失败: {str(e)}")
            
            total_pages = len(images)
            result_text = ""
            
            # 配置Tesseract参数
            custom_config = r'--oem 3 --psm 6 -l chi_sim+eng+equ'  # 使用LSTM引擎，自动页面分割，支持中文、英文和公式
            
            # 对每一页进行OCR
            for i, image in enumerate(images):
                try:
                    self.log.emit(f"正在处理第{i+1}页...")
                    
                    # 预处理图像
                    # 转换为灰度图
                    if image.mode != 'L':
                        image = image.convert('L')
                    
                    # 增强对比度
                    from PIL import ImageEnhance
                    enhancer = ImageEnhance.Contrast(image)
                    image = enhancer.enhance(1.5)  # 增加对比度
                    
                    # 进行OCR识别
                    text = pytesseract.image_to_string(
                        image,
                        config=custom_config,
                        lang='chi_sim+eng+equ'  # 使用中文、英文和公式识别
                    )
                    
                    result_text += f"=== 第 {i+1} 页 ===\n{text}\n\n"
                    progress = int((i + 1) / total_pages * 100)
                    self.progress.emit(progress)
                except Exception as e:
                    self.log.emit(f"OCR识别错误: {str(e)}")
                    self.log.emit(traceback.format_exc())
                    raise Exception(f"OCR识别失败: {str(e)}")
                
            self.finished.emit(result_text)
        except Exception as e:
            error_msg = f"错误: {str(e)}\n\n详细信息:\n{traceback.format_exc()}"
            self.finished.emit(error_msg)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF OCR识别工具")
        self.setGeometry(100, 100, 800, 600)
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        
        # 创建水平布局用于按钮
        button_layout = QHBoxLayout()
        
        # 创建按钮
        self.select_button = QPushButton("选择PDF文件")
        self.select_button.clicked.connect(self.select_pdf)
        
        self.copy_button = QPushButton("复制文本")
        self.copy_button.clicked.connect(self.copy_text)
        self.copy_button.setEnabled(False)
        
        self.export_button = QPushButton("导出文本")
        self.export_button.clicked.connect(self.export_text)
        self.export_button.setEnabled(False)
        
        # 添加按钮到水平布局
        button_layout.addWidget(self.select_button)
        button_layout.addWidget(self.copy_button)
        button_layout.addWidget(self.export_button)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        # 创建日志显示区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(100)  # 限制日志区域高度
        self.log_text.setStyleSheet("background-color: #f0f0f0;")  # 设置背景色
        
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        
        # 添加部件到主布局
        layout.addLayout(button_layout)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.log_text)  # 添加日志显示区域
        layout.addWidget(self.result_text)
        
        main_widget.setLayout(layout)
        
        self.ocr_thread = None
        
        # 检查依赖
        errors = check_dependencies()
        if errors:
            error_msg = "程序初始化失败:\n" + "\n".join(errors)
            QMessageBox.critical(self, "错误", error_msg)
            self.select_button.setEnabled(False)
            self.result_text.setText(error_msg)
        
    def copy_text(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.result_text.toPlainText())
        QMessageBox.information(self, "提示", "文本已复制到剪贴板")
    
    def export_text(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出文本文件",
            "",
            "文本文件 (*.txt)"
        )
        
        if file_path:
            try:
                # 确保文件路径有.txt扩展名
                if not file_path.lower().endswith('.txt'):
                    file_path += '.txt'
                
                # 获取文本内容
                text_content = self.result_text.toPlainText()
                
                # 检查文本内容是否为空
                if not text_content.strip():
                    QMessageBox.warning(self, "警告", "没有可导出的文本内容")
                    return
                
                # 写入文件
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(text_content)
                
                QMessageBox.information(self, "提示", f"文本已成功导出到：\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}\n请检查文件路径是否有写入权限。")
    
    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf)"
        )
        
        if file_path:
            self.start_ocr(file_path)
            
    def update_log(self, message):
        self.log_text.append(message)
        # 滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def start_ocr(self, pdf_path):
        self.select_button.setEnabled(False)
        self.copy_button.setEnabled(False)
        self.export_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_text.clear()
        self.log_text.clear()  # 清空日志
        
        self.ocr_thread = OCRThread(pdf_path)
        self.ocr_thread.progress.connect(self.update_progress)
        self.ocr_thread.finished.connect(self.ocr_finished)
        self.ocr_thread.log.connect(self.update_log)  # 连接日志信号
        self.ocr_thread.start()
        
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        
    def ocr_finished(self, result):
        self.result_text.setText(result)
        self.select_button.setEnabled(True)
        self.copy_button.setEnabled(True)
        self.export_button.setEnabled(True)
        self.progress_bar.setVisible(False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 