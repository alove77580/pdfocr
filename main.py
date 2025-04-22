import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                           QVBoxLayout, QWidget, QFileDialog, QProgressBar, QTextEdit,
                           QMessageBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
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
    tesseract_path = get_resource_path(os.path.join('tesseract', 'tesseract.exe'))
    if os.path.exists(tesseract_path):
        return tesseract_path
    return None

def get_poppler_path():
    poppler_path = get_resource_path('poppler')
    if os.path.exists(poppler_path):
        return poppler_path
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
            for lang in ['chi_sim.traineddata', 'eng.traineddata']:
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
    
    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path
        
    def run(self):
        try:
            # 设置Tesseract路径
            tesseract_path = get_tesseract_path()
            if tesseract_path:
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                print(f"Tesseract路径: {tesseract_path}")
            else:
                raise Exception("找不到Tesseract-OCR")
            
            # 设置Poppler路径
            poppler_path = get_poppler_path()
            if not poppler_path:
                raise Exception("找不到Poppler，请确保程序完整性")
            print(f"Poppler路径: {poppler_path}")
            
            # 将PDF转换为图像
            try:
                print("开始转换PDF...")
                images = convert_from_path(
                    self.pdf_path,
                    poppler_path=poppler_path
                )
                print(f"PDF转换完成，共{len(images)}页")
            except Exception as e:
                print(f"PDF转换错误: {str(e)}")
                print(traceback.format_exc())
                raise Exception(f"PDF转换失败: {str(e)}")
            
            total_pages = len(images)
            result_text = ""
            
            # 对每一页进行OCR
            for i, image in enumerate(images):
                try:
                    print(f"正在处理第{i+1}页...")
                    text = pytesseract.image_to_string(image, lang='chi_sim+eng')
                    result_text += f"=== 第 {i+1} 页 ===\n{text}\n\n"
                    progress = int((i + 1) / total_pages * 100)
                    self.progress.emit(progress)
                except Exception as e:
                    print(f"OCR识别错误: {str(e)}")
                    print(traceback.format_exc())
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
        
        # 创建按钮和标签
        self.select_button = QPushButton("选择PDF文件")
        self.select_button.clicked.connect(self.select_pdf)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        
        # 添加部件到布局
        layout.addWidget(self.select_button)
        layout.addWidget(self.progress_bar)
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
        
    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf)"
        )
        
        if file_path:
            self.start_ocr(file_path)
            
    def start_ocr(self, pdf_path):
        self.select_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_text.clear()
        
        self.ocr_thread = OCRThread(pdf_path)
        self.ocr_thread.progress.connect(self.update_progress)
        self.ocr_thread.finished.connect(self.ocr_finished)
        self.ocr_thread.start()
        
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        
    def ocr_finished(self, result):
        self.result_text.setText(result)
        self.select_button.setEnabled(True)
        self.progress_bar.setVisible(False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 