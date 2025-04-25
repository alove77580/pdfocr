import sys
import os
import shutil
import json
import hashlib
import threading
import time
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                           QVBoxLayout, QWidget, QFileDialog, QProgressBar, QTextEdit,
                           QMessageBox, QHBoxLayout, QComboBox, QSpinBox, QSlider, QDialog,
                           QDialogButtonBox, QFontComboBox, QListWidget, QSplitter, QMenu)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMimeData, QThreadPool, QRunnable, QMetaObject, Q_ARG, QObject
from PyQt5.QtGui import QClipboard, QDragEnterEvent, QDropEvent
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

def get_tessdata_path():
    # 首先检查打包后的路径
    tessdata_path = get_resource_path('tessdata')
    if os.path.exists(tessdata_path):
        return tessdata_path
    
    # 检查Tesseract安装目录下的tessdata
    tesseract_path = get_tesseract_path()
    if tesseract_path:
        tessdata_path = os.path.join(os.path.dirname(tesseract_path), 'tessdata')
        if os.path.exists(tessdata_path):
            return tessdata_path
    
    # 检查常见安装路径
    common_paths = [
        r'D:\Program Files\Tesseract-OCR\tessdata',
        r'C:\Program Files\Tesseract-OCR\tessdata',
        r'C:\Program Files (x86)\Tesseract-OCR\tessdata'
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
        tessdata_path = get_tessdata_path()
        if not tessdata_path:
            errors.append("找不到tessdata目录")
        else:
            # 设置TESSDATA_PREFIX环境变量
            os.environ['TESSDATA_PREFIX'] = tessdata_path
            
            # 检查语言文件
            for lang in ['chi_sim.traineddata', 'eng.traineddata', 'equ.traineddata']:
                lang_path = os.path.join(tessdata_path, lang)
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

class OCRConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("OCR参数配置")
        self.setModal(True)
        
        layout = QVBoxLayout()
        
        # 添加配置选项
        self.oem_combo = QComboBox()
        self.oem_combo.addItems(["0 - 传统模式", "1 - LSTM模式", "2 - 传统+LSTM", "3 - 默认"])
        self.oem_combo.setCurrentIndex(3)
        layout.addWidget(QLabel("OCR引擎模式:"))
        layout.addWidget(self.oem_combo)
        
        self.psm_combo = QComboBox()
        self.psm_combo.addItems([
            "0 - 仅方向检测",
            "1 - 自动页面分割+方向检测",
            "2 - 自动页面分割，无方向检测",
            "3 - 全自动页面分割，无方向检测",
            "4 - 假设单列变长文本",
            "5 - 假设统一垂直对齐文本",
            "6 - 假设统一块文本",
            "7 - 假设单行文本",
            "8 - 假设单个单词",
            "9 - 假设单个单词圆形",
            "10 - 假设单个字符",
            "11 - 稀疏文本",
            "12 - 稀疏文本+方向检测",
            "13 - 原始行"
        ])
        self.psm_combo.setCurrentIndex(6)
        layout.addWidget(QLabel("页面分割模式:"))
        layout.addWidget(self.psm_combo)
        
        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(100, 1200)
        self.dpi_spin.setValue(600)
        self.dpi_spin.setSingleStep(100)
        layout.addWidget(QLabel("DPI:"))
        layout.addWidget(self.dpi_spin)
        
        self.contrast_slider = QSlider(Qt.Horizontal)
        self.contrast_slider.setRange(0, 200)
        self.contrast_slider.setValue(100)
        layout.addWidget(QLabel("对比度:"))
        layout.addWidget(self.contrast_slider)
        
        self.brightness_slider = QSlider(Qt.Horizontal)
        self.brightness_slider.setRange(0, 200)
        self.brightness_slider.setValue(100)
        layout.addWidget(QLabel("亮度:"))
        layout.addWidget(self.brightness_slider)
        
        # 添加按钮
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)
    
    def get_config(self):
        return {
            'oem': self.oem_combo.currentIndex(),
            'psm': self.psm_combo.currentIndex(),
            'dpi': self.dpi_spin.value(),
            'contrast': self.contrast_slider.value() / 100.0,
            'brightness': self.brightness_slider.value() / 100.0
        }

class OCRWorker(QRunnable):
    def __init__(self, pdf_path, config, progress_callback, log_callback, finished_callback):
        super().__init__()
        self.pdf_path = pdf_path
        self.config = config
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.finished_callback = finished_callback
        self._stop_event = threading.Event()
    
    def run(self):
        try:
            # 检查缓存
            cache_key = self._get_cache_key()
            cached_result = self._get_cached_result(cache_key)
            if cached_result:
                self.log_callback("使用缓存结果")
                self.finished_callback(cached_result)
                return
            
            # 设置Tesseract路径
            tesseract_path = get_tesseract_path()
            if tesseract_path:
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                self.log_callback(f"Tesseract路径: {tesseract_path}")
            else:
                raise Exception("找不到Tesseract-OCR")
            
            # 设置Poppler路径
            poppler_path = get_poppler_path()
            if not poppler_path:
                raise Exception("找不到Poppler，请确保程序完整性")
            self.log_callback(f"Poppler路径: {poppler_path}")
            
            # 将PDF转换为图像
            try:
                self.log_callback("开始转换PDF...")
                images = convert_from_path(
                    self.pdf_path,
                    poppler_path=poppler_path,
                    dpi=self.config['dpi']
                )
                self.log_callback(f"PDF转换完成，共{len(images)}页")
            except Exception as e:
                self.log_callback(f"PDF转换错误: {str(e)}")
                self.log_callback(traceback.format_exc())
                raise Exception(f"PDF转换失败: {str(e)}")
            
            total_pages = len(images)
            result_text = ""
            
            # 配置Tesseract参数
            tessdata_path = get_tessdata_path()
            if not tessdata_path:
                raise Exception("找不到tessdata目录")
            
            # 设置TESSDATA_PREFIX环境变量
            os.environ['TESSDATA_PREFIX'] = tessdata_path
            self.log_callback(f"TESSDATA_PREFIX设置为: {tessdata_path}")
            
            # 检查语言文件是否存在
            for lang in ['chi_sim.traineddata', 'eng.traineddata', 'equ.traineddata']:
                lang_path = os.path.join(tessdata_path, lang)
                if not os.path.exists(lang_path):
                    raise Exception(f"找不到语言文件: {lang}")
                self.log_callback(f"找到语言文件: {lang}")
            
            # 使用配置的参数
            custom_config = f'--oem {self.config["oem"]} --psm {self.config["psm"]}'
            
            # 对每一页进行OCR
            for i, image in enumerate(images):
                if self._stop_event.is_set():
                    self.log_callback("处理已取消")
                    return
                
                try:
                    self.log_callback(f"正在处理第{i+1}页...")
                    
                    # 预处理图像
                    if image.mode != 'L':
                        image = image.convert('L')
                    
                    # 应用配置的对比度和亮度
                    from PIL import ImageEnhance
                    enhancer = ImageEnhance.Contrast(image)
                    image = enhancer.enhance(self.config['contrast'])
                    enhancer = ImageEnhance.Brightness(image)
                    image = enhancer.enhance(self.config['brightness'])
                    
                    # 进行OCR识别
                    text = pytesseract.image_to_string(
                        image,
                        config=custom_config,
                        lang='chi_sim+eng+equ'
                    )
                    
                    result_text += f"=== 第 {i+1} 页 ===\n{text}\n\n"
                    progress = int((i + 1) / total_pages * 100)
                    self.progress_callback(progress)
                except Exception as e:
                    self.log_callback(f"OCR识别错误: {str(e)}")
                    self.log_callback(traceback.format_exc())
                    raise Exception(f"OCR识别失败: {str(e)}")
            
            # 缓存结果
            self._cache_result(cache_key, result_text)
            
            self.finished_callback(result_text)
        except Exception as e:
            error_msg = f"错误: {str(e)}\n\n详细信息:\n{traceback.format_exc()}"
            self.finished_callback(error_msg)
    
    def stop(self):
        self._stop_event.set()
    
    def _get_cache_key(self):
        # 使用文件内容和配置生成缓存键
        with open(self.pdf_path, 'rb') as f:
            file_hash = hashlib.md5(f.read()).hexdigest()
        config_hash = hashlib.md5(json.dumps(self.config, sort_keys=True).encode()).hexdigest()
        return f"{file_hash}_{config_hash}"
    
    def _get_cached_result(self, cache_key):
        cache_dir = os.path.join(os.path.dirname(self.pdf_path), '.ocr_cache')
        os.makedirs(cache_dir, exist_ok=True)
        cache_file = os.path.join(cache_dir, f"{cache_key}.txt")
        
        if os.path.exists(cache_file):
            # 检查缓存是否过期（24小时）
            if os.path.getmtime(cache_file) > (time.time() - 24 * 3600):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    return f.read()
        return None
    
    def _cache_result(self, cache_key, result):
        cache_dir = os.path.join(os.path.dirname(self.pdf_path), '.ocr_cache')
        os.makedirs(cache_dir, exist_ok=True)
        cache_file = os.path.join(cache_dir, f"{cache_key}.txt")
        
        with open(cache_file, 'w', encoding='utf-8') as f:
            f.write(result)

class OCRSignals(QObject):
    log = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF OCR识别工具")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout()
        
        # 添加线程池
        self.thread_pool = QThreadPool()
        self.thread_pool.setMaxThreadCount(4)  # 最大4个线程
        
        # 创建信号对象
        self.signals = OCRSignals()
        
        # 创建左侧面板（历史记录）
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        self.history_list = QListWidget()
        self.history_list.itemClicked.connect(self.load_history_item)
        self.history_list.setMinimumWidth(300)  # 设置最小宽度
        self.history_list.setMaximumWidth(400)  # 设置最大宽度
        self.history_list.setContextMenuPolicy(Qt.CustomContextMenu)  # 启用右键菜单
        self.history_list.customContextMenuRequested.connect(self.show_history_context_menu)  # 连接右键菜单信号
        left_layout.addWidget(QLabel("历史记录"))
        left_layout.addWidget(self.history_list)
        left_panel.setLayout(left_layout)
        
        # 创建右侧面板（主功能）
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        
        # 创建水平布局用于按钮
        button_layout = QHBoxLayout()
        
        # 创建按钮
        self.select_button = QPushButton("选择PDF文件")
        self.select_button.clicked.connect(self.select_pdf)
        
        self.select_multiple_button = QPushButton("批量选择PDF")
        self.select_multiple_button.clicked.connect(self.select_multiple_pdf)
        
        self.copy_button = QPushButton("复制文本")
        self.copy_button.clicked.connect(self.copy_text)
        self.copy_button.setEnabled(False)
        
        self.export_button = QPushButton("导出文本")
        self.export_button.clicked.connect(self.export_text)
        self.export_button.setEnabled(False)
        
        self.export_word_button = QPushButton("导出Word")
        self.export_word_button.clicked.connect(self.export_word)
        self.export_word_button.setEnabled(False)
        
        # 添加配置按钮
        self.config_button = QPushButton("OCR配置")
        self.config_button.clicked.connect(self.show_config_dialog)
        
        # 添加取消按钮
        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self.cancel_ocr)
        self.cancel_button.setEnabled(False)
        
        # 添加按钮到水平布局
        button_layout.addWidget(self.select_button)
        button_layout.addWidget(self.select_multiple_button)
        button_layout.addWidget(self.copy_button)
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.export_word_button)
        button_layout.addWidget(self.config_button)
        button_layout.addWidget(self.cancel_button)
        
        # 创建设置区域
        settings_layout = QHBoxLayout()
        
        # 添加格式选项
        self.format_combo = QComboBox()
        self.format_combo.addItems(["保留原始格式", "纯文本", "Markdown"])
        settings_layout.addWidget(QLabel("输出格式:"))
        settings_layout.addWidget(self.format_combo)
        
        # 添加DPI设置
        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(300, 1200)
        self.dpi_spin.setValue(600)
        self.dpi_spin.setSingleStep(100)
        settings_layout.addWidget(QLabel("DPI:"))
        settings_layout.addWidget(self.dpi_spin)
        
        # 添加主题设置
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["浅色", "深色"])
        self.theme_combo.currentTextChanged.connect(self.change_theme)
        settings_layout.addWidget(QLabel("主题:"))
        settings_layout.addWidget(self.theme_combo)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        # 创建日志显示区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(100)
        self.log_text.setStyleSheet("background-color: #f0f0f0;")
        
        # 创建结果显示区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(False)  # 允许编辑以进行校对
        
        # 添加校对按钮
        self.proofread_button = QPushButton("校对文本")
        self.proofread_button.clicked.connect(self.proofread_text)
        self.proofread_button.setEnabled(False)
        
        # 添加统计信息显示
        self.stats_label = QLabel()
        self.stats_label.setStyleSheet("color: #666666;")
        
        # 添加部件到右侧布局
        right_layout.addLayout(button_layout)
        right_layout.addLayout(settings_layout)
        right_layout.addWidget(self.progress_bar)
        right_layout.addWidget(self.log_text)
        right_layout.addWidget(self.result_text)
        right_layout.addWidget(self.proofread_button)
        right_layout.addWidget(self.stats_label)
        
        right_panel.setLayout(right_layout)
        
        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 4)
        
        main_layout.addWidget(splitter)
        main_widget.setLayout(main_layout)
        
        self.ocr_thread = None
        self.current_theme = "浅色"
        self.history = []
        self.load_history()
        
        # 启用拖放
        self.setAcceptDrops(True)
        
        # 检查依赖
        errors = check_dependencies()
        if errors:
            error_msg = "程序初始化失败:\n" + "\n".join(errors)
            QMessageBox.critical(self, "错误", error_msg)
            self.select_button.setEnabled(False)
            self.select_multiple_button.setEnabled(False)
            self.result_text.setText(error_msg)
        
        self.ocr_config = {
            'oem': 3,
            'psm': 6,
            'dpi': 600,
            'contrast': 1.0,
            'brightness': 1.0
        }
        
        # 连接信号
        self.signals.log.connect(self._update_log)
        self.signals.progress.connect(self._update_progress)
        self.signals.finished.connect(self._ocr_finished)
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        if pdf_files:
            if len(pdf_files) == 1:
                self.start_ocr(pdf_files[0])
            else:
                self.batch_process_pdfs(pdf_files)
    
    def load_history(self):
        try:
            with open('history.json', 'r', encoding='utf-8') as f:
                self.history = json.load(f)
                self.update_history_list()
        except FileNotFoundError:
            self.history = []
    
    def save_history(self):
        with open('history.json', 'w', encoding='utf-8') as f:
            json.dump(self.history, f, ensure_ascii=False, indent=2)
    
    def update_history_list(self):
        self.history_list.clear()
        for item in reversed(self.history):
            # 获取文件大小
            file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), item['filename'])
            file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
            file_size_str = f"{file_size/1024:.1f}KB" if file_size < 1024*1024 else f"{file_size/1024/1024:.1f}MB"
            
            # 获取文本统计信息
            text = item['result']
            lines = len(text.split('\n'))
            words = sum(len(line.split()) for line in text.split('\n'))
            chars = len(text)
            
            # 格式化显示
            display_text = (
                f"📄 {item['filename']}\n"
                f"⏰ {item['time']}\n"
                f"📊 {file_size_str} | {lines}行 | {words}字 | {chars}字符"
            )
            self.history_list.addItem(display_text)
    
    def add_to_history(self, filename, result):
        history_item = {
            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'filename': filename,  # 保存完整路径
            'result': result
        }
        self.history.append(history_item)
        self.save_history()
        self.update_history_list()
    
    def load_history_item(self, item):
        index = self.history_list.row(item)
        history_item = self.history[-(index + 1)]
        self.result_text.setText(history_item['result'])
        self.copy_button.setEnabled(True)
        self.export_button.setEnabled(True)
        self.export_word_button.setEnabled(True)
        self.proofread_button.setEnabled(True)
    
    def proofread_text(self):
        # 这里可以添加更复杂的校对逻辑
        text = self.result_text.toPlainText()
        # 简单的校对示例：去除多余的空行
        text = '\n'.join(line for line in text.split('\n') if line.strip())
        self.result_text.setText(text)
        QMessageBox.information(self, "校对完成", "文本校对已完成")
    
    def export_word(self):
        try:
            from docx import Document
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "导出Word文档",
                "",
                "Word文档 (*.docx)"
            )
            
            if file_path:
                if not file_path.lower().endswith('.docx'):
                    file_path += '.docx'
                
                doc = Document()
                doc.add_paragraph(self.result_text.toPlainText())
                doc.save(file_path)
                QMessageBox.information(self, "提示", f"Word文档已成功导出到：\n{file_path}")
        except ImportError:
            QMessageBox.warning(self, "警告", "请先安装python-docx库：\npip install python-docx")
    
    def show_config_dialog(self):
        dialog = OCRConfigDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.ocr_config = dialog.get_config()
            self.log_text.append("OCR配置已更新")
    
    def update_stats(self, text):
        # 更新统计信息
        lines = text.split('\n')
        words = sum(len(line.split()) for line in lines)
        chars = sum(len(line) for line in lines)
        self.stats_label.setText(
            f"行数: {len(lines)} | 字数: {words} | 字符数: {chars}"
        )
    
    def _update_log(self, message):
        self.log_text.append(message)
        # 滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def _update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def _ocr_finished(self, result):
        self.result_text.setText(result)
        self.select_button.setEnabled(True)
        self.select_multiple_button.setEnabled(True)
        self.copy_button.setEnabled(True)
        self.export_button.setEnabled(True)
        self.export_word_button.setEnabled(True)
        self.proofread_button.setEnabled(True)
        self.config_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.progress_bar.setVisible(False)
        
        # 更新统计信息
        self.update_stats(result)
        
        # 添加到历史记录
        if hasattr(self, 'current_pdf_path'):
            self.add_to_history(self.current_pdf_path, result)
    
    def change_theme(self, theme):
        if theme == "深色" and self.current_theme == "浅色":
            self.setStyleSheet("""
                QMainWindow, QWidget {
                    background-color: #2b2b2b;
                    color: #ffffff;
                }
                QPushButton {
                    background-color: #3c3f41;
                    color: #ffffff;
                    border: 1px solid #555555;
                }
                QPushButton:hover {
                    background-color: #4c4f51;
                }
                QTextEdit {
                    background-color: #323232;
                    color: #ffffff;
                }
                QComboBox, QSpinBox {
                    background-color: #3c3f41;
                    color: #ffffff;
                }
            """)
            self.current_theme = "深色"
        elif theme == "浅色" and self.current_theme == "深色":
            self.setStyleSheet("")
            self.current_theme = "浅色"
    
    def select_multiple_pdf(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择多个PDF文件",
            "",
            "PDF文件 (*.pdf)"
        )
        
        if file_paths:
            self.batch_process_pdfs(file_paths)
    
    def batch_process_pdfs(self, file_paths):
        total_files = len(file_paths)
        self.progress_bar.setMaximum(total_files)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        
        for i, file_path in enumerate(file_paths):
            self.log_text.append(f"正在处理文件 {i+1}/{total_files}: {os.path.basename(file_path)}")
            self.start_ocr(file_path)
            self.progress_bar.setValue(i + 1)
        
        self.progress_bar.setVisible(False)
        QMessageBox.information(self, "完成", f"已处理 {total_files} 个文件")
    
    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf)"
        )
        
        if file_path:
            self.start_ocr(file_path)
    
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
    
    def start_ocr(self, pdf_path):
        self.select_button.setEnabled(False)
        self.select_multiple_button.setEnabled(False)
        self.copy_button.setEnabled(False)
        self.export_button.setEnabled(False)
        self.export_word_button.setEnabled(False)
        self.proofread_button.setEnabled(False)
        self.config_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_text.clear()
        self.log_text.clear()
        
        self.current_pdf_path = pdf_path
        self.current_worker = OCRWorker(
            pdf_path,
            self.ocr_config,
            self.signals.progress.emit,
            self.signals.log.emit,
            self.signals.finished.emit
        )
        self.thread_pool.start(self.current_worker)
    
    def cancel_ocr(self):
        if hasattr(self, 'current_worker'):
            self.current_worker.stop()
            self.cancel_button.setEnabled(False)
            self.select_button.setEnabled(True)
            self.select_multiple_button.setEnabled(True)
            self.config_button.setEnabled(True)
            self.progress_bar.setVisible(False)
            self.log_text.append("处理已取消")
    
    def show_history_context_menu(self, position):
        # 创建右键菜单
        menu = QMenu()
        delete_action = menu.addAction("删除")
        
        # 获取右键点击的项目
        item = self.history_list.itemAt(position)
        if item:
            # 执行删除操作
            action = menu.exec_(self.history_list.mapToGlobal(position))
            if action == delete_action:
                self.delete_history_item(item)
    
    def delete_history_item(self, item):
        # 获取要删除的项目的索引
        index = self.history_list.row(item)
        # 从历史记录中删除
        del self.history[-(index + 1)]
        # 保存更新后的历史记录
        self.save_history()
        # 更新历史记录列表显示
        self.update_history_list()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 