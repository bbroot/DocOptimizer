import os  
import sys  
import zipfile  
import tempfile  
import logging  
from PIL import Image  
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,  
                           QPushButton, QLabel, QFileDialog, QProgressBar,  
                           QWidget, QMessageBox, QSpinBox, QGroupBox, QFrame,  
                           QListWidget, QListWidgetItem, QCheckBox, QSizePolicy,  
                           QGraphicsDropShadowEffect)  
from PyQt5.QtCore import (Qt, QThread, pyqtSignal, QMimeData, QSize, QPropertyAnimation,   
                         QEasingCurve, QSequentialAnimationGroup)  
from PyQt5.QtGui import (QDragEnterEvent, QDropEvent, QFont, QIcon, QPixmap,   
                        QColor, QPalette, QLinearGradient, QBrush, QPainter,  
                        QGuiApplication, QPainterPath)  

# 版本信息  
VERSION = "1.1.0"  
AUTHOR = "bbroot"  

# 配置日志  
logging.basicConfig(  
    level=logging.INFO,  
    format='%(asctime)s - %(levelname)s - %(message)s',  
    filename='word_compressor.log'  
)  

class CompressionThread(QThread):  
    progress_updated = pyqtSignal(int, str)  
    finished_signal = pyqtSignal(bool, str)  

    def __init__(self, input_path, output_path, quality):  
        super().__init__()  
        self.input_path = input_path  
        self.output_path = output_path  
        self.quality = quality  
        self.canceled = False  

    def run(self):  
        try:  
            # 验证文件  
            if not os.path.exists(self.input_path):  
                raise FileNotFoundError("输入文件不存在")  
            
            if not self.input_path.lower().endswith(('.docx', '.doc')):  
                raise ValueError("仅支持Word文档 (.docx/.doc)")  

            # 创建临时目录  
            with tempfile.TemporaryDirectory() as temp_dir:  
                # 解压docx文件  
                with zipfile.ZipFile(self.input_path, 'r') as zip_ref:  
                    zip_ref.extractall(temp_dir)  

                # 验证docx结构  
                if not os.path.exists(os.path.join(temp_dir, 'word/media')):  
                    raise ValueError("无效的Word文档结构")  

                # 处理图片  
                media_dir = os.path.join(temp_dir, 'word', 'media')  
                if os.path.exists(media_dir):  
                    self.process_images(media_dir)  

                # 重新打包  
                self.repackage(temp_dir)  

                # 验证输出  
                if not os.path.exists(self.output_path):  
                    raise RuntimeError("创建输出文件失败")  

                # 计算压缩率  
                orig_size = os.path.getsize(self.input_path)  
                comp_size = os.path.getsize(self.output_path)  
                ratio = (orig_size - comp_size) / orig_size * 100  

                self.finished_signal.emit(  
                    True,   
                    f"压缩成功！\n原始大小: {orig_size/1024:.2f}KB "  
                    f"压缩后: {comp_size/1024:.2f}KB "  
                    f"(缩小了 {ratio:.1f}%)"  
                )  

        except Exception as e:  
            self.finished_signal.emit(False, f"压缩失败: {str(e)}")  

    def process_images(self, media_dir):  
        """压缩文档中的图片"""  
        image_files = [f for f in os.listdir(media_dir)   
                      if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))]  
        
        total = len(image_files)  
        for i, img_file in enumerate(image_files):  
            if self.canceled:  
                break  

            img_path = os.path.join(media_dir, img_file)  
            try:  
                with Image.open(img_path) as img:  
                    # 保持原始格式  
                    ext = os.path.splitext(img_file)[1].lower()  
                    if ext in ('.jpg', '.jpeg'):  
                        img.save(img_path, quality=self.quality, optimize=True)  
                    elif ext == '.png':  
                        img.save(img_path, optimize=True)  
                    else:  
                        img.save(img_path)  
                
                progress = int((i + 1) / total * 100)  
                self.progress_updated.emit(progress, f"正在处理图片: {img_file}")  

            except Exception as e:  
                logging.warning(f"图片处理失败: {img_file} - {str(e)}")  
                continue  

    def repackage(self, temp_dir):  
        """重新打包Word文档"""  
        # 确保输出目录存在  
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)  
        
        # 删除现有文件  
        if os.path.exists(self.output_path):  
            os.remove(self.output_path)  
        
        # 创建新的zip文件  
        with zipfile.ZipFile(self.output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:  
            for root, dirs, files in os.walk(temp_dir):  
                for file in files:  
                    file_path = os.path.join(root, file)  
                    arcname = os.path.relpath(file_path, temp_dir)  
                    zipf.write(file_path, arcname)  

class ShadowFrame(QFrame):  
    def __init__(self, parent=None):  
        super().__init__(parent)  
        self.shadow = QGraphicsDropShadowEffect()  
        self.shadow.setBlurRadius(15)  
        self.shadow.setOffset(3)  
        self.shadow.setColor(QColor(0, 0, 0, 100))  
        self.setGraphicsEffect(self.shadow)  

class DropArea(ShadowFrame):  
    files_dropped = pyqtSignal(list)  
    clicked = pyqtSignal()  

    def __init__(self, parent=None):  
        super().__init__(parent)  
        self.setAcceptDrops(True)  
        self.setFrameShape(QFrame.StyledPanel)  
        self.setStyleSheet("""  
            ShadowFrame {  
                border: 2px dashed #aaa;  
                border-radius: 12px;  
                background-color: rgba(255, 255, 255, 100);  
                padding: 20px;  
            }  
        """)  
        self.setMinimumSize(300, 180)  

        layout = QVBoxLayout(self)  
        layout.setAlignment(Qt.AlignCenter)  
        layout.setSpacing(10)  
        
        self.icon = QLabel()  
        self.icon.setPixmap(QIcon.fromTheme("folder-documents").pixmap(64, 64))  
        layout.addWidget(self.icon, 0, Qt.AlignCenter)  
        
        self.label = QLabel("拖放Word文档到此处\n或点击选择文件")  
        self.label.setAlignment(Qt.AlignCenter)  
        self.label.setWordWrap(True)  
        layout.addWidget(self.label)  

    def dragEnterEvent(self, event):  
        if event.mimeData().hasUrls():  
            event.acceptProposedAction()  
            animate = QPropertyAnimation(self, b"radius")  
            animate.setDuration(200)  
            animate.setStartValue(12)  
            animate.setEndValue(24)  
            animate.setEasingCurve(QEasingCurve.OutQuad)  
            animate.start()  
            
            # 使用渐变创建悬停效果  
            gradient = QLinearGradient(0, 0, 0, self.height())  
            gradient.setColorAt(0, QColor(76, 175, 80, 60))  
            gradient.setColorAt(1, QColor(56, 142, 60, 60))  
            
            self.styledEffect = QWidget(self)  
            self.styledEffect.setStyleSheet("background: transparent;")  
            self.styledEffect.resize(self.size())  
            
            def paintEvent(event):  
                painter = QPainter(self.styledEffect)  
                painter.setRenderHint(QPainter.Antialiasing)  
                path = QPainterPath()  
                path.addRoundedRect(self.rect(), 12, 12)  
                painter.setClipPath(path)  
                painter.setBrush(QBrush(gradient))  
                painter.setPen(Qt.NoPen)  
                painter.drawRoundedRect(self.rect(), 12, 12)  
                painter.end()  
                
            self.styledEffect.paintEvent = paintEvent  
            self.styledEffect.update()  
            
            self.label.raise_()  
            self.icon.raise_()  

    def dragLeaveEvent(self, event):  
        self.clear_effects()  

    def dropEvent(self, event):  
        self.clear_effects()  
        urls = event.mimeData().urls()  
        if urls:  
            files = [url.toLocalFile() for url in urls if url.isLocalFile()]  
            if files:  
                self.files_dropped.emit(files)  

    def clear_effects(self):  
        if hasattr(self, 'styledEffect'):  
            self.styledEffect.deleteLater()  
            del self.styledEffect  
        self.setStyleSheet("""  
            ShadowFrame {  
                border: 2px dashed #aaa;  
                border-radius: 12px;  
                background-color: rgba(255, 255, 255, 100);  
                padding: 20px;  
            }  
        """)  

    def mousePressEvent(self, event):  
        if event.button() == Qt.LeftButton:  
            # 点击动画  
            self.click_animation = QPropertyAnimation(self, b"size")  
            self.click_animation.setDuration(100)  
            self.click_animation.setStartValue(self.size())  
            self.click_animation.setEndValue(QSize(self.width()-6, self.height()-6))  
            self.click_animation.setEasingCurve(QEasingCurve.OutQuad)  
            
            self.click_animation2 = QPropertyAnimation(self, b"size")  
            self.click_animation2.setDuration(100)  
            self.click_animation2.setStartValue(QSize(self.width()-6, self.height()-6))  
            self.click_animation2.setEndValue(self.size())  
            self.click_animation2.setEasingCurve(QEasingCurve.InQuad)  
            
            self.anim_group = QSequentialAnimationGroup()  
            self.anim_group.addAnimation(self.click_animation)  
            self.anim_group.addAnimation(self.click_animation2)  
            self.anim_group.start()  
            
            self.clicked.emit()  
            
    def resizeEvent(self, event):  
        super().resizeEvent(event)  
        if hasattr(self, 'styledEffect'):  
            self.styledEffect.resize(self.size())  

class ModernButton(QPushButton):  
    def __init__(self, text, parent=None):  
        super().__init__(text, parent)  
        self.setCursor(Qt.PointingHandCursor)  
        self.setMinimumHeight(36)  
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)  
        
        # 默认样式  
        self.normal_style = """  
            QPushButton {  
                background-color: #4CAF50;  
                color: white;  
                border: none;  
                border-radius: 6px;  
                padding: 8px 16px;  
                font-weight: 500;  
                min-width: 100px;  
            }  
            QPushButton:hover {  
                background-color: #45a049;  
            }  
            QPushButton:pressed {  
                background-color: #3d8b40;  
            }  
            QPushButton:disabled {  
                background-color: #cccccc;  
                color: #666666;  
            }  
        """  
        
        self.setStyleSheet(self.normal_style)  
        
        # 阴影效果  
        self.shadow = QGraphicsDropShadowEffect()  
        self.shadow.setBlurRadius(8)  
        self.shadow.setOffset(2, 2)  
        self.shadow.setColor(QColor(0, 0, 0, 50))  
        self.setGraphicsEffect(self.shadow)  

class SecondaryButton(ModernButton):  
    def __init__(self, text, parent=None):  
        super().__init__(text, parent)  
        self.normal_style = """  
            QPushButton {  
                background-color: #f0f0f0;  
                color: #333333;  
                border: 1px solid #dddddd;  
                border-radius: 6px;  
                padding: 8px 16px;  
                font-weight: 500;  
                min-width: 100px;  
            }  
            QPushButton:hover {  
                background-color: #e0e0e0;  
            }  
            QPushButton:pressed {  
                background-color: #d0d0d0;  
            }  
            QPushButton:disabled {  
                background-color: #f5f5f5;  
                color: #aaaaaa;  
            }  
        """  
        self.setStyleSheet(self.normal_style)  

class DangerButton(ModernButton):  
    def __init__(self, text, parent=None):  
        super().__init__(text, parent)  
        self.normal_style = """  
            QPushButton {  
                background-color: #f44336;  
                color: white;  
                border: none;  
                border-radius: 6px;  
                padding: 8px 16px;  
                font-weight: 500;  
                min-width: 100px;  
            }  
            QPushButton:hover {  
                background-color: #e53935;  
            }  
            QPushButton:pressed {  
                background-color: #d32f2f;  
            }  
            QPushButton:disabled {  
                background-color: #ffcdd2;  
                color: #666666;  
            }  
        """  
        self.setStyleSheet(self.normal_style)  

class MainWindow(QMainWindow):  
    def __init__(self):  
        super().__init__()  
        self.setWindowTitle(f"Word文档压缩工具 v{VERSION}")  
        self.setMinimumSize(1000, 700)  
        
        # 初始化变量  
        self.input_files = []  
        self.output_dir = ""  
        self.compression_thread = None  
        
        # 设置UI  
        self.init_ui()  
        self.update_styles()  
        
        # 居中窗口  
        self.center_window()  

    def center_window(self):  
        frame = self.frameGeometry()  
        center_point = QGuiApplication.primaryScreen().availableGeometry().center()  
        frame.moveCenter(center_point)  
        self.move(frame.topLeft())  

    def init_ui(self):  
        # 创建中央部件  
        central_widget = QWidget()  
        self.setCentralWidget(central_widget)  
        
        # 主布局  
        main_layout = QVBoxLayout(central_widget)  
        main_layout.setContentsMargins(30, 30, 30, 30)  
        main_layout.setSpacing(20)  
        
        # 内容区域  
        content_frame = ShadowFrame()  
        content_frame.setStyleSheet("""  
            ShadowFrame {  
                background-color: white;  
                border-radius: 12px;  
            }  
        """)  
        content_layout = QVBoxLayout(content_frame)  
        content_layout.setContentsMargins(20, 20, 20, 20)  
        content_layout.setSpacing(20)  
        
        # 标题区域  
        title_frame = QFrame()  
        title_frame.setStyleSheet("""  
            QFrame {  
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,  
                    stop:0 #4CAF50, stop:1 #81C784);  
                border-radius: 8px;  
                padding: 15px;  
            }  
        """)  
        title_layout = QVBoxLayout(title_frame)  
        title_layout.setSpacing(8)  
        
        title_label = QLabel(f"Word文档压缩工具 by {AUTHOR}")  
        title_label.setAlignment(Qt.AlignCenter)  
        title_label.setStyleSheet("""  
            QLabel {  
                color: white;  
                font-size: 27px;  
                font-weight: bold;  
            }  
        """)  
        title_layout.addWidget(title_label)  
        
        subtitle_label = QLabel("优化Word文档中的图片，显著减小文件大小")  
        subtitle_label.setAlignment(Qt.AlignCenter)  
        subtitle_label.setStyleSheet("""  
            QLabel {  
                color: rgba(255, 255, 255, 0.9);  
                font-size: 15px;  
            }  
        """)  
        title_layout.addWidget(subtitle_label)  
        
        content_layout.addWidget(title_frame)  

        # 拖放区域  
        self.drop_area = DropArea()  
        self.drop_area.clicked.connect(self.select_input_files)  
        self.drop_area.files_dropped.connect(self.handle_dropped_files)  
        content_layout.addWidget(self.drop_area)  

        # 文件列表  
        self.file_list = QListWidget()  
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)  
        self.file_list.setStyleSheet("""  
            QListWidget {  
                border: 1px solid #e0e0e0;  
                border-radius: 8px;  
                padding: 5px;  
                background-color: white;  
                alternate-background-color: #f9f9f9;  
            }  
            QListWidget::item {  
                padding: 10px;  
                border-bottom: 1px solid #f0f0f0;  
            }  
            QListWidget::item:hover {  
                background-color: #f5f5f5;  
            }  
            QListWidget::item:selected {  
                background-color: #e3f2fd;  
                color: #1976d2;  
            }  
        """)  
        content_layout.addWidget(self.file_list)  

        # 列表操作按钮  
        list_btn_layout = QHBoxLayout()  
        list_btn_layout.setSpacing(15)  
        
        self.clear_btn = SecondaryButton("清空列表")  
        self.clear_btn.clicked.connect(self.clear_file_list)  
        list_btn_layout.addWidget(self.clear_btn)  
        
        self.remove_btn = SecondaryButton("移除选中")  
        self.remove_btn.clicked.connect(self.remove_selected_files)  
        list_btn_layout.addWidget(self.remove_btn)  
        
        list_btn_layout.addStretch()  
        content_layout.addLayout(list_btn_layout)  

        # 设置组  
        settings_group = ShadowFrame()  
        settings_group.setStyleSheet("""  
            ShadowFrame {  
                background-color: white;  
                border-radius: 8px;  
                padding: 15px;  
            }  
        """)  
        
        settings_layout = QVBoxLayout(settings_group)  
        settings_layout.setContentsMargins(10, 10, 10, 10)  
        settings_layout.setSpacing(20)  

        # 质量设置  
        quality_layout = QHBoxLayout()  
        quality_layout.setSpacing(15)  
        
        quality_label = QLabel("图片质量:")  
        quality_label.setStyleSheet("font-weight: bold;")  
        quality_layout.addWidget(quality_label)  
        
        self.quality_spin = QSpinBox()  
        self.quality_spin.setRange(1, 100)  
        self.quality_spin.setValue(75)  
        self.quality_spin.setFixedWidth(100)  
        self.quality_spin.setStyleSheet("""  
            QSpinBox {  
                padding: 5px;  
                border: 1px solid #ddd;  
                border-radius: 4px;  
            }  
            QSpinBox:hover {  
                border-color: #aaa;  
            }  
        """)  
        quality_layout.addWidget(self.quality_spin)  
        
        quality_layout.addStretch()  
        
        self.same_dir_check = QCheckBox("输出到原文件所在目录")  
        self.same_dir_check.setChecked(True)  
        self.same_dir_check.setStyleSheet("""  
            QCheckBox {  
                spacing: 5px;  
            }  
            QCheckBox::indicator {  
                width: 16px;  
                height: 16px;  
            }  
        """)  
        quality_layout.addWidget(self.same_dir_check)  
        
        settings_layout.addLayout(quality_layout)  

        # 输出目录  
        output_layout = QHBoxLayout()  
        output_layout.setSpacing(15)  
        
        output_label = QLabel("输出目录:")  
        output_label.setStyleSheet("font-weight: bold;")  
        output_layout.addWidget(output_label)  
        
        self.output_label = QLabel("未选择")  
        self.output_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)  
        self.output_label.setWordWrap(True)  
        self.output_label.setStyleSheet("""  
            QLabel {  
                color: #555;  
                padding: 5px;  
                border: 1px solid #eee;  
                border-radius: 4px;  
                background-color: #f9f9f9;  
            }  
        """)  
        output_layout.addWidget(self.output_label)  
        
        self.select_dir_btn = SecondaryButton("浏览...")  
        self.select_dir_btn.clicked.connect(self.select_output_directory)  
        output_layout.addWidget(self.select_dir_btn)  
        
        settings_layout.addLayout(output_layout)  
        content_layout.addWidget(settings_group)  

        # 进度条  
        self.progress_bar = QProgressBar()  
        self.progress_bar.setRange(0, 100)  
        self.progress_bar.setTextVisible(True)  
        self.progress_bar.setStyleSheet("""  
            QProgressBar {  
                border: 1px solid #e0e0e0;  
                border-radius: 6px;  
                text-align: center;  
                height: 20px;  
            }  
            QProgressBar::chunk {  
                background-color: #4CAF50;  
                border-radius: 5px;  
            }  
        """)  
        content_layout.addWidget(self.progress_bar)  

        # 状态标签  
        self.status_label = QLabel("就绪")  
        self.status_label.setAlignment(Qt.AlignCenter)  
        self.status_label.setStyleSheet("""  
            QLabel {  
                color: #666;  
                font-style: italic;  
            }  
        """)  
        content_layout.addWidget(self.status_label)  

        # 操作按钮  
        btn_layout = QHBoxLayout()  
        btn_layout.setSpacing(15)  
        
        self.compress_btn = ModernButton("压缩")  
        self.compress_btn.clicked.connect(self.start_compression)  
        btn_layout.addWidget(self.compress_btn)  
        
        self.batch_btn = ModernButton("批量压缩")  
        self.batch_btn.clicked.connect(self.start_batch_compression)  
        btn_layout.addWidget(self.batch_btn)  
        
        self.cancel_btn = DangerButton("取消")  
        self.cancel_btn.setEnabled(False)  
        self.cancel_btn.clicked.connect(self.cancel_compression)  
        btn_layout.addWidget(self.cancel_btn)  
        
        btn_layout.addStretch()  
        content_layout.addLayout(btn_layout)  
        
        main_layout.addWidget(content_frame)  

    def update_styles(self):  
        # 设置应用字体  
        font = QFont()  
        font.setFamily("Microsoft YaHei" if sys.platform == "win32" else   
                      "PingFang SC" if sys.platform == "darwin" else   
                      "Noto Sans CJK SC")  
        font.setStyleHint(QFont.SansSerif)  
        self.setFont(font)  
        
        # 设置窗口背景  
        palette = self.palette()  
        gradient = QLinearGradient(0, 0, 0, 400)  
        gradient.setColorAt(0, QColor(240, 248, 255))  
        gradient.setColorAt(1, QColor(230, 240, 250))  
        palette.setBrush(QPalette.Window, QBrush(gradient))  
        self.setPalette(palette)  

    def select_input_files(self):  
        files, _ = QFileDialog.getOpenFileNames(  
            self, "选择Word文档", "",  
            "Word文档 (*.docx *.doc);;所有文件 (*)"  
        )  
        if files:  
            self.input_files = files  
            self.update_file_list()  
            
            if self.same_dir_check.isChecked() and files:  
                self.output_dir = os.path.dirname(files[0])  
                self.output_label.setText(self.output_dir)  

    def handle_dropped_files(self, file_paths):  
        valid_files = [f for f in file_paths if f.lower().endswith(('.docx', '.doc'))]  
        if not valid_files:  
            QMessageBox.warning(self, "错误", "请拖放有效的Word文档 (.docx/.doc)")  
            return  
        
        self.input_files = valid_files  
        self.update_file_list()  
        
        if self.same_dir_check.isChecked() and valid_files:  
            self.output_dir = os.path.dirname(valid_files[0])  
            self.output_label.setText(self.output_dir)  

    def update_file_list(self):  
        self.file_list.clear()  
        for file in self.input_files:  
            item = QListWidgetItem(file)  
            icon = QIcon.fromTheme("x-office-document")  
            item.setIcon(icon)  
            self.file_list.addItem(item)  
        self.status_label.setText(f"已准备压缩 {len(self.input_files)} 个文件")  

    def clear_file_list(self):  
        self.input_files = []  
        self.file_list.clear()  
        self.status_label.setText("文件列表已清空")  

    def remove_selected_files(self):  
        selected = self.file_list.selectedItems()  
        if not selected:  
            return  
        
        for item in selected:  
            self.input_files.remove(item.text())  
            self.file_list.takeItem(self.file_list.row(item))  
        
        self.status_label.setText(f"已移除 {len(selected)} 个文件")  

    def select_output_directory(self):  
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录")  
        if directory:  
            self.output_dir = directory  
            self.output_label.setText(directory)  
            self.same_dir_check.setChecked(False)  

    def start_compression(self):  
        if not self.input_files:  
            QMessageBox.warning(self, "错误", "请先选择要压缩的文件")  
            return  
        
        if not self.output_dir:  
            QMessageBox.warning(self, "错误", "请选择输出目录")  
            return  
        
        input_path = self.input_files[0]  
        filename = os.path.basename(input_path)  
        output_path = os.path.join(self.output_dir, f"compressed_{filename}")  
        
        quality = self.quality_spin.value()  
        
        self.compression_thread = CompressionThread(input_path, output_path, quality)  
        self.compression_thread.progress_updated.connect(self.update_progress)  
        self.compression_thread.finished_signal.connect(self.compression_finished)  
        
        self.set_controls_enabled(False)  
        self.compression_thread.start()  

    def start_batch_compression(self):  
        if not self.input_files:  
            QMessageBox.warning(self, "错误", "请先选择要压缩的文件")  
            return  
        
        if not self.output_dir:  
            QMessageBox.warning(self, "错误", "请选择输出目录")  
            return  
        
        # 简化处理，实际应为每个文件创建压缩任务  
        self.start_compression()  

    def cancel_compression(self):  
        if self.compression_thread and self.compression_thread.isRunning():  
            self.compression_thread.canceled = True  
            self.compression_thread.terminate()  
            self.status_label.setText("操作已取消")  
        
        self.set_controls_enabled(True)  

    def update_progress(self, value, text):  
        self.progress_bar.setValue(value)  
        self.status_label.setText(text)  

    def compression_finished(self, success, message):  
        if success:  
            # 成功消息  
            msg = QMessageBox(self)  
            msg.setIcon(QMessageBox.Information)  
            msg.setWindowTitle("成功")  
            msg.setText(message)  
            msg.setStandardButtons(QMessageBox.Ok)  
            msg.setStyleSheet("""  
                QMessageBox {  
                    background-color: white;  
                }  
                QMessageBox QLabel {  
                    color: #333;  
                }  
            """)  
            msg.exec_()  
        else:  
            # 错误消息  
            msg = QMessageBox(self)  
            msg.setIcon(QMessageBox.Warning)  
            msg.setWindowTitle("错误")  
            msg.setText(message)  
            msg.setStandardButtons(QMessageBox.Ok)  
            msg.setStyleSheet("""  
                QMessageBox {  
                    background-color: white;  
                }  
                QMessageBox QLabel {  
                    color: #333;  
                }  
            """)  
            msg.exec_()  
        
        self.set_controls_enabled(True)  
        self.progress_bar.setValue(0)  

    def set_controls_enabled(self, enabled):  
        self.drop_area.setEnabled(enabled)  
        self.clear_btn.setEnabled(enabled)  
        self.remove_btn.setEnabled(enabled)  
        self.quality_spin.setEnabled(enabled)  
        self.same_dir_check.setEnabled(enabled)  
        self.select_dir_btn.setEnabled(enabled)  
        self.compress_btn.setEnabled(enabled)  
        self.batch_btn.setEnabled(enabled)  
        self.cancel_btn.setEnabled(not enabled)  

    def closeEvent(self, event):  
        if self.compression_thread and self.compression_thread.isRunning():  
            # 确认对话框  
            msg = QMessageBox(self)  
            msg.setIcon(QMessageBox.Question)  
            msg.setWindowTitle("确认退出")  
            msg.setText("压缩正在进行中，确定要退出吗？")  
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)  
            msg.setDefaultButton(QMessageBox.No)  
            msg.setStyleSheet("""  
                QMessageBox {  
                    background-color: white;  
                }  
                QMessageBox QLabel {  
                    color: #333;  
                }  
            """)  
            
            reply = msg.exec_()  
            
            if reply == QMessageBox.Yes:  
                self.compression_thread.terminate()  
                event.accept()  
            else:  
                event.ignore()  
        else:  
            event.accept()  

if __name__ == "__main__":  
    app = QApplication(sys.argv)  
    
    # 设置应用信息  
    app.setApplicationName("WordCompressor")  
    app.setApplicationDisplayName(f"Word文档压缩工具 v{VERSION}")  
    app.setApplicationVersion(VERSION)  
    
    # 创建并显示主窗口  
    window = MainWindow()  
    window.show()  
    
    sys.exit(app.exec_())  