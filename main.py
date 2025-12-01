"""
Automation Hub - Main Window & UI Shell
PySide6 Desktop Application with Custom Title Bar
"""
import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QStackedWidget, QFrame)
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QFont

from modules import PDFEditorModule, OCRTrainerModule, SchedulerModule

class CustomTitleBar(QWidget):
    """Custom draggable title bar with window controls"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = parent
        self.start_pos = None
        self.setFixedHeight(45)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 0, 10, 0)
        layout.setSpacing(0)
        
        # App Title
        title = QLabel("🗂️ Automation Hub")
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        title.setStyleSheet("color: #00D9FF; letter-spacing: 1px;")
        layout.addWidget(title)
        layout.addStretch()
        
        # Window Control Buttons
        btn_style = """
            QPushButton {
                background: transparent;
                border: none;
                color: white;
                font-size: 16px;
                padding: 8px 12px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background: rgba(255, 255, 255, 0.1);
            }
        """
        
        self.btn_minimize = QPushButton("─")
        self.btn_maximize = QPushButton("□")
        self.btn_close = QPushButton("✕")
        
        for btn in [self.btn_minimize, self.btn_maximize, self.btn_close]:
            btn.setStyleSheet(btn_style)
            btn.setFixedSize(40, 35)
        
        self.btn_close.setStyleSheet(btn_style + """
            QPushButton:hover { background: #E81123; }
        """)
        
        self.btn_minimize.clicked.connect(parent.showMinimized)
        self.btn_maximize.clicked.connect(self.toggle_maximize)
        self.btn_close.clicked.connect(parent.close)
        
        layout.addWidget(self.btn_minimize)
        layout.addWidget(self.btn_maximize)
        layout.addWidget(self.btn_close)
    
    def toggle_maximize(self):
        if self.parent_window.isMaximized():
            self.parent_window.showNormal()
        else:
            self.parent_window.showMaximized()
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_pos = event.globalPosition().toPoint()
    
    def mouseMoveEvent(self, event):
        if self.start_pos:
            delta = event.globalPosition().toPoint() - self.start_pos
            self.parent_window.move(self.parent_window.pos() + delta)
            self.start_pos = event.globalPosition().toPoint()
    
    def mouseReleaseEvent(self, event):
        self.start_pos = None

class MainWindow(QMainWindow):
    """Main application window with sidebar navigation"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automation Hub")
        self.setGeometry(100, 100, 1400, 900)
        self.setWindowFlags(Qt.FramelessWindowHint)
        
        # Central Widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Custom Title Bar
        self.title_bar = CustomTitleBar(self)
        main_layout.addWidget(self.title_bar)
        
        # Content Area (Sidebar + Main Content)
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)
        
        # Sidebar
        self.sidebar = self.create_sidebar()
        content_layout.addWidget(self.sidebar)
        
        # Stacked Widget for Modules
        self.content_stack = QStackedWidget()
        
        # Instantiate Modules
        self.pdf_module = PDFEditorModule()
        self.ocr_module = OCRTrainerModule()
        self.scheduler_module = SchedulerModule()
        
        self.content_stack.addWidget(self.pdf_module)
        self.content_stack.addWidget(self.ocr_module)
        self.content_stack.addWidget(self.scheduler_module)
        
        content_layout.addWidget(self.content_stack)
        main_layout.addLayout(content_layout)
        
        # Apply Styling
        self.apply_styles()
    
    def create_sidebar(self):
        """Create navigation sidebar"""
        sidebar = QFrame()
        sidebar.setFixedWidth(240)
        sidebar.setObjectName("sidebar")
        
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(12, 25, 12, 12)
        layout.setSpacing(10)
        
        # Info Label
        info = QLabel("MODULES")
        info.setStyleSheet("color: #888; font-size: 11px; font-weight: 600; letter-spacing: 1px; padding-left: 10px;")
        layout.addWidget(info)
        
        # Navigation Buttons
        self.nav_buttons = []
        modules = [
            ("📄 PDF Editor", 0),
            ("🔍 OCR Trainer", 1),
            ("⏰ Scheduler", 2),
        ]
        
        for text, index in modules:
            btn = QPushButton(text)
            btn.setObjectName("navButton")
            btn.setFixedHeight(48)
            btn.clicked.connect(lambda checked, i=index: self.switch_module(i))
            layout.addWidget(btn)
            self.nav_buttons.append(btn)
        
        layout.addStretch()
        
        # Footer
        footer = QLabel("v1.0.0")
        footer.setStyleSheet("color: #555; font-size: 10px; padding: 10px;")
        footer.setAlignment(Qt.AlignCenter)
        layout.addWidget(footer)
        
        # Activate first module
        self.nav_buttons[0].setProperty("active", True)
        
        return sidebar
    
    def switch_module(self, index):
        """Switch between modules"""
        self.content_stack.setCurrentIndex(index)
        
        # Update button states
        for i, btn in enumerate(self.nav_buttons):
            btn.setProperty("active", i == index)
            btn.style().unpolish(btn)
            btn.style().polish(btn)
    
    def apply_styles(self):
        """Apply global stylesheet"""
        self.setStyleSheet("""
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #0A0E27, stop:1 #1A1F3A);
            }
            
            #sidebar {
                background: rgba(15, 20, 35, 0.95);
                border-right: 1px solid rgba(255, 255, 255, 0.08);
            }
            
            QPushButton#navButton {
                background: transparent;
                color: #B0B0B0;
                border: none;
                border-radius: 10px;
                text-align: left;
                padding-left: 18px;
                font-size: 14px;
                font-weight: 500;
            }
            
            QPushButton#navButton:hover {
                background: rgba(255, 255, 255, 0.06);
                color: white;
            }
            
            QPushButton#navButton[active="true"] {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #667EEA, stop:1 #764BA2);
                color: white;
                font-weight: 600;
            }
            
            QScrollArea {
                background: rgba(255, 255, 255, 0.03);
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 10px;
            }
            
            QTableWidget {
                background: rgba(255, 255, 255, 0.03);
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 8px;
                color: white;
                gridline-color: rgba(255, 255, 255, 0.05);
            }
            
            QTableWidget::item {
                padding: 5px;
            }
            
            QHeaderView::section {
                background: rgba(255, 255, 255, 0.05);
                color: white;
                padding: 8px;
                border: none;
                font-weight: 600;
            }
            
            QLineEdit, QSpinBox, QComboBox {
                background: rgba(255, 255, 255, 0.05);
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 6px;
                color: white;
                padding: 8px 12px;
                font-size: 13px;
            }
            
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
                border: 1px solid #667EEA;
                background: rgba(255, 255, 255, 0.08);
            }
            
            QListWidget {
                background: rgba(255, 255, 255, 0.03);
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 8px;
                color: white;
                padding: 5px;
            }
            
            QListWidget::item {
                padding: 10px;
                border-radius: 6px;
            }
            
            QListWidget::item:hover {
                background: rgba(255, 255, 255, 0.05);
            }
            
            QListWidget::item:selected {
                background: rgba(102, 126, 234, 0.3);
                color: white;
            }
        """)

def main():
    """Application entry point"""
    # Enable High DPI
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    app.setApplicationName("Automation Hub")
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
