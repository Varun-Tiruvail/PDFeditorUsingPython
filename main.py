"""
Automation Hub - Main Window & UI Shell
PySide6 Desktop Application with Custom Title Bar
"""
import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QStackedWidget, QFrame)
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QFont

from modules import PDFEditorModule, OCRTrainerModule, SchedulerModule, MailDrafterModule

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
        title = QLabel("🗂️ Custom Reporting Automation Hub")
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
        
        # Theme Toggle
        self.btn_theme = QPushButton("🌓")
        self.btn_theme.setStyleSheet(btn_style)
        self.btn_theme.setFixedSize(40, 35)
        self.btn_theme.clicked.connect(parent.toggle_theme)
        
        layout.addWidget(self.btn_theme)
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
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        self.current_theme = "dark"
        
        # Central Widget
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10) # Margin for window shadow/rounded corners
        main_layout.setSpacing(0)
        
        # Main Container (for rounded corners)
        self.container = QFrame()
        self.container.setObjectName("mainContainer")
        container_layout = QVBoxLayout(self.container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)
        main_layout.addWidget(self.container)
        
        # Custom Title Bar
        self.title_bar = CustomTitleBar(self)
        container_layout.addWidget(self.title_bar)
        
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
        self.mail_module = MailDrafterModule(self.pdf_module)
        
        self.content_stack.addWidget(self.pdf_module)
        self.content_stack.addWidget(self.ocr_module)
        self.content_stack.addWidget(self.scheduler_module)
        self.content_stack.addWidget(self.mail_module)
        
        content_layout.addWidget(self.content_stack)
        container_layout.addLayout(content_layout)
        
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
        info.setObjectName("modulesLabel")
        layout.addWidget(info)
        
        # Navigation Buttons
        self.nav_buttons = []
        modules = [
            ("📄 PDF Editor", 0),
            ("🔍 OCR Trainer", 1),
            ("⏰ Scheduler", 2),
            ("📧 Mail Drafter", 3), # Added Mail Drafter module
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
        footer.setObjectName("footerLabel")
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
            
    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_styles()
    
    def apply_styles(self):
        """Apply global stylesheet based on theme"""
        
        if self.current_theme == "dark":
            # Dark Glassmorphism Theme
            bg_gradient = "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #0f172a, stop:1 #1e293b)"
            sidebar_bg = "rgba(30, 41, 59, 0.7)"
            text_color = "#e2e8f0"
            secondary_text = "#94a3b8"
            accent_color = "#3b82f6"
            glass_bg = "rgba(255, 255, 255, 0.05)"
            border_color = "rgba(255, 255, 255, 0.1)"
            btn_hover = "rgba(255, 255, 255, 0.1)"
            input_bg = "rgba(15, 23, 42, 0.6)"
        else:
            # Light Glassmorphism Theme
            bg_gradient = "qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f8fafc, stop:1 #e2e8f0)"
            sidebar_bg = "rgba(255, 255, 255, 0.7)"
            text_color = "#1e293b"
            secondary_text = "#64748b"
            accent_color = "#3b82f6"
            glass_bg = "rgba(255, 255, 255, 0.5)"
            border_color = "rgba(0, 0, 0, 0.05)"
            btn_hover = "rgba(0, 0, 0, 0.05)"
            input_bg = "rgba(255, 255, 255, 0.8)"

        self.setStyleSheet(f"""
            QMainWindow {{
                background: transparent;
            }}
            
            #mainContainer {{
                background: {bg_gradient};
                border-radius: 16px;
                border: 1px solid {border_color};
            }}
            
            #sidebar {{
                background: {sidebar_bg};
                border-right: 1px solid {border_color};
                border-top-left-radius: 0px;
                border-bottom-left-radius: 16px;
            }}
            
            QLabel {{
                color: {text_color};
            }}
            
            #modulesLabel {{
                color: {secondary_text};
                font-size: 11px; 
                font-weight: 700; 
                letter-spacing: 1px; 
                padding-left: 10px;
            }}
            
            #footerLabel {{
                color: {secondary_text};
                font-size: 10px; 
                padding: 10px;
            }}
            
            QPushButton#navButton {{
                background: transparent;
                color: {secondary_text};
                border: none;
                border-radius: 10px;
                text-align: left;
                padding-left: 18px;
                font-size: 14px;
                font-weight: 500;
            }}
            
            QPushButton#navButton:hover {{
                background: {btn_hover};
                color: {text_color};
            }}
            
            QPushButton#navButton[active="true"] {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6366f1, stop:1 #8b5cf6);
                color: white;
                font-weight: 600;
            }}
            
            /* Glassmorphism Panels */
            QScrollArea, QTableWidget, QListWidget, QTabWidget::pane {{
                background: {glass_bg};
                border: 1px solid {border_color};
                border-radius: 12px;
            }}
            
            QTableWidget, QListWidget {{
                color: {text_color};
                gridline-color: {border_color};
            }}
            
            QHeaderView::section {{
                background: {glass_bg};
                color: {text_color};
                padding: 8px;
                border: none;
                font-weight: 600;
            }}
            
            /* Inputs */
            QLineEdit, QSpinBox, QComboBox {{
                background: {input_bg};
                border: 1px solid {border_color};
                border-radius: 8px;
                color: {text_color};
                padding: 8px 12px;
                font-size: 13px;
            }}
            
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{
                border: 1px solid {accent_color};
                background: {input_bg};
            }}
            
            /* Tabs */
            QTabBar::tab {{
                background: {glass_bg};
                color: {secondary_text};
                padding: 8px 16px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 4px;
            }}
            
            QTabBar::tab:selected {{
                background: {accent_color};
                color: white;
            }}
            
            /* Scrollbars */
            QScrollBar:vertical {{
                border: none;
                background: transparent;
                width: 8px;
                margin: 0px;
            }}
            QScrollBar::handle:vertical {{
                background: {border_color};
                min-height: 20px;
                border-radius: 4px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
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
