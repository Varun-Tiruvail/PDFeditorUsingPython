"""
Automation Hub - All Business Logic & Modules
Contains: PDF Editor, OCR Trainer, Scheduler, Database, Utilities
"""
import os
import re
import fitz  # PyMuPDF
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                               QLabel, QFileDialog, QScrollArea, QTableWidget,
                               QTableWidgetItem, QLineEdit, QSpinBox, QComboBox,
                               QTextEdit, QListWidget, QDialog, QDialogButtonBox,
                               QMessageBox, QGraphicsScene, QGraphicsView,
                               QGraphicsRectItem, QTabWidget, QMainWindow, QInputDialog,QApplication,
                               QRubberBand, QMenu, QCheckBox)
from PySide6.QtCore import Qt, QPointF, QRectF, Signal, QThread, QPoint, QRect, QSize
from PySide6.QtGui import QPixmap, QImage, QPen, QColor, QBrush, QPainter
from sqlalchemy import create_engine, Column, Integer, String, Float, ForeignKey, Boolean, DateTime
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.interval import IntervalTrigger
from apscheduler.triggers.date import DateTrigger
import subprocess
import uuid
import pandas as pd
import datetime
import win32com.client
import pythoncom

# ============================================================================
# OFFICE CONVERTER
# ============================================================================

class OfficeConverter:
    @staticmethod
    def convert_to_pdf(input_path):
        """Convert PPT/Excel/Word to PDF using win32com"""
        input_path = os.path.abspath(input_path)
        base, ext = os.path.splitext(input_path)
        output_path = base + "_converted.pdf"
        
        try:
            pythoncom.CoInitialize()
            ext = ext.lower()
            
            if ext in ['.pptx', '.ppt']:
                powerpoint = win32com.client.Dispatch("Powerpoint.Application")
                presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
                presentation.SaveAs(output_path, 32) # 32 = ppSaveAsPDF
                presentation.Close()
                # powerpoint.Quit() # Keep open for performance?
                
            elif ext in ['.xlsx', '.xls']:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Open(input_path)
                wb.ExportAsFixedFormat(0, output_path) # 0 = xlTypePDF
                wb.Close(False)
                # excel.Quit()
                
            elif ext in ['.docx', '.doc']:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(input_path)
                doc.SaveAs(output_path, 17) # 17 = wdFormatPDF
                doc.Close()
                # word.Quit()
                
            return output_path
        except Exception as e:
            print(f"Conversion failed: {e}")
            return None

# ============================================================================
# DATABASE SETUP
# ============================================================================

Base = declarative_base()
DB_PATH = "data/automation_hub.db"
os.makedirs("data", exist_ok=True)
engine = create_engine(f"sqlite:///{DB_PATH}")
SessionLocal = sessionmaker(bind=engine)

class Template(Base):
    __tablename__ = "templates"
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    base_width = Column(Float)
    base_height = Column(Float)
    fields = relationship("Field", back_populates="template", cascade="all, delete-orphan")

class Field(Base):
    __tablename__ = "fields"
    id = Column(Integer, primary_key=True)
    template_id = Column(Integer, ForeignKey("templates.id"))
    name = Column(String)
    x = Column(Float)
    y = Column(Float)
    width = Column(Float)
    height = Column(Float)
    template = relationship("Template", back_populates="fields")

# ============================================================================
# OCR TEMPLATE DATABASE MODELS
# ============================================================================

class OCRTemplate(Base):
    """Template for extracting data from PDFs using labeled boxes"""
    __tablename__ = "ocr_templates"
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)
    pages = relationship("OCRPage", back_populates="template", cascade="all, delete-orphan", order_by="OCRPage.order_index")

class OCRPage(Base):
    """Represents a page in a PDF within a template"""
    __tablename__ = "ocr_pages"
    id = Column(Integer, primary_key=True)
    template_id = Column(Integer, ForeignKey("ocr_templates.id"))
    pdf_filename = Column(String)  # Original filename for reference
    page_number = Column(Integer)
    page_width = Column(Float)
    page_height = Column(Float)
    page_rotation = Column(Integer, default=0)  # 0, 90, 180, or 270
    order_index = Column(Integer)  # Order in the template
    template = relationship("OCRTemplate", back_populates="pages")
    boxes = relationship("LabeledBox", back_populates="page", cascade="all, delete-orphan")

class LabeledBox(Base):
    """A labeled extraction region with optional anchor/value sub-regions"""
    __tablename__ = "labeled_boxes"
    id = Column(Integer, primary_key=True)
    page_id = Column(Integer, ForeignKey("ocr_pages.id"))
    parent_box_id = Column(Integer, ForeignKey("labeled_boxes.id"), nullable=True)
    name = Column(String)
    box_type = Column(String)  # 'label' (parent), 'anchor', 'value'
    x = Column(Float)
    y = Column(Float)
    width = Column(Float)
    height = Column(Float)
    anchor_text = Column(String, nullable=True)  # OCR-captured text for anchor boxes
    page = relationship("OCRPage", back_populates="boxes")
    # Self-referential relationship: parent has many children
    children = relationship("LabeledBox", 
                           foreign_keys="LabeledBox.parent_box_id",
                           back_populates="parent_box",
                           lazy="joined")
    parent_box = relationship("LabeledBox", 
                             remote_side="LabeledBox.id",
                             foreign_keys="LabeledBox.parent_box_id",
                             back_populates="children")


class Job(Base):
    __tablename__ = "jobs"
    id = Column(Integer, primary_key=True)
    name = Column(String)
    script_path = Column(String)
    job_type = Column(String)  # 'one_time' or 'recurring'
    run_date = Column(DateTime, nullable=True)  # For one-time jobs
    recurrence = Column(String, nullable=True)  # 'daily', 'weekly', 'monthly', 'interval'
    interval_seconds = Column(Integer, nullable=True)
    cron_expression = Column(String, nullable=True)
    recurrence_time = Column(String, nullable=True)  # Time of day for daily/weekly/monthly (HH:MM)
    day_of_week = Column(String, nullable=True)  # For weekly (e.g., "0,2,4" for Mon/Wed/Fri)
    day_of_month = Column(Integer, nullable=True)  # For monthly
    last_run = Column(DateTime, nullable=True)
    next_run = Column(DateTime, nullable=True)
    enabled = Column(Boolean, default=True)
    misfire_grace_time = Column(Integer, default=300)  # 5 minutes default

Base.metadata.create_all(engine)

# ============================================================================
# UTILITY CLASSES
# ============================================================================

class PDFCanvas(QLabel):
    """Custom label that supports interactive selection with resize handles"""
    selection_confirmed = Signal(QRect)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)  # Improve responsiveness
        self.selection_mode = False
        
        # State
        self.current_rect = QRect()
        self.drag_start = QPoint()
        self.drag_mode = None  # None, 'create', 'move', 'handle'
        self.active_handle = None
        
        # Appearance
        self.handle_size = 8
        self.border_color = QColor(255, 0, 0)
        self.fill_color = QColor(255, 0, 0, 50)
        self.handle_color = QColor(255, 255, 255)

    def set_selection_mode(self, enabled):
        self.selection_mode = enabled
        if enabled:
            self.setCursor(Qt.CrossCursor)
            self.setFocus()
        else:
            self.setCursor(Qt.ArrowCursor)
            self.current_rect = QRect() # Clear selection
            self.update()

    def paintEvent(self, event):
        super().paintEvent(event) # Draw the pixmap
        
        if self.selection_mode and not self.current_rect.isNull():
            painter = QPainter(self)
            painter.setPen(QPen(self.border_color, 2))
            painter.setBrush(QBrush(self.fill_color))
            
            # Draw Main Rect
            painter.drawRect(self.current_rect)
            
            # Draw Handles
            painter.setBrush(QBrush(self.handle_color))
            painter.setPen(QPen(self.border_color, 1))
            for handle_rect in self._get_handles().values():
                painter.drawRect(handle_rect)

    def _get_handles(self):
        """Calculate handle rectangles for current selection"""
        if self.current_rect.isNull(): return {}
        
        r = self.current_rect
        s = self.handle_size
        hs = s // 2
        
        handles = {
            'tl': QRect(r.left() - hs, r.top() - hs, s, s),
            't':  QRect(r.center().x() - hs, r.top() - hs, s, s),
            'tr': QRect(r.right() - hs, r.top() - hs, s, s),
            'r':  QRect(r.right() - hs, r.center().y() - hs, s, s),
            'br': QRect(r.right() - hs, r.bottom() - hs, s, s),
            'b':  QRect(r.center().x() - hs, r.bottom() - hs, s, s),
            'bl': QRect(r.left() - hs, r.bottom() - hs, s, s),
            'l':  QRect(r.left() - hs, r.center().y() - hs, s, s),
        }
        return handles

    def _get_handle_at(self, pos):
        for name, rect in self._get_handles().items():
            if rect.contains(pos):
                return name
        return None

    def mousePressEvent(self, event):
        if not self.selection_mode or event.button() != Qt.LeftButton:
            return
            
        pos = event.position().toPoint()
        
        # Check handles first
        handle = self._get_handle_at(pos)
        if handle:
            self.drag_mode = 'handle'
            self.active_handle = handle
            self.drag_start = pos
            return
            
        # Check move
        if self.current_rect.contains(pos):
            self.drag_mode = 'move'
            self.drag_start = pos
            self.setCursor(Qt.SizeAllCursor)
            return
            
        # Create new
        self.drag_mode = 'create'
        self.drag_start = pos
        self.current_rect = QRect(pos, QSize())
        self.update()

    def mouseMoveEvent(self, event):
        if not self.selection_mode: return

        pos = event.position().toPoint()
        
        # Update cursor hover feedback
        if not self.drag_mode:
            handle = self._get_handle_at(pos)
            if handle:
                if handle in ['tl', 'br']: self.setCursor(Qt.SizeFDiagCursor)
                elif handle in ['tr', 'bl']: self.setCursor(Qt.SizeBDiagCursor)
                elif handle in ['l', 'r']: self.setCursor(Qt.SizeHorCursor)
                elif handle in ['t', 'b']: self.setCursor(Qt.SizeVerCursor)
            elif self.current_rect.contains(pos):
                self.setCursor(Qt.SizeAllCursor)
            else:
                self.setCursor(Qt.CrossCursor)
            return

        # Handle Dragging
        dx = pos.x() - self.drag_start.x()
        dy = pos.y() - self.drag_start.y()
        
        if self.drag_mode == 'create':
            self.current_rect = QRect(self.drag_start, pos).normalized()
            
        elif self.drag_mode == 'move':
            self.current_rect.translate(dx, dy)
            self.drag_start = pos
            
        elif self.drag_mode == 'handle':
            r = self.current_rect
            # Adjust specific edges based on handle
            if 'l' in self.active_handle: r.setLeft(r.left() + dx)
            if 'r' in self.active_handle: r.setRight(r.right() + dx)
            if 't' in self.active_handle: r.setTop(r.top() + dy)
            if 'b' in self.active_handle: r.setBottom(r.bottom() + dy)
            self.current_rect = r.normalized()
            self.drag_start = pos
            
        self.update()

    def mouseReleaseEvent(self, event):
        if self.selection_mode and event.button() == Qt.LeftButton:
            self.drag_mode = None
            self.active_handle = None
            self.update() # Refreshes handles position
            
            # Ensure 0-size rects are ignored but don't finish yet
            if self.current_rect.width()<5 and self.current_rect.height()<5:
                self.current_rect = QRect()
                
    def keyPressEvent(self, event):
        if not self.selection_mode:
            super().keyPressEvent(event)
            return
            
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            if not self.current_rect.isNull():
                self.selection_confirmed.emit(self.current_rect)
        elif event.key() == Qt.Key_Escape:
            self.current_rect = QRect()
            self.update()
            # Optionally exit mode? For now just clear selection

# ============================================================================
# PDF EDITOR MODULE
# ============================================================================

class PDFTab(QWidget):
    def __init__(self, doc, path=None, is_temp=False, temp_path=None):
        super().__init__()
        self.doc = doc
        self.path = path
        self.current_page = 0
        self.scale = 1.5
        self.is_temp = is_temp
        self.temp_path = temp_path
        self.parent_dock = None  # Will be set by PDFEditorModule
        self.setup_ui()
        self.setFocusPolicy(Qt.ClickFocus)

    def focusInEvent(self, event):
        super().focusInEvent(event)
        # Notify parent PDFEditorModule that this tab is active
        parent = self.parent()
        while parent and not isinstance(parent, PDFEditorModule):
            parent = parent.parent()
        if parent:
            parent._last_active_tab = self

    def mousePressEvent(self, event):
        self.setFocus()
        super().mousePressEvent(event)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Toolbar Container
        toolbar_layout = QHBoxLayout()
        
        # Navigation Toolbar (Left)
        nav_layout = QHBoxLayout()
        nav_layout.setAlignment(Qt.AlignLeft)
        
        self.btn_prev = QPushButton("◀")
        self.btn_prev.setFixedSize(40, 30)
        self.btn_prev.clicked.connect(self.prev_page)
        
        self.lbl_page = QLabel("Page 0 / 0")
        self.lbl_page.setStyleSheet("font-weight: bold; color: #ccc; padding: 0 10px;")
        
        self.btn_next = QPushButton("▶")
        self.btn_next.setFixedSize(40, 30)
        self.btn_next.clicked.connect(self.next_page)
        
        nav_layout.addWidget(self.btn_prev)
        nav_layout.addWidget(self.lbl_page)
        nav_layout.addWidget(self.btn_next)
        
        # File Name (Center)
        self.lbl_filename = QLabel(os.path.basename(self.path) if self.path else "Untitled")
        self.lbl_filename.setStyleSheet("font-weight: bold; color: #fff; padding: 0 20px; font-size: 14px;")
        self.lbl_filename.setAlignment(Qt.AlignCenter)
        
        #  Zoom Toolbar (Right)
        zoom_layout = QHBoxLayout()
        zoom_layout.setAlignment(Qt.AlignRight)
        
        self.btn_zoom_out = QPushButton("−")
        self.btn_zoom_out.setFixedSize(40, 30)
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        
        self.lbl_zoom = QLabel("150%")
        self.lbl_zoom.setStyleSheet("font-weight: bold; color: #ccc; padding: 0 10px;")
        
        self.btn_zoom_in = QPushButton("+")
        self.btn_zoom_in.setFixedSize(40, 30)
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        
        self.btn_fit_width = QPushButton("Fit W")
        self.btn_fit_width.setFixedSize(55, 30)
        self.btn_fit_width.clicked.connect(self.fit_to_width)
        
        self.btn_fit_height = QPushButton("Fit H")
        self.btn_fit_height.setFixedSize(55, 30)
        self.btn_fit_height.clicked.connect(self.fit_to_height)
        
        self.btn_fit = QPushButton("Fit")
        self.btn_fit.setFixedSize(50, 30)
        self.btn_fit.clicked.connect(self.fit_to_screen)
        
        self.btn_close = QPushButton("✖")
        self.btn_close.setFixedSize(40, 30)
        self.btn_close.clicked.connect(self.close_self)
        self.btn_close.setStyleSheet("background-color: #dc2626; color: white;")
        
        self.btn_popout = QPushButton("⬜")
        self.btn_popout.setFixedSize(40, 30)
        self.btn_popout.clicked.connect(self.pop_out)
        
        zoom_layout.addWidget(self.btn_zoom_out)
        zoom_layout.addWidget(self.lbl_zoom)
        zoom_layout.addWidget(self.btn_zoom_in)
        zoom_layout.addWidget(self.btn_fit_width)
        zoom_layout.addWidget(self.btn_fit_height)
        zoom_layout.addWidget(self.btn_fit)
        zoom_layout.addWidget(self.btn_close)
        zoom_layout.addWidget(self.btn_popout)
        
        # Combine toolbars
        toolbar_layout.addLayout(nav_layout)
        toolbar_layout.addWidget(self.lbl_filename, stretch=1)
        toolbar_layout.addLayout(zoom_layout)
        layout.addLayout(toolbar_layout)
        
        # Scroll Area
        self.scroll = QScrollArea()
        self.label = PDFCanvas()
        self.label.setAlignment(Qt.AlignCenter)

        self.scroll.setWidget(self.label)
        self.scroll.setWidgetResizable(True)
        layout.addWidget(self.scroll)
        
        self.render()


    
    def zoom_in(self):
        self.scale *= 1.2
        self.update_zoom_label()
        self.render()
    
    def zoom_out(self):
        self.scale /= 1.2
        self.update_zoom_label()
        self.render()
    
    def fit_to_screen(self):
        """Fit to width (same as fit_to_width for backward compatibility)"""
        self.fit_to_width()
    
    def fit_to_width(self):
        if not self.doc: return
        try:
            page = self.doc.load_page(self.current_page)
            page_width = page.rect.width
            scroll_width = self.scroll.width() - 40  # Account for margins
            self.scale = scroll_width / page_width
            self.update_zoom_label()
            self.render()
        except Exception as e:
            print(f"Fit width error: {e}")
    
    def fit_to_height(self):
        if not self.doc: return
        try:
            page = self.doc.load_page(self.current_page)
            page_height = page.rect.height
            scroll_height = self.scroll.height() - 40  # Account for margins
            self.scale = scroll_height / page_height
            self.update_zoom_label()
            self.render()
        except Exception as e:
            print(f"Fit height error: {e}")
    
    def close_self(self):
        """Close this dock"""
        if self.parent_dock:
            # Find parent PDFEditorModule
            parent = self.parent()
            while parent and not isinstance(parent, PDFEditorModule):
                parent = parent.parent()
            if parent:
                parent.close_tab(self.parent_dock)
    
    def pop_out(self):
        """Pop out to floating window"""
        if self.parent_dock:
            self.parent_dock.setFloating(True)
    
    def update_zoom_label(self):
        zoom_pct = int(self.scale * 100)
        self.lbl_zoom.setText(f"{zoom_pct}%")

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render()

    def next_page(self):
        if self.doc and self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.render()

    def render(self):
        if not self.doc: return
        try:
            # Update Page Label
            total_pages = len(self.doc)
            self.lbl_page.setText(f"Page {self.current_page + 1} / {total_pages}")
            
            # Enable/Disable buttons
            self.btn_prev.setEnabled(self.current_page > 0)
            self.btn_next.setEnabled(self.current_page < total_pages - 1)
            
            page = self.doc.load_page(self.current_page)
            pix = page.get_pixmap(matrix=fitz.Matrix(self.scale, self.scale))
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            self.label.setPixmap(QPixmap.fromImage(img))
        except Exception as e:
            print(f"Render error: {e}")
    
    def cleanup(self):
        """Clean up temp files and close document"""
        if self.doc:
            try:
                self.doc.close()
                self.doc = None
            except Exception as e:
                print(f"Failed to close doc: {e}")

        if self.is_temp and self.temp_path and os.path.exists(self.temp_path):
            try:
                os.remove(self.temp_path)
                print(f"Deleted temp file: {self.temp_path}")
            except Exception as e:
                print(f"Failed to delete temp file: {e}")


class PDFEditorModule(QWidget):
    def __init__(self):
        super().__init__()
        # Create temp directory
        self.temp_dir = os.path.join(os.getcwd(), ".temp_pdfs")
        os.makedirs(self.temp_dir, exist_ok=True)
        self._last_active_tab = None
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Title
        title = QLabel("📄 PDF Editor")
        title.setObjectName("moduleTitle")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        layout.addWidget(title)

        # Toolbar
        toolbar = QHBoxLayout()
        toolbar.setSpacing(10)

        self.btn_open = self.create_btn("📂 Open (PDF/Office)", self.open_pdf)
        self.btn_save = self.create_btn("💾 Save", lambda: self.save_pdf())
        self.btn_close_all = self.create_btn("❌ Close All", self.close_all)
        self.btn_ppt = self.create_btn("📊 PPT to PDF", self.ppt_to_pdf)
        self.btn_compress = self.create_btn("🗜️ Compress", self.compress_pdf)
        self.btn_merge = self.create_btn("📑 Merge", self.merge_pdfs)
        self.btn_split = self.create_btn("✂️ Split", self.split_pdf)
        self.btn_redact_custom = self.create_btn("🎯 Redact Custom", self.redact_custom_location)
        self.btn_pagenum = self.create_btn("🔢 Add Page #", self.add_page_numbers)
        self.btn_header = self.create_btn("📝 Header/Footer", self.add_header_footer)
        self.btn_advanced = self.create_btn("🔧 Advanced Tools", self.show_advanced_menu)
        
        for btn in [self.btn_open, self.btn_save, self.btn_close_all, self.btn_ppt, self.btn_compress, self.btn_merge, self.btn_split, 
                   self.btn_redact_custom, self.btn_pagenum, self.btn_header, self.btn_advanced]:
            toolbar.addWidget(btn)
        toolbar.addStretch()
        layout.addLayout(toolbar)
        
        # Dock Manager (QMainWindow embedded)
        self.dock_manager = QMainWindow()
        self.dock_manager.setWindowFlags(Qt.Widget) # Embeddable
        self.dock_manager.setDockOptions(
            QMainWindow.AllowTabbedDocks | 
            QMainWindow.AllowNestedDocks | 
            QMainWindow.AnimatedDocks |
            QMainWindow.GroupedDragging
        )
        
        # Enable all dock orientations
        self.dock_manager.setCorner(Qt.TopLeftCorner, Qt.LeftDockWidgetArea)
        self.dock_manager.setCorner(Qt.TopRightCorner, Qt.RightDockWidgetArea)
        self.dock_manager.setCorner(Qt.BottomLeftCorner, Qt.LeftDockWidgetArea)
        self.dock_manager.setCorner(Qt.BottomRightCorner, Qt.RightDockWidgetArea)
        
        # Set tab position to bottom
        self.dock_manager.setTabPosition(Qt.AllDockWidgetAreas, QTabWidget.South)
        
        # Central widget (minimal size to allow splits)
        self.central_widget = QWidget()
        self.central_widget.setMaximumSize(1, 1)
        self.central_widget.setStyleSheet("background: transparent;")
        self.dock_manager.setCentralWidget(self.central_widget)
        
        layout.addWidget(self.dock_manager)
        
        # Track open docs
        self.docks = []

    def create_btn(self, text, callback):
        btn = QPushButton(text)
        btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #667EEA, stop:1 #764BA2);
                color: white; border: none; padding: 10px 20px;
                border-radius: 6px; font-size: 14px; font-weight: 500;
            }
            QPushButton:hover { background: #764BA2; }
        """)
        btn.clicked.connect(callback)
        return btn
    
    def current_tab(self):
        from PySide6.QtWidgets import QApplication, QTabBar
        
        # 1. Check if any tab's child widget currently has focus
        focus_widget = QApplication.focusWidget()
        if focus_widget:
            for dock in self.docks:
                tab = dock.widget()
                if tab == focus_widget or tab.isAncestorOf(focus_widget):
                    self._last_active_tab = tab
                    return tab
        
        # 2. Find the active dock in a tabbed group by finding tab bars
        # In Qt, when docks are tabbed, there's a QTabBar child of the dock area
        for child in self.dock_manager.findChildren(QTabBar):
            current_index = child.currentIndex()
            if current_index >= 0:
                # Get the text of the current tab to match with dock titles
                tab_text = child.tabText(current_index)
                for dock in self.docks:
                    if dock.windowTitle() == tab_text:
                        self._last_active_tab = dock.widget()
                        return dock.widget()
        
        # 3. Fall back to last known active tab
        if self._last_active_tab and self._last_active_tab in [d.widget() for d in self.docks]:
            return self._last_active_tab

        # 4. Fallback: Return first visible dock
        for dock in self.docks:
            if dock.isVisible() and not dock.isHidden():
                return dock.widget()
        
        # 5. Last resort: just the latest dock
        if self.docks:
            return self.docks[-1].widget()
        return None

    def close_tab(self, dock):
        if dock in self.docks:
            # Cleanup temp files
            tab = dock.widget()
            if tab and hasattr(tab, 'cleanup'):
                tab.cleanup()
            
            # Check for unsaved changes (mockup)
            reply = QMessageBox.question(self, "Close", "Save changes before closing?", 
                                       QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                return
            if reply == QMessageBox.Yes:
                self.save_pdf(dock.widget())
            
            self.dock_manager.removeDockWidget(dock)
            dock.deleteLater()
            self.docks.remove(dock)

    def close_all(self):
        reply = QMessageBox.question(self, "Close All", "Close all tabs without saving?", 
                                   QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            for dock in list(self.docks):
                self.dock_manager.removeDockWidget(dock)
                dock.deleteLater()
            self.docks.clear()
            self._last_active_tab = None

    def on_dock_visibility_changed(self, visible):
        """Track active dock using visibility signals"""
        if visible:
            dock = self.sender()
            if dock and not dock.isFloating():
                self._last_active_tab = dock.widget()

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Files (*.pdf *.pptx *.xlsx *.docx)")
        if path:
            try:
                is_temp = False
                temp_path = None
                original_path = path
                
                # Convert if Office file
                if path.lower().endswith(('.pptx', '.xlsx', '.docx')):
                    # Generate temp filename
                    import uuid
                    temp_filename = f"{uuid.uuid4().hex}.pdf"
                    temp_path = os.path.join(self.temp_dir, temp_filename)
                    
                    # Convert to temp location
                    import shutil
                    converted_path = OfficeConverter.convert_to_pdf(path)
                    if not converted_path:
                        raise Exception("Conversion failed")
                    
                    shutil.move(converted_path, temp_path)
                    path = temp_path
                    is_temp = True
                
                doc = fitz.open(path)
                tab = PDFTab(doc, original_path, is_temp=is_temp, temp_path=temp_path)
                
                # Create Dock Widget
                from PySide6.QtWidgets import QDockWidget
                dock = QDockWidget(os.path.basename(original_path), self)
                dock.setWidget(tab)
                dock.setAllowedAreas(Qt.AllDockWidgetAreas)
                dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable | QDockWidget.DockWidgetClosable)
                
                # Set parent_dock reference
                tab.parent_dock = dock
                
                # Rename feature via context menu
                dock.setContextMenuPolicy(Qt.CustomContextMenu)
                dock.customContextMenuRequested.connect(lambda pos, d=dock: self.dock_context_menu(pos, d))
                
                # Signal for active tab tracking
                dock.visibilityChanged.connect(self.on_dock_visibility_changed)

                # Connect interactive selection signal
                tab.label.selection_confirmed.connect(lambda rect: self.apply_custom_redaction(tab, rect))
                
                self.dock_manager.addDockWidget(Qt.RightDockWidgetArea, dock)
                if self.docks:
                    self.dock_manager.tabifyDockWidget(self.docks[-1], dock)
                
                self.docks.append(dock)
                dock.show()
                # Explicitly set as active if it's the only one
                if len(self.docks) == 1:
                    self._last_active_tab = tab
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open file: {e}")
    
    def dock_context_menu(self, pos, dock):
        from PySide6.QtWidgets import QMenu
        menu = QMenu()
        rename_action = menu.addAction("Rename")
        close_action = menu.addAction("Close")
        action = menu.exec(dock.mapToGlobal(pos))
        if action == rename_action:
            new_name, ok = QInputDialog.getText(self, "Rename", "New Name:", text=dock.windowTitle())
            if ok and new_name:
                dock.setWindowTitle(new_name)
        elif action == close_action:
            self.close_tab(dock)

    def save_pdf(self, tab=None):
        if not tab: tab = self.current_tab()
        if not tab: return
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
        if path:
            try:
                tab.doc.save(path)
                QMessageBox.information(self, "Success", "PDF saved successfully!")
                # Update dock title
                for dock in self.docks:
                    if dock.widget() == tab:
                        dock.setWindowTitle(os.path.basename(path))
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
    
    def ppt_to_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select PPT", "", "PowerPoint (*.pptx *.ppt)")
        if path:
            pdf_path = OfficeConverter.convert_to_pdf(path)
            if pdf_path:
                QMessageBox.information(self, "Success", f"Converted to: {pdf_path}")
            else:
                QMessageBox.critical(self, "Error", "Conversion failed")

    def compress_pdf(self):
        tab = self.current_tab()
        if not tab: return
        
        path, _ = QFileDialog.getSaveFileName(self, "Save Compressed PDF", "", "PDF Files (*.pdf)")
        if path:
            try:
                # Save compressed to new file
                tab.doc.save(path, garbage=4, deflate=True)
                # Open result in new tab
                new_doc = fitz.open(path)
                new_tab = PDFTab(new_doc, path)
                
                # Create Dock Widget
                from PySide6.QtWidgets import QDockWidget
                dock = QDockWidget(os.path.basename(path), self)
                dock.setWidget(new_tab)
                dock.setAllowedAreas(Qt.AllDockWidgetAreas)
                dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable | QDockWidget.DockWidgetClosable)
                dock.setContextMenuPolicy(Qt.CustomContextMenu)
                dock.customContextMenuRequested.connect(lambda pos, d=dock: self.dock_context_menu(pos, d))
                
                # Set parent_dock reference
                new_tab.parent_dock = dock
                
                self.dock_manager.addDockWidget(Qt.RightDockWidgetArea, dock)
                if self.docks:
                    self.dock_manager.tabifyDockWidget(self.docks[-1], dock)
                self.docks.append(dock)
                dock.show()
                
                QMessageBox.information(self, "Success", "Compressed PDF opened in new tab!")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def merge_pdfs(self):
        """Show merge options: Simple or Header-Based"""
        choice_dialog = QDialog(self)
        choice_dialog.setWindowTitle("Choose Merge Type")
        choice_dialog.resize(400, 200)
        layout = QVBoxLayout(choice_dialog)
        
        layout.addWidget(QLabel("<h3>How would you like to merge PDFs?</h3>"))
        
        btn_simple = QPushButton("📑 Simple Merge with Page Rearranging")
        btn_simple.clicked.connect(lambda: (choice_dialog.accept(), self.merge_simple()))
        layout.addWidget(btn_simple)
        
        btn_headers = QPushButton("📌 Header-Based Merge (Insert PDFs after headers)")
        btn_headers.clicked.connect(lambda: (choice_dialog.accept(), self.merge_with_headers()))
        layout.addWidget(btn_headers)
        
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(choice_dialog.reject)
        layout.addWidget(btn_cancel)
        
        choice_dialog.exec()
    
    def merge_simple(self):
        """Simple merge with page-level rearranging"""
        from PySide6.QtWidgets import QListWidgetItem
        from PySide6.QtCore import QSize
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Simple Merge - Arrange Pages")
        dialog.resize(800, 600)
        layout = QVBoxLayout(dialog)
        
        # Layout: Left side for PDF list, Right side for Thumbnails
        content_layout = QHBoxLayout()
        
        # LEFT PANEL: PDF List
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.addWidget(QLabel("<b>1. Add PDFs</b>"))
        
        pdf_listwidget = QListWidget()
        pdf_listwidget.setDragDropMode(QListWidget.InternalMove)
        left_layout.addWidget(pdf_listwidget)
        
        btn_add = QPushButton("+ Add PDFs")
        def add_pdfs():
            files, _ = QFileDialog.getOpenFileNames(self, "Select PDFs", "", "PDF Files (*.pdf)")
            for f in files:
                pdf_listwidget.addItem(f)
        btn_add.clicked.connect(add_pdfs)
        left_layout.addWidget(btn_add)
        
        btn_load_pages = QPushButton("Load Pages →")
        left_layout.addWidget(btn_load_pages)
        left_layout.addStretch()
        
        content_layout.addWidget(left_panel, stretch=1)
        
        # RIGHT PANEL: Page Thumbnails
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.addWidget(QLabel("<b>2. Arrange Pages</b>"))
        
        page_listwidget = QListWidget()
        page_listwidget.setDragDropMode(QListWidget.InternalMove)
        page_listwidget.setViewMode(QListWidget.ListMode)
        page_listwidget.setIconSize(QSize(100, 140))
        page_listwidget.setSpacing(5)
        right_layout.addWidget(page_listwidget)
        
        # Move Buttons
        btn_layout = QHBoxLayout()
        btn_up = QPushButton("▲ Move Up")
        btn_down = QPushButton("▼ Move Down")
        
        def move_item(direction):
            row = page_listwidget.currentRow()
            if row < 0: return
            
            new_row = row + direction
            if 0 <= new_row < page_listwidget.count():
                item = page_listwidget.takeItem(row)
                page_listwidget.insertItem(new_row, item)
                page_listwidget.setCurrentRow(new_row)
        
        btn_up.clicked.connect(lambda: move_item(-1))
        btn_down.clicked.connect(lambda: move_item(1))
        
        btn_layout.addWidget(btn_up)
        btn_layout.addWidget(btn_down)
        right_layout.addWidget(page_listwidget)
        right_layout.addLayout(btn_layout)
        
        content_layout.addWidget(right_panel, stretch=2)
        layout.addLayout(content_layout)
        
        # Load pages logic
        def load_pages():
            page_listwidget.clear()
            for i in range(pdf_listwidget.count()):
                pdf_path = pdf_listwidget.item(i).text()
                try:
                    doc = fitz.open(pdf_path)
                    pdf_name = os.path.basename(pdf_path)
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        pix = page.get_pixmap(matrix=fitz.Matrix(0.3, 0.3))
                        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                        item = QListWidgetItem(QPixmap.fromImage(img), f"{pdf_name}\nP{page_num + 1}")
                        item.setData(Qt.UserRole, (i, page_num))
                        page_listwidget.addItem(item)
                    doc.close()
                except Exception as e:
                    print(f"Error: {e}")
        btn_load_pages.clicked.connect(load_pages)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec() == QDialog.Accepted and page_listwidget.count() > 0:
            try:
                merged = fitz.open()
                pdf_docs = [fitz.open(pdf_listwidget.item(i).text()) for i in range(pdf_listwidget.count())]
                
                for i in range(page_listwidget.count()):
                    item = page_listwidget.item(i)
                    pdf_idx, page_num = item.data(Qt.UserRole)
                    merged.insert_pdf(pdf_docs[pdf_idx], from_page=page_num, to_page=page_num)
                
                for doc in pdf_docs:
                    doc.close()
                
                tab = PDFTab(merged, "Merged.pdf")
                
                # Create Dock Widget
                from PySide6.QtWidgets import QDockWidget
                dock = QDockWidget("Merged.pdf", self)
                dock.setWidget(tab)
                dock.setAllowedAreas(Qt.AllDockWidgetAreas)
                dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable | QDockWidget.DockWidgetClosable)
                dock.setContextMenuPolicy(Qt.CustomContextMenu)
                dock.customContextMenuRequested.connect(lambda pos, d=dock: self.dock_context_menu(pos, d))
                
                # Set parent_dock reference
                tab.parent_dock = dock
                
                self.dock_manager.addDockWidget(Qt.RightDockWidgetArea, dock)
                if self.docks:
                    self.dock_manager.tabifyDockWidget(self.docks[-1], dock)
                self.docks.append(dock)
                dock.show()
                
                QMessageBox.information(self, "Success", f"Merged {page_listwidget.count()} pages!")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
    
    def merge_with_headers(self):
        """Header-based merge: Insert PDFs after specific header pages"""
        from PySide6.QtWidgets import QListWidgetItem, QStackedWidget
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Header-Based Merge")
        dialog.resize(700, 600)
        layout = QVBoxLayout(dialog)
        
        stack = QStackedWidget()
        layout.addWidget(stack)
        
        # Nav buttons
        nav_layout = QHBoxLayout()
        btn_back = QPushButton("◀ Back")
        btn_next = QPushButton("Next ▶")
        btn_merge = QPushButton("✓ Merge")
        btn_merge.setVisible(False)
        nav_layout.addWidget(btn_back)
        nav_layout.addStretch()
        nav_layout.addWidget(btn_next)
        nav_layout.addWidget(btn_merge)
        layout.addLayout(nav_layout)
        
        base_pdf = None
        headers = []  # [(page_num, label), ...]
        insertions = {}  # {page_num: [pdf_paths]}
        
        # STEP 1: Select base PDF
        step1 = QWidget()
        step1_layout = QVBoxLayout(step1)
        step1_layout.addWidget(QLabel("<h2>Step 1: Select Base PDF</h2>"))
        step1_layout.addWidget(QLabel("This PDF contains the header pages:"))
        
        base_label = QLabel("No PDF selected")
        step1_layout.addWidget(base_label)
        
        def select_base():
            nonlocal base_pdf
            path, _ = QFileDialog.getOpenFileName(dialog, "Select Base PDF", "", "PDF Files (*.pdf)")
            if path:
                base_pdf = path
                base_label.setText(f"Selected: {os.path.basename(path)}")
        
        btn_select_base = QPushButton("Browse...")
        btn_select_base.clicked.connect(select_base)
        step1_layout.addWidget(btn_select_base)
        stack.addWidget(step1)
        
        # STEP 2: Mark headers
        step2 = QWidget()
        step2_layout = QVBoxLayout(step2)
        step2_layout.addWidget(QLabel("<h2>Step 2: Mark Header Pages</h2>"))
        
        header_scroll = QScrollArea()
        header_container = QWidget()
        header_container_layout = QVBoxLayout(header_container)
        header_scroll.setWidget(header_container)
        header_scroll.setWidgetResizable(True)
        step2_layout.addWidget(header_scroll)
        stack.addWidget(step2)
        
        # STEP 3: Insert PDFs after headers
        step3 = QWidget()
        step3_layout = QVBoxLayout(step3)
        step3_layout.addWidget(QLabel("<h2>Step 3: Insert PDFs After Headers</h2>"))
        
        insert_scroll = QScrollArea()
        insert_container = QWidget()
        insert_container_layout = QVBoxLayout(insert_container)
        insert_scroll.setWidget(insert_container)
        insert_scroll.setWidgetResizable(True)
        step3_layout.addWidget(insert_scroll)
        stack.addWidget(step3)
        
        # Navigation
        def go_step1():
            stack.setCurrentIndex(0)
            btn_back.setVisible(False)
            btn_next.setVisible(True)
            btn_merge.setVisible(False)
        
        def go_step2():
            if not base_pdf:
                QMessageBox.warning(dialog, "Warning", "Please select a base PDF")
                return
            
            # Load base PDF and show pages
            for i in reversed(range(header_container_layout.count())):
                header_container_layout.itemAt(i).widget().deleteLater()
            
            try:
                doc = fitz.open(base_pdf)
                from PySide6.QtWidgets import QCheckBox
                for page_num in range(len(doc)):
                    row = QHBoxLayout()
                    cb = QCheckBox(f"Page {page_num + 1}")
                    cb.setProperty("page_num", page_num)
                    label_input = QLineEdit()
                    label_input.setPlaceholderText("Header label (e.g., 'Section 1')")
                    label_input.setEnabled(False)
                    
                    # Fix: Use a separate function to capture closure correctly
                    def connect_cb(checkbox, input_field):
                        checkbox.stateChanged.connect(lambda state: input_field.setEnabled(state == 2))
                    
                    connect_cb(cb, label_input)
                    
                    row.addWidget(cb)
                    row.addWidget(label_input)
                    
                    widget = QWidget()
                    widget.setLayout(row)
                    widget.setProperty("checkbox", cb)
                    widget.setProperty("label_input", label_input)
                    header_container_layout.addWidget(widget)
                
                doc.close()
            except Exception as e:
                QMessageBox.critical(dialog, "Error", str(e))
                return
            
            stack.setCurrentIndex(1)
            btn_back.setVisible(True)
            btn_next.setVisible(True)
            btn_merge.setVisible(False)
        
        def go_step3():
            # Collect headers
            nonlocal headers
            headers = []
            
            for i in range(header_container_layout.count()):
                widget = header_container_layout.itemAt(i).widget()
                cb = widget.property("checkbox")
                label_inp = widget.property("label_input")
                
                if cb and cb.isChecked():
                    page_num = cb.property("page_num")
                    label = label_inp.text() if label_inp and label_inp.text() else f"Header {len(headers) + 1}"
                    headers.append((page_num, label))
            
            if not headers:
                QMessageBox.warning(dialog, "Warning", "Please mark at least one header page")
                return
            
            headers.sort()  # Sort by page number
            
            # Build insertion UI
            for i in reversed(range(insert_container_layout.count())):
                insert_container_layout.itemAt(i).widget().deleteLater()
            
            from PySide6.QtWidgets import QFrame
            for page_num, label in headers:
                group = QFrame()
                group.setFrameStyle(QFrame.Box)
                group_layout = QVBoxLayout(group)
                group_layout.addWidget(QLabel(f"<b>📌 After '{label}' (Page {page_num + 1})</b>"))
                
                list_widget = QListWidget()
                list_widget.setProperty("page_num", page_num)
                
                btn_add_pdfs = QPushButton("+ Add PDFs")
                def add_pdfs_for_header(pg=page_num, lst=list_widget):
                    files, _ = QFileDialog.getOpenFileNames(dialog, "Select PDFs", "", "PDF Files (*.pdf)")
                    for f in files:
                        lst.addItem(f)
                
                btn_add_pdfs.clicked.connect(add_pdfs_for_header)
                
                group_layout.addWidget(list_widget)
                group_layout.addWidget(btn_add_pdfs)
                insert_container_layout.addWidget(group)
            
            stack.setCurrentIndex(2)
            btn_back.setVisible(True)
            btn_next.setVisible(False)
            btn_merge.setVisible(True)
        
        def do_merge():
            # Collect insertion data
            nonlocal insertions
            insertions = {}
            
            for i in range(insert_container_layout.count()):
                group_widget = insert_container_layout.itemAt(i).widget()
                if not group_widget: continue
                
                # Find the list widget
                for j in range(group_widget.layout().count()):
                    item = group_widget.layout().itemAt(j)
                    if not item: continue
                    widget = item.widget()
                    if isinstance(widget, QListWidget):
                        page_num = widget.property("page_num")
                        pdfs = [widget.item(k).text() for k in range(widget.count())]
                        if pdfs:
                            insertions[page_num] = pdfs
            
            dialog.accept()
        
        btn_back.clicked.connect(lambda: go_step1() if stack.currentIndex() == 1 else go_step2())
        btn_next.clicked.connect(lambda: go_step2() if stack.currentIndex() == 0 else go_step3())
        btn_merge.clicked.connect(do_merge)
        
        go_step1()
        
        if dialog.exec() == QDialog.Accepted:
            try:
                # Build final merged PDF
                base_doc = fitz.open(base_pdf)
                merged = fitz.open()
                
                for page_num in range(len(base_doc)):
                    # Insert base page
                    merged.insert_pdf(base_doc, from_page=page_num, to_page=page_num)
                    
                    # If this is a header, insert PDFs after it
                    if page_num in insertions:
                        for pdf_path in insertions[page_num]:
                            insert_doc = fitz.open(pdf_path)
                            merged.insert_pdf(insert_doc)
                            insert_doc.close()
                
                base_doc.close()
                
                tab = PDFTab(merged, "Merged_Headers.pdf")
                
                # Create Dock Widget
                from PySide6.QtWidgets import QDockWidget
                dock = QDockWidget("Merged_Headers.pdf", self)
                dock.setWidget(tab)
                dock.setAllowedAreas(Qt.AllDockWidgetAreas)
                dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable | QDockWidget.DockWidgetClosable)
                dock.setContextMenuPolicy(Qt.CustomContextMenu)
                dock.customContextMenuRequested.connect(lambda pos, d=dock: self.dock_context_menu(pos, d))
                
                # Set parent_dock reference
                tab.parent_dock = dock
                
                self.dock_manager.addDockWidget(Qt.RightDockWidgetArea, dock)
                if self.docks:
                    self.dock_manager.tabifyDockWidget(self.docks[-1], dock)
                self.docks.append(dock)
                dock.show()
                
                QMessageBox.information(self, "Success", "Header-based merge complete!")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def split_pdf(self):
        """Dynamic PDF split with user-specified page ranges"""
        tab = self.current_tab()
        if not tab: return
        
        total_pages = len(tab.doc)
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Split PDF")
        layout = QVBoxLayout(dialog)
        
        layout.addWidget(QLabel(f"<h3>Split PDF ({total_pages} pages)</h3>"))
        layout.addWidget(QLabel("Enter page ranges (e.g., '1-3, 5, 7-10'):"))
        
        range_input = QLineEdit()
        range_input.setPlaceholderText("1-3, 5-7")
        layout.addWidget(range_input)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec() == QDialog.Accepted:
            range_str = range_input.text().strip()
            if not range_str:
                QMessageBox.warning(self, "Warning", "Please enter page ranges")
                return
            
            try:
                # Parse ranges
                pages = set()
                for part in range_str.split(','):
                    part = part.strip()
                    if '-' in part:
                        start, end = map(int, part.split('-'))
                        pages.update(range(start - 1, end))  # 0-indexed
                    else:
                        pages.add(int(part) - 1)
                
                # Validate
                pages = sorted([p for p in pages if 0 <= p < total_pages])
                
                if not pages:
                    QMessageBox.warning(self, "Warning", "No valid pages selected")
                    return
                
                # Create split PDF
                new_doc = fitz.open()
                for page_num in pages:
                    new_doc.insert_pdf(tab.doc, from_page=page_num, to_page=page_num)
                
                new_tab = PDFTab(new_doc, "Split.pdf")
                
                # Create Dock Widget
                from PySide6.QtWidgets import QDockWidget
                dock = QDockWidget("Split.pdf", self)
                dock.setWidget(new_tab)
                dock.setAllowedAreas(Qt.AllDockWidgetAreas)
                dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable | QDockWidget.DockWidgetClosable)
                dock.setContextMenuPolicy(Qt.CustomContextMenu)
                dock.customContextMenuRequested.connect(lambda pos, d=dock: self.dock_context_menu(pos, d))
                
                # Set parent_dock reference
                new_tab.parent_dock = dock
                
                self.dock_manager.addDockWidget(Qt.RightDockWidgetArea, dock)
                if self.docks:
                    self.dock_manager.tabifyDockWidget(self.docks[-1], dock)
                self.docks.append(dock)
                dock.show()
                
                QMessageBox.information(self, "Success", f"Split {len(pages)} pages into new tab!")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def redact_custom_location(self):
        """Custom redaction - works with rotated pages via coordinate transformation"""
        self.redact_mode = "standard"
        tab = self.current_tab()
        if not tab:
            QMessageBox.warning(self, "No PDF", "Please open a PDF first.")
            return
        
        # Check if any page has rotation and inform user
        has_rotation = any(page.rotation != 0 for page in tab.doc)
        if has_rotation:
            QMessageBox.information(
                self, "Note",
                "Some pages have rotation. Redaction coordinates will be automatically transformed to match the visual position."
            )
        
        # Enable selection mode for redaction
        tab.label.set_selection_mode(True)
        tab.label.setCursor(Qt.CrossCursor)
        QMessageBox.information(self, "Custom Redaction", "Draw a box around the area you wish to redact. A preview will be shown before applying.")


    def prepare_rasterize_redaction(self):
        """Start coordinate selection for rasterization redaction"""
        self.redact_mode = "rasterize"
        tab = self.current_tab()
        if not tab: return
        
        tab.label.set_selection_mode(True)
        tab.label.setCursor(Qt.CrossCursor)
        QMessageBox.information(self, "Select Redaction Area", "Draw a box around the area (e.g. page number) to redact on ALL pages during rasterization.")

    def apply_custom_redaction(self, tab, ui_rect):
        """Apply redaction with preview/confirm/undo flow"""
        tab.label.set_selection_mode(False)
        tab.label.setCursor(Qt.ArrowCursor)
        
        if ui_rect.width() < 5 or ui_rect.height() < 5:
            return

        pixmap = tab.label.pixmap()
        if not pixmap or pixmap.isNull():
            return

        try:
            # 1. Calculate normalized coordinates (0.0 to 1.0)
            p_width = pixmap.width()
            p_height = pixmap.height()
            
            pixmap_rect = pixmap.rect()
            label_rect = tab.label.rect()
            
            # Center offset
            offset_x = (label_rect.width() - pixmap_rect.width()) / 2
            offset_y = (label_rect.height() - pixmap_rect.height()) / 2
            
            # Get Selection coordinates relative to the Pixmap
            vis_x0 = (ui_rect.left() - offset_x)
            vis_y0 = (ui_rect.top() - offset_y)
            vis_x1 = (ui_rect.right() - offset_x)
            vis_y1 = (ui_rect.bottom() - offset_y)
            
            # Normalize coordinates (0.0 to 1.0)
            n_x0 = max(0.0, min(vis_x0 / p_width, 1.0))
            n_y0 = max(0.0, min(vis_y0 / p_height, 1.0))
            n_x1 = max(0.0, min(vis_x1 / p_width, 1.0))
            n_y1 = max(0.0, min(vis_y1 / p_height, 1.0))

            # --- BRANCH BASED ON MODE ---
            if getattr(self, "redact_mode", "standard") == "rasterize":
                # Rasterization mode - existing logic
                page = tab.doc.load_page(tab.current_page)
                vis_w_pts, vis_h_pts = page.rect.width, page.rect.height
                
                rect_w = (n_x1 - n_x0) * vis_w_pts
                rect_h = (n_y1 - n_y0) * vis_h_pts
                dist_right = vis_w_pts - (n_x1 * vis_w_pts)
                dist_bottom = vis_h_pts - (n_y1 * vis_h_pts)
                
                reply = QMessageBox.question(self, "Confirm Rasterize & Redact", 
                                           "This will convert all pages to images and redact this area.\n\nProceed?",
                                           QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    geometry = (rect_w, rect_h, dist_right, dist_bottom)
                    self.rasterize_with_redaction(tab, geometry)
                return

            # Standard Mode - Preview and Confirm
            # Store the selection for visual preview (the PDFCanvas already shows the selection box)
            # Keep the selection visible for preview
            tab.label.selection = ui_rect
            tab.label.update()
            
            # Show confirm/undo dialog
            from PySide6.QtWidgets import QDialog, QDialogButtonBox
            
            dialog = QDialog(self)
            dialog.setWindowTitle("Confirm Redaction")
            dialog_layout = QVBoxLayout(dialog)
            
            preview_label = QLabel("Preview: The selected area (shown as blue box) will be redacted with white.")
            preview_label.setWordWrap(True)
            dialog_layout.addWidget(preview_label)
            
            all_pages_label = QLabel("Apply redaction to ALL pages at this relative position?")
            all_pages_label.setStyleSheet("font-weight: bold;")
            dialog_layout.addWidget(all_pages_label)
            
            button_box = QDialogButtonBox()
            btn_confirm = button_box.addButton("Confirm (All Pages)", QDialogButtonBox.AcceptRole)
            btn_undo = button_box.addButton("Undo", QDialogButtonBox.RejectRole)
            dialog_layout.addWidget(button_box)
            
            btn_confirm.clicked.connect(dialog.accept)
            btn_undo.clicked.connect(dialog.reject)
            
            result = dialog.exec()
            
            # Clear the preview selection
            tab.label.selection = None
            tab.label.update()
            
            if result != QDialog.Accepted:
                # User clicked Undo
                return
            
            # Apply redaction to all pages using PyMuPDF's derotation matrix
            for pg_idx in range(len(tab.doc)):
                pg = tab.doc.load_page(pg_idx)
                
                # Visual dimensions (from pg.rect which accounts for rotation)
                w_vis = pg.rect.width
                h_vis = pg.rect.height
                
                # Map normalized coordinates to visual points for this page
                vx0 = n_x0 * w_vis
                vy0 = n_y0 * h_vis
                vx1 = n_x1 * w_vis
                vy1 = n_y1 * h_vis
                
                # Use derotation_matrix to transform visual coords to internal (MediaBox) coords
                # This handles all rotation cases correctly
                derot = pg.derotation_matrix
                
                # Transform corner points
                p0 = fitz.Point(vx0, vy0) * derot
                p1 = fitz.Point(vx1, vy1) * derot
                
                # Create rect from transformed points and normalize
                rect = fitz.Rect(p0, p1).normalize()

                pg.add_redact_annot(rect, fill=(1, 1, 1))
                pg.apply_redactions()
            
            tab.render()
            QMessageBox.information(self, "Success", f"Redaction applied to all {len(tab.doc)} pages.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def show_advanced_menu(self):
        tab = self.current_tab()
        if not tab: return
        
        from PySide6.QtGui import QCursor
        menu = QMenu(self)
        
        act_sanitize = menu.addAction("🔓 Sanitize & Unlock PDF")
        act_sanitize.setToolTip("Remove passwords, encryption, and restriction flags.")
        
        act_rasterize = menu.addAction("🖼️ Rasterize & Redact Bottom (Draw Box)")
        act_rasterize.setToolTip("Convert pages to images to fix orientation/font issues, then redact a selected area.")
        
        # Show menu at mouse cursor position
        action = menu.exec(QCursor.pos())
        
        if action == act_sanitize:
            self.sanitize_pdf(tab)
        elif action == act_rasterize:
            self.prepare_rasterize_redaction()

    def rasterize_with_redaction(self, tab, geometry):
        """Convert pages to images and redact using relative geometry
        geometry: (width, height, dist_from_right, dist_from_bottom)
        """
        import traceback
        import uuid
        
        rect_w, rect_h, dist_right, dist_bottom = geometry
        
        try:
            QMessageBox.information(self, "Processing", "Rasterizing and redacting... ensure coordinates are correct.")
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            src_doc = tab.doc
            new_doc = fitz.open() # New empty PDF
            
            for i, page in enumerate(src_doc):
                try:
                    # Render image
                    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                    new_page = new_doc.new_page(width=pix.width, height=pix.height)
                    new_page.insert_image(new_page.rect, stream=pix.tobytes("jpg"), keep_proportion=True)
                    
                    # Calculate redaction rect for THIS page dimensions
                    p_w, p_h = float(pix.width), float(pix.height)
                    
                    # Note: pixmap dimensions might differ from PDF point dimensions if scaled?
                    # fitz.Matrix(2.0) scales the output image by 2x.
                    # new_page.rect is set to pix.width/height, so coordinate space matches pixels.
                    # HOWEVER, the geometry passed in was from PDF point coordinates (unscaled).
                    # We must scale the redaction geometry by 2.0 to match the high-res image page.
                    
                    scale_factor = 2.0
                    r_w = rect_w * scale_factor
                    r_h = rect_h * scale_factor
                    d_r = dist_right * scale_factor
                    d_b = dist_bottom * scale_factor
                    
                    x1 = p_w - d_r
                    y1 = p_h - d_b
                    x0 = x1 - r_w
                    y0 = y1 - r_h
                    
                    redact_rect = fitz.Rect(x0, y0, x1, y1)
                    new_page.draw_rect(redact_rect, color=(1, 1, 1), fill=(1, 1, 1))
                    
                    pix = None
                except Exception as inner_e:
                    print(f"Error processing page {i+1}: {inner_e}")
                    raise inner_e
            
            # Save new PDF
            new_filename = f"rasterized_redacted_{uuid.uuid4().hex[:8]}.pdf"
            new_path = os.path.join(self.temp_dir, new_filename)
            new_doc.save(new_path)
            new_doc.close()
            
            QApplication.restoreOverrideCursor()
            self.open_pdf_file(new_path)
            QMessageBox.information(self, "Success", "Rasterization complete! output opened in new tab.")
            
        except Exception as e:
            QApplication.restoreOverrideCursor()
            error_msg = f"Rasterization failed: {e}\n{traceback.format_exc()}"
            print(error_msg)
            QMessageBox.critical(self, "Error", f"Rasterization failed: {e}")

    def sanitize_pdf(self, tab):
        """Remove security and saving as a clean copy"""
        try:
            import uuid
            new_filename = f"sanitized_{uuid.uuid4().hex[:8]}.pdf"
            new_path = os.path.join(self.temp_dir, new_filename)
            
            # Save without encryption first
            tab.doc.save(new_path, encryption=fitz.PDF_ENCRYPT_NONE)
            
            # Open source for baking
            src_doc = fitz.open(new_path)
            out_doc = fitz.open()
            
            # Iterate and bake rotation
            for page in src_doc:
                rot = page.rotation
                
                # Create new page with VISUAL dimensions (page.rect reflects rotation already)
                # If page is rotated 90, page.rect is already swapped (e.g. 842x595).
                # So we just transform the visual rect to the new page.
                new_page = out_doc.new_page(width=page.rect.width, height=page.rect.height)
                
                # Sync MediaBox to CropBox to prevent zoom-out/scaling issues naturally
                # [FIX]: Normalize the origin to (0,0) so that show_pdf_page doesn't translate (shift) the content.
                page.set_mediabox(page.cropbox)
                
                # Draw the page with its rotation baked in
                # We use 'rotate=rot' (positive) to preserve the VISUAL orientation.
                # If page is Rot 90 (Visual Landscape, Top is Right), we want result to be Landscape, Top on Right.
                # rotate=90 achieves this mapping.
                new_page.show_pdf_page(new_page.rect, src_doc, page.number, rotate=rot, clip=page.cropbox)
            
            # Save final baked PDF to a NEW path to avoid Windows file locking issues
            final_path = new_path.replace(".pdf", "_baked.pdf")
            out_doc.save(final_path)
            out_doc.close()
            
            # Close source
            src_doc.close()
            
            # Try to cleanup intermediate file (soft fail)
            try:
                os.remove(new_path)
            except:
                pass # If locked, let OS/cleanup handle it later
            
            # Open the new file
            self.open_pdf_file(final_path)
            QMessageBox.information(self, "Success", "PDF sanitized (rotation baked) and opened in new tab!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Sanitization failed: {e}")

    def add_page_numbers(self):
        tab = self.current_tab()
        if not tab: return
        
        # Default values (stored as class attributes)
        if not hasattr(self, 'pn_defaults'):
            self.pn_defaults = {
                'dist_bottom': 25,
                'dist_edge': 20,
                'font_size': 10,
                'format': 0,  # index in combo
                'position': 0  # index in combo
            }
            
        dialog = QDialog(self)
        dialog.setWindowTitle("Add/Remove Page Numbers")
        dialog.resize(400, 450)
        layout = QVBoxLayout(dialog)
        
        # Remove Button at top
        btn_remove = QPushButton("🗑️ Remove Added Page Numbers")
        btn_remove.setStyleSheet("background-color: #dc2626; color: white; padding: 8px;")
        btn_remove.clicked.connect(lambda: self.remove_page_numbers(tab, dialog))
        layout.addWidget(btn_remove)
        
        layout.addWidget(QLabel("<hr>"))
        
        layout.addWidget(QLabel("Format:"))
        fmt_combo = QComboBox()
        fmt_combo.addItems(["Page n of n", "n"])
        fmt_combo.setCurrentIndex(self.pn_defaults['format'])
        layout.addWidget(fmt_combo)
        
        layout.addWidget(QLabel("Skip Pages (No Number, No Count - e.g. 1, 3-5):"))
        skip_input = QLineEdit()
        layout.addWidget(skip_input)

        layout.addWidget(QLabel("Omit Numbers (Count continues, but hide text - e.g. 2, 6):"))
        omit_input = QLineEdit()
        layout.addWidget(omit_input)
        
        layout.addWidget(QLabel("Position:"))
        pos_combo = QComboBox()
        pos_combo.addItems(["Bottom Center", "Bottom Right", "Bottom Left", "Top Center", "Top Right"])
        pos_combo.setCurrentIndex(self.pn_defaults['position'])
        layout.addWidget(pos_combo)
        
        # Distance inputs
        dist_layout = QHBoxLayout()
        dist_layout.addWidget(QLabel("Distance from Bottom/Top (pts):"))
        dist_bottom_spin = QSpinBox()
        dist_bottom_spin.setRange(5, 200)
        dist_bottom_spin.setValue(self.pn_defaults['dist_bottom'])
        dist_layout.addWidget(dist_bottom_spin)
        layout.addLayout(dist_layout)
        
        edge_layout = QHBoxLayout()
        edge_layout.addWidget(QLabel("Distance from Edge (pts):"))
        dist_edge_spin = QSpinBox()
        dist_edge_spin.setRange(5, 200)
        dist_edge_spin.setValue(self.pn_defaults['dist_edge'])
        edge_layout.addWidget(dist_edge_spin)
        layout.addLayout(edge_layout)
        
        layout.addWidget(QLabel("Font Size:"))
        size_spin = QSpinBox()
        size_spin.setRange(6, 72)
        size_spin.setValue(self.pn_defaults['font_size'])
        layout.addWidget(size_spin)
        
        # Set as Default button
        def save_defaults():
            self.pn_defaults = {
                'dist_bottom': dist_bottom_spin.value(),
                'dist_edge': dist_edge_spin.value(),
                'font_size': size_spin.value(),
                'format': fmt_combo.currentIndex(),
                'position': pos_combo.currentIndex()
            }
            QMessageBox.information(dialog, "Saved", "Current settings saved as default!")
        
        btn_default = QPushButton("💾 Set as Default")
        btn_default.clicked.connect(save_defaults)
        layout.addWidget(btn_default)
        
        # Flatten checkbox (hides from comments panel)
        flatten_check = QCheckBox("Flatten (hide from comments panel - cannot be removed later)")
        layout.addWidget(flatten_check)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        def parse_pages(p_str):
            pages = set()
            if not p_str: return pages
            for part in p_str.split(','):
                try:
                    if '-' in part:
                        start, end = map(int, part.strip().split('-'))
                        pages.update(range(start, end + 1))
                    else:
                        pages.add(int(part.strip()))
                except: pass
            return pages

        if dialog.exec() == QDialog.Accepted:
            try:
                doc = tab.doc
                skipped = parse_pages(skip_input.text())
                omitted = parse_pages(omit_input.text())
                
                total_eligible = len(doc) - len([p for p in skipped if 1 <= p <= len(doc)])
                fmt = fmt_combo.currentText()
                font_size = size_spin.value()
                dist_bottom = dist_bottom_spin.value()
                dist_edge = dist_edge_spin.value()
                
                # Generate unique tag for this batch
                tag = f"PDFEDITOR_PN_{uuid.uuid4().hex[:8]}"
                
                current_seq_num = 1
                for i in range(len(doc)):
                    pg_index = i + 1
                    
                    if pg_index in skipped:
                        continue
                    
                    if pg_index not in omitted:
                        page = doc.load_page(i)
                        
                        if fmt == "n":
                            text = f"{current_seq_num}"
                        else:
                            text = f"Page {current_seq_num} of {total_eligible}"
                        
                        # Use visual dimensions (page.rect accounts for rotation)
                        vis_width = page.rect.width
                        vis_height = page.rect.height
                        pos_idx = pos_combo.currentIndex()
                        
                        # Calculate annotation rectangle in VISUAL coordinates
                        text_width = len(text) * (font_size * 0.6)
                        text_height = font_size * 1.5
                        
                        if pos_idx == 0:  # Bottom Center
                            vx0 = (vis_width - text_width) / 2
                            vy0 = vis_height - dist_bottom - text_height
                        elif pos_idx == 1:  # Bottom Right
                            vx0 = vis_width - dist_edge - text_width
                            vy0 = vis_height - dist_bottom - text_height
                        elif pos_idx == 2:  # Bottom Left
                            vx0 = dist_edge
                            vy0 = vis_height - dist_bottom - text_height
                        elif pos_idx == 3:  # Top Center
                            vx0 = (vis_width - text_width) / 2
                            vy0 = dist_bottom
                        else:  # Top Right
                            vx0 = vis_width - dist_edge - text_width
                            vy0 = dist_bottom
                        
                        vx1 = vx0 + text_width
                        vy1 = vy0 + text_height
                        
                        # Transform visual coords to internal coords using derotation matrix
                        derot = page.derotation_matrix
                        p0 = fitz.Point(vx0, vy0) * derot
                        p1 = fitz.Point(vx1, vy1) * derot
                        annot_rect = fitz.Rect(p0, p1).normalize()
                        
                        # Determine text rotation for the annotation based on page rotation
                        rotate_angle = page.rotation
                        
                        # Create FreeText annotation
                        annot = page.add_freetext_annot(
                            annot_rect,
                            text,
                            fontsize=font_size,
                            fontname="helv",
                            text_color=(0, 0, 0),
                            fill_color=None,
                            border_color=None,
                            align=fitz.TEXT_ALIGN_CENTER,
                            rotate=rotate_angle
                        )
                        # Tag for later removal
                        annot.set_info(title=tag)
                        annot.update()
                    
                    current_seq_num += 1
                
                # Flatten annotations if checkbox is checked
                if flatten_check.isChecked():
                    for page in doc:
                        # Get all annotations with our tag and flatten them
                        annots_to_flatten = []
                        for annot in page.annots():
                            info = annot.info
                            title = info.get("title", "")
                            if title == tag:  # Only flatten the ones we just added
                                annots_to_flatten.append(annot)
                        
                        for annot in annots_to_flatten:
                            # Render annotation to pixmap and insert as image
                            annot_rect = annot.rect
                            # Use the annot appearance directly via update
                            page.delete_annot(annot)
                            # Actually we use a simpler approach - insert text directly
                        
                        # Better approach: re-insert as text instead of annotation
                    # Re-do with insert_text for flattened version
                    current_seq_num = 1
                    for i in range(len(doc)):
                        pg_index = i + 1
                        if pg_index in skipped:
                            continue
                        if pg_index not in omitted:
                            page = doc.load_page(i)
                            if fmt == "n":
                                text = f"{current_seq_num}"
                            else:
                                text = f"Page {current_seq_num} of {total_eligible}"
                            
                            vis_width = page.rect.width
                            vis_height = page.rect.height
                            pos_idx = pos_combo.currentIndex()
                            text_width = len(text) * (font_size * 0.6)
                            text_height = font_size * 1.5
                            
                            if pos_idx == 0:
                                vx0 = (vis_width - text_width) / 2
                                vy0 = vis_height - dist_bottom - text_height
                            elif pos_idx == 1:
                                vx0 = vis_width - dist_edge - text_width
                                vy0 = vis_height - dist_bottom - text_height
                            elif pos_idx == 2:
                                vx0 = dist_edge
                                vy0 = vis_height - dist_bottom - text_height
                            elif pos_idx == 3:
                                vx0 = (vis_width - text_width) / 2
                                vy0 = dist_bottom
                            else:
                                vx0 = vis_width - dist_edge - text_width
                                vy0 = dist_bottom
                            
                            # Transform for rotation
                            derot = page.derotation_matrix
                            pt = fitz.Point(vx0, vy0 + text_height) * derot
                            
                            # Insert as regular text (not annotation)
                            page.insert_text(pt, text, fontname="helv", fontsize=font_size, color=(0, 0, 0), rotate=page.rotation)
                        current_seq_num += 1
                    
                    msg = "Page numbers added (flattened - not removable)!"
                else:
                    msg = "Page numbers added! Use 'Remove' to delete."
                
                tab.render()
                QMessageBox.information(self, "Success", msg)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
    
    def remove_page_numbers(self, tab, parent_dialog):
        """Remove only page number annotations that were added by this tool (tagged with PDFEDITOR_PN_)"""
        try:
            doc = tab.doc
            removed_count = 0
            
            for page in doc:
                annots_to_delete = []
                for annot in page.annots():
                    info = annot.info
                    title = info.get("title", "")
                    if title.startswith("PDFEDITOR_PN_"):
                        annots_to_delete.append(annot)
                
                for annot in annots_to_delete:
                    page.delete_annot(annot)
                    removed_count += 1
            
            tab.render()
            parent_dialog.accept()
            QMessageBox.information(self, "Success", f"Removed {removed_count} page number annotations!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def add_header_footer(self):
        tab = self.current_tab()
        if not tab: return
        
        # Default values (stored as class attributes)
        if not hasattr(self, 'hf_defaults'):
            self.hf_defaults = {
                'dist_top_bottom': 15,
                'dist_edge': 20,
                'font_size': 26,
                'color': 'Red',
                'type': 0,  # 0=Header, 1=Footer
                'align': 0  # 0=Center, 1=Left, 2=Right
            }
            
        dialog = QDialog(self)
        dialog.setWindowTitle("Add/Remove Header/Footer")
        dialog.resize(450, 450)
        layout = QVBoxLayout(dialog)
        
        # Remove Button at top
        btn_remove = QPushButton("🗑️ Remove Added Headers/Footers")
        btn_remove.setStyleSheet("background-color: #dc2626; color: white; padding: 8px;")
        btn_remove.clicked.connect(lambda: self.remove_header_footer(tab, dialog))
        layout.addWidget(btn_remove)
        
        layout.addWidget(QLabel("<hr>"))
        
        # Preset Button
        btn_draft = QPushButton("Load 'DRAFT' Preset")
        layout.addWidget(btn_draft)
        
        layout.addWidget(QLabel("Text:"))
        text_input = QLineEdit()
        layout.addWidget(text_input)
        
        layout.addWidget(QLabel("Type:"))
        type_combo = QComboBox()
        type_combo.addItems(["Header", "Footer"])
        type_combo.setCurrentIndex(self.hf_defaults['type'])
        layout.addWidget(type_combo)
        
        layout.addWidget(QLabel("Alignment:"))
        align_combo = QComboBox()
        align_combo.addItems(["Center", "Left", "Right"])
        align_combo.setCurrentIndex(self.hf_defaults['align'])
        layout.addWidget(align_combo)
        
        # Distance inputs
        dist_layout = QHBoxLayout()
        dist_layout.addWidget(QLabel("Distance from Top/Bottom (pts):"))
        dist_tb_spin = QSpinBox()
        dist_tb_spin.setRange(5, 200)
        dist_tb_spin.setValue(self.hf_defaults['dist_top_bottom'])
        dist_layout.addWidget(dist_tb_spin)
        layout.addLayout(dist_layout)
        
        edge_layout = QHBoxLayout()
        edge_layout.addWidget(QLabel("Distance from Edge (pts):"))
        dist_edge_spin = QSpinBox()
        dist_edge_spin.setRange(5, 200)
        dist_edge_spin.setValue(self.hf_defaults['dist_edge'])
        edge_layout.addWidget(dist_edge_spin)
        layout.addLayout(edge_layout)
        
        # Styling (Font is always Times New Roman)
        style_layout = QHBoxLayout()
        
        style_layout.addWidget(QLabel("Size:"))
        size_spin = QSpinBox()
        size_spin.setRange(8, 72)
        size_spin.setValue(self.hf_defaults['font_size'])
        style_layout.addWidget(size_spin)
        
        style_layout.addWidget(QLabel("Color:"))
        color_combo = QComboBox()
        color_combo.addItems(["Black", "Red", "Blue", "Green", "Gray"])
        color_combo.setCurrentText(self.hf_defaults['color'])
        style_layout.addWidget(color_combo)
        
        layout.addLayout(style_layout)
        
        layout.addWidget(QLabel("Font: Times New Roman (fixed)"))
        
        # Preset Logic
        def load_draft():
            text_input.setText("DRAFT")
            type_combo.setCurrentText("Header")
            align_combo.setCurrentText("Center")
            size_spin.setValue(26)
            color_combo.setCurrentText("Red")
            dist_tb_spin.setValue(15)
        
        btn_draft.clicked.connect(load_draft)
        
        # Set as Default button
        def save_defaults():
            self.hf_defaults = {
                'dist_top_bottom': dist_tb_spin.value(),
                'dist_edge': dist_edge_spin.value(),
                'font_size': size_spin.value(),
                'color': color_combo.currentText(),
                'type': type_combo.currentIndex(),
                'align': align_combo.currentIndex()
            }
            QMessageBox.information(dialog, "Saved", "Current settings saved as default!")
        
        btn_default = QPushButton("💾 Set as Default")
        btn_default.clicked.connect(save_defaults)
        layout.addWidget(btn_default)
        
        # Flatten checkbox (hides from comments panel)
        flatten_check = QCheckBox("Flatten (hide from comments panel - cannot be removed later)")
        layout.addWidget(flatten_check)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec() == QDialog.Accepted:
            text = text_input.text()
            if not text: return
            
            try:
                doc = tab.doc
                is_header = type_combo.currentText() == "Header"
                align = align_combo.currentText()
                size = size_spin.value()
                color_name = color_combo.currentText().lower()
                dist_tb = dist_tb_spin.value()
                dist_edge = dist_edge_spin.value()
                
                # Map color names to RGB tuples
                colors = {
                    "black": (0, 0, 0),
                    "red": (1, 0, 0),
                    "blue": (0, 0, 1),
                    "green": (0, 0.5, 0),
                    "gray": (0.5, 0.5, 0.5)
                }
                color = colors.get(color_name, (0, 0, 0))
                
                # Generate unique tag for this batch of annotations
                tag = f"PDFEDITOR_HF_{uuid.uuid4().hex[:8]}"
                
                for page in doc:
                    # Use visual dimensions (page.rect accounts for rotation)
                    vis_width = page.rect.width
                    vis_height = page.rect.height
                    
                    # Calculate text dimensions
                    text_width = len(text) * (size * 0.6)
                    text_height = size * 1.5
                    
                    # Calculate position in VISUAL coordinates
                    if is_header:
                        vy0 = dist_tb
                    else:
                        vy0 = vis_height - dist_tb - text_height
                    vy1 = vy0 + text_height
                    
                    if align == "Center":
                        vx0 = (vis_width - text_width) / 2
                    elif align == "Left":
                        vx0 = dist_edge
                    else:
                        vx0 = vis_width - dist_edge - text_width
                    vx1 = vx0 + text_width
                    
                    # Transform visual coords to internal coords using derotation matrix
                    derot = page.derotation_matrix
                    p0 = fitz.Point(vx0, vy0) * derot
                    p1 = fitz.Point(vx1, vy1) * derot
                    annot_rect = fitz.Rect(p0, p1).normalize()
                    
                    # Determine text rotation for the annotation based on page rotation
                    rotate_angle = page.rotation
                    
                    # Create FreeText annotation - Always use Times Roman font
                    annot = page.add_freetext_annot(
                        annot_rect,
                        text,
                        fontsize=size,
                        fontname="tiro",  # Times Roman
                        text_color=color,
                        fill_color=None,
                        border_color=None,
                        align=fitz.TEXT_ALIGN_CENTER if align == "Center" else (fitz.TEXT_ALIGN_LEFT if align == "Left" else fitz.TEXT_ALIGN_RIGHT),
                        rotate=rotate_angle
                    )
                    # Tag the annotation for later removal
                    annot.set_info(title=tag)
                    annot.update()
                
                # Flatten annotations if checkbox is checked
                if flatten_check.isChecked():
                    # Delete annotations and re-insert as text
                    for page in doc:
                        annots_to_delete = []
                        for annot in page.annots():
                            info = annot.info
                            title = info.get("title", "")
                            if title == tag:
                                annots_to_delete.append(annot)
                        for annot in annots_to_delete:
                            page.delete_annot(annot)
                    
                    # Re-insert as flattened text
                    for page in doc:
                        vis_width = page.rect.width
                        vis_height = page.rect.height
                        text_width = len(text) * (size * 0.6)
                        text_height = size * 1.5
                        
                        if is_header:
                            vy0 = dist_tb
                        else:
                            vy0 = vis_height - dist_tb - text_height
                        
                        if align == "Center":
                            vx0 = (vis_width - text_width) / 2
                        elif align == "Left":
                            vx0 = dist_edge
                        else:
                            vx0 = vis_width - dist_edge - text_width
                        
                        # Transform for rotation
                        derot = page.derotation_matrix
                        pt = fitz.Point(vx0, vy0 + text_height) * derot
                        
                        # Insert as regular text (not annotation)
                        page.insert_text(pt, text, fontname="tiro", fontsize=size, color=color, rotate=page.rotation)
                    
                    msg = "Header/Footer added (flattened - not removable)!"
                else:
                    msg = "Header/Footer added! Use 'Remove' to delete."
                
                tab.render()
                QMessageBox.information(self, "Success", msg)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
    
    def remove_header_footer(self, tab, parent_dialog):
        """Remove only FreeText annotations that were added by this tool (tagged with PDFEDITOR_HF_)"""
        try:
            doc = tab.doc
            removed_count = 0
            
            for page in doc:
                # Get all annotations on this page
                annots_to_delete = []
                for annot in page.annots():
                    # Check if this is a tagged header/footer annotation
                    info = annot.info
                    title = info.get("title", "")
                    if title.startswith("PDFEDITOR_HF_"):
                        annots_to_delete.append(annot)
                
                # Delete the tagged annotations
                for annot in annots_to_delete:
                    page.delete_annot(annot)
                    removed_count += 1
            
            tab.render()
            parent_dialog.accept()
            QMessageBox.information(self, "Success", f"Removed {removed_count} header/footer annotations!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))



    def open_pdf_file(self, path):
        """Helper to open a PDF file given a path"""
        try:
            doc = fitz.open(path)
            # Check if likely a temp file
            is_temp = ".temp_pdfs" in path
            tab = PDFTab(doc, path, is_temp=is_temp, temp_path=path if is_temp else None)
            
            from PySide6.QtWidgets import QDockWidget
            dock = QDockWidget(os.path.basename(path), self)
            dock.setWidget(tab)
            dock.setAllowedAreas(Qt.AllDockWidgetAreas)
            dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable | QDockWidget.DockWidgetClosable)
            
            tab.parent_dock = dock
            dock.setContextMenuPolicy(Qt.CustomContextMenu)
            dock.customContextMenuRequested.connect(lambda pos, d=dock: self.dock_context_menu(pos, d))
            dock.visibilityChanged.connect(self.on_dock_visibility_changed)
            
            # Connect interactive selection signal
            tab.label.selection_confirmed.connect(lambda rect: self.apply_custom_redaction(tab, rect))
            
            self.dock_manager.addDockWidget(Qt.RightDockWidgetArea, dock)
            if self.docks:
                self.dock_manager.tabifyDockWidget(self.docks[-1], dock)
            self.docks.append(dock)
            dock.show()
            if len(self.docks) == 1: self._last_active_tab = tab
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open file: {e}")

# ============================================================================
# OCR TRAINER MODULE - Enhanced with Multi-PDF and Hierarchical Boxes
# ============================================================================

class OCRBox:
    """Represents a box with type (label/anchor/value) and optional parent"""
    def __init__(self, rect, name, box_type='label', parent=None):
        self.rect = rect  # QRectF
        self.name = name
        self.box_type = box_type  # 'label', 'anchor', 'value'
        self.parent = parent  # Parent OCRBox for anchor/value
        self.children = []  # Child boxes (anchors/values)
        self.id = None  # Database ID, set after saving
        self.anchor_text = ""  # OCR-captured text inside anchor box
    
    def add_child(self, child):
        child.parent = self
        self.children.append(child)

class OCRCanvasWidget(QWidget):
    """Enhanced canvas for drawing hierarchical OCR boxes"""
    box_created = Signal(object)  # Emits OCRBox when created
    box_selected = Signal(object)  # Emits selected OCRBox
    
    # Colors for different box types
    COLORS = {
        'label': QColor(66, 133, 244),      # Blue - Parent/Label boxes
        'anchor': QColor(52, 168, 83),      # Green - Anchor boxes
        'value': QColor(251, 188, 5),       # Orange/Yellow - Value boxes
        'selected': QColor(234, 67, 53),    # Red - Selected box
        'drawing': QColor(100, 100, 100)    # Gray - Currently drawing
    }
    
    def __init__(self):
        super().__init__()
        self.pixmap = None
        self.boxes = []  # List of OCRBox
        self.start_point = None
        self.current_rect = None
        self.scale_factor = 1.0
        self.setMinimumSize(400, 400)
        self.selected_box = None
        self.current_mode = 'label'  # 'label', 'anchor', 'value'
        self.active_parent_box = None  # For anchor/value drawing
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        
        # Handle size for resize operations
        self.handle_size = 8
        self.drag_mode = None  # 'move', 'resize_*', None
        self.drag_start = None
        self.resize_handle = None
    
    def set_image(self, pixmap, scale_factor=1.0):
        self.pixmap = pixmap
        self.boxes = []
        self.scale_factor = scale_factor
        self.selected_box = None
        self.active_parent_box = None
        if pixmap:
            self.setFixedSize(pixmap.size())
        self.update()
    
    def set_boxes(self, boxes):
        """Set boxes from loaded template"""
        self.boxes = boxes
        self.update()
    
    def set_mode(self, mode):
        """Set drawing mode: 'label', 'anchor', or 'value'"""
        self.current_mode = mode
        if mode == 'label':
            self.active_parent_box = None
        self.update()
    
    def set_active_parent(self, parent_box):
        """Set the active parent box for anchor/value drawing"""
        self.active_parent_box = parent_box
        self.selected_box = parent_box
        self.update()
    
    def get_handle_rects(self, box):
        """Get resize handle rectangles for a box"""
        r = box.rect.toRect()
        s = self.handle_size
        hs = s // 2
        return {
            'tl': QRect(r.left() - hs, r.top() - hs, s, s),
            'tr': QRect(r.right() - hs, r.top() - hs, s, s),
            'bl': QRect(r.left() - hs, r.bottom() - hs, s, s),
            'br': QRect(r.right() - hs, r.bottom() - hs, s, s),
        }
    
    def get_handle_at(self, pos, box):
        """Check if position is on a resize handle"""
        for name, rect in self.get_handle_rects(box).items():
            if rect.contains(pos):
                return name
        return None
    
    def paintEvent(self, event):
        if not self.pixmap:
            return
        
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.drawPixmap(0, 0, self.pixmap)
        
        # Draw all boxes
        for box in self.boxes:
            self._draw_box(painter, box)
            # Draw children
            for child in box.children:
                self._draw_box(painter, child)
        
        # Draw current drawing rect
        if self.current_rect:
            pen = QPen(self.COLORS['drawing'], 2, Qt.DashLine)
            painter.setPen(pen)
            painter.setBrush(QBrush(QColor(100, 100, 100, 50)))
            painter.drawRect(self.current_rect.toRect())
        
        # Draw mode indicator
        self._draw_mode_indicator(painter)
    
    def _draw_box(self, painter, box):
        """Draw a single box with appropriate styling"""
        is_selected = box == self.selected_box
        color = self.COLORS['selected'] if is_selected else self.COLORS.get(box.box_type, self.COLORS['label'])
        
        # Box outline
        pen_width = 3 if is_selected else 2
        pen = QPen(color, pen_width)
        painter.setPen(pen)
        
        # Semi-transparent fill
        fill_color = QColor(color)
        fill_color.setAlpha(40)
        painter.setBrush(QBrush(fill_color))
        
        rect = box.rect.toRect()
        painter.drawRect(rect)
        
        # Draw label with background
        label = f"[{box.box_type[0].upper()}] {box.name}"
        font = painter.font()
        font.setBold(True)
        font.setPointSize(9)
        painter.setFont(font)
        
        text_rect = painter.fontMetrics().boundingRect(label)
        label_bg = QRect(rect.left(), rect.top() - text_rect.height() - 4, 
                        text_rect.width() + 8, text_rect.height() + 4)
        
        # Label background
        painter.fillRect(label_bg, color)
        painter.setPen(QPen(Qt.white))
        painter.drawText(label_bg.adjusted(4, 2, -4, -2), Qt.AlignLeft | Qt.AlignVCenter, label)
        
        # Draw resize handles if selected
        if is_selected:
            painter.setBrush(QBrush(Qt.white))
            painter.setPen(QPen(color, 1))
            for handle_rect in self.get_handle_rects(box).values():
                painter.drawRect(handle_rect)
    
    def _draw_mode_indicator(self, painter):
        """Draw current mode indicator in corner"""
        mode_labels = {
            'label': '📦 Label Mode',
            'anchor': '🎯 Anchor Mode',
            'value': '📝 Value Mode'
        }
        label = mode_labels.get(self.current_mode, '')
        
        font = painter.font()
        font.setBold(True)
        font.setPointSize(10)
        painter.setFont(font)
        
        color = self.COLORS.get(self.current_mode, self.COLORS['label'])
        text_rect = painter.fontMetrics().boundingRect(label)
        bg_rect = QRect(10, 10, text_rect.width() + 16, text_rect.height() + 8)
        
        painter.fillRect(bg_rect, QColor(0, 0, 0, 180))
        painter.setPen(QPen(color))
        painter.drawText(bg_rect, Qt.AlignCenter, label)
    
    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton or not self.pixmap:
            return
        
        pos = event.position().toPoint()
        
        # Check if clicking on selected box's resize handle
        if self.selected_box:
            handle = self.get_handle_at(pos, self.selected_box)
            if handle:
                self.drag_mode = f'resize_{handle}'
                self.drag_start = pos
                self.resize_handle = handle
                return
        
        # In anchor/value mode, allow drawing inside the active parent box
        if self.current_mode in ('anchor', 'value') and self.active_parent_box:
            # Check if click is inside the active parent box - allow drawing
            if self.active_parent_box.rect.contains(QPointF(pos)):
                # Start drawing new sub-box
                self.start_point = event.position()
                self.current_rect = QRectF(self.start_point, self.start_point)
                self.update()
                return
        
        # Check if clicking on any existing box (for selection)
        clicked_box = None
        for box in reversed(self.boxes):  # Check topmost first
            # First check children (they're on top)
            for child in box.children:
                if child.rect.contains(QPointF(pos)):
                    clicked_box = child
                    break
            if clicked_box:
                break
            if box.rect.contains(QPointF(pos)):
                clicked_box = box
                break
        
        if clicked_box:
            if clicked_box == self.selected_box:
                # Start moving
                self.drag_mode = 'move'
                self.drag_start = pos
            else:
                # Select the box
                self.selected_box = clicked_box
                # If it's a label box, make it the active parent
                if clicked_box.box_type == 'label':
                    self.active_parent_box = clicked_box
                self.box_selected.emit(clicked_box)
                self.update()
        else:
            # Start drawing new box
            self.selected_box = None
            self.start_point = event.position()
            self.current_rect = QRectF(self.start_point, self.start_point)
        
        self.update()
    
    def mouseMoveEvent(self, event):
        pos = event.position().toPoint()
        
        # Update cursor based on context
        if self.selected_box and not self.drag_mode:
            handle = self.get_handle_at(pos, self.selected_box)
            if handle:
                if handle in ['tl', 'br']:
                    self.setCursor(Qt.SizeFDiagCursor)
                else:
                    self.setCursor(Qt.SizeBDiagCursor)
            elif self.selected_box.rect.contains(QPointF(pos)):
                self.setCursor(Qt.SizeAllCursor)
            else:
                self.setCursor(Qt.CrossCursor)
        else:
            self.setCursor(Qt.CrossCursor)
        
        # Handle dragging
        if self.drag_mode and self.drag_start:
            dx = pos.x() - self.drag_start.x()
            dy = pos.y() - self.drag_start.y()
            
            if self.drag_mode == 'move':
                self.selected_box.rect.translate(dx, dy)
            elif self.drag_mode.startswith('resize_'):
                r = self.selected_box.rect
                handle = self.drag_mode.split('_')[1]
                if 'l' in handle:
                    r.setLeft(r.left() + dx)
                if 'r' in handle:
                    r.setRight(r.right() + dx)
                if 't' in handle:
                    r.setTop(r.top() + dy)
                if 'b' in handle:
                    r.setBottom(r.bottom() + dy)
                self.selected_box.rect = r.normalized()
            
            self.drag_start = pos
            self.update()
        
        # Drawing new box
        elif self.start_point:
            self.current_rect = QRectF(self.start_point, event.position()).normalized()
            self.update()
    
    def mouseReleaseEvent(self, event):
        if self.drag_mode:
            self.drag_mode = None
            self.drag_start = None
            self.resize_handle = None
            self.update()
            return
        
        if self.current_rect and self.current_rect.width() > 10 and self.current_rect.height() > 10:
            # Create new box based on mode
            name, ok = QInputDialog.getText(self, "Box Name", 
                f"Enter name for {self.current_mode} box:")
            if ok and name:
                new_box = OCRBox(self.current_rect, name, self.current_mode)
                
                if self.current_mode == 'label':
                    self.boxes.append(new_box)
                    self.active_parent_box = new_box
                elif self.active_parent_box:
                    self.active_parent_box.add_child(new_box)
                else:
                    QMessageBox.warning(self, "No Parent", 
                        "Please select a Label box first before adding anchor/value boxes.")
                
                self.selected_box = new_box
                self.box_created.emit(new_box)
        
        self.current_rect = None
        self.start_point = None
        self.update()
    
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete and self.selected_box:
            self.delete_selected_box()
    
    def delete_selected_box(self):
        """Delete the currently selected box"""
        if not self.selected_box:
            return
        
        box = self.selected_box
        
        # Remove from parent's children or main list
        if box.parent:
            box.parent.children.remove(box)
        elif box in self.boxes:
            self.boxes.remove(box)
        
        self.selected_box = None
        self.update()
    
    def clear_boxes(self):
        """Clear all boxes"""
        self.boxes = []
        self.selected_box = None
        self.active_parent_box = None
        self.update()


class OCRTrainerModule(QWidget):
    """Enhanced OCR Trainer with multi-PDF support and hierarchical boxes"""
    
    def __init__(self):
        super().__init__()
        # PDF management
        self.loaded_pdfs = []  # List of (filename, fitz.Document)
        self.current_pdf_index = -1
        self.current_page_index = 0
        
        # Page dimensions and rotation cache
        self.page_dimensions = {}  # (pdf_idx, page_idx) -> (width, height)
        self.page_rotations = {}   # (pdf_idx, page_idx) -> rotation (0, 90, 180, 270)
        
        # Zoom scale for display
        self.zoom_scale = 1.0
        
        # Box data per page
        self.page_boxes = {}  # (pdf_idx, page_idx) -> [OCRBox]
        
        # Extraction results
        self.extraction_results = []
        
        self.setup_ui()
        self.load_template_list()
    
    def setup_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # =================== LEFT PANEL ===================
        left_panel = QVBoxLayout()
        left_panel.setSpacing(8)
        
        # Title
        title = QLabel("🔍 OCR Template Builder")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #4285F4;")
        left_panel.addWidget(title)
        
        # --- PDF Management Section ---
        pdf_section = QLabel("📁 PDF Files")
        pdf_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(pdf_section)
        
        btn_layout = QHBoxLayout()
        self.btn_add_pdf = QPushButton("➕ Add")
        self.btn_add_pdf.clicked.connect(self.add_pdfs)
        self.btn_remove_pdf = QPushButton("➖ Remove")
        self.btn_remove_pdf.clicked.connect(self.remove_pdf)
        btn_layout.addWidget(self.btn_add_pdf)
        btn_layout.addWidget(self.btn_remove_pdf)
        left_panel.addLayout(btn_layout)
        
        self.pdf_list = QListWidget()
        self.pdf_list.setMaximumHeight(120)
        self.pdf_list.currentRowChanged.connect(self.on_pdf_selected)
        left_panel.addWidget(self.pdf_list)
        
        # --- Page Navigation ---
        nav_section = QLabel("📄 Page Navigation")
        nav_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(nav_section)
        
        nav_layout = QHBoxLayout()
        self.btn_prev_page = QPushButton("◀")
        self.btn_prev_page.setFixedWidth(40)
        self.btn_prev_page.clicked.connect(lambda: self.navigate_page(-1))
        self.lbl_page = QLabel("Page 0/0")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        self.btn_next_page = QPushButton("▶")
        self.btn_next_page.setFixedWidth(40)
        self.btn_next_page.clicked.connect(lambda: self.navigate_page(1))
        nav_layout.addWidget(self.btn_prev_page)
        nav_layout.addWidget(self.lbl_page, stretch=1)
        nav_layout.addWidget(self.btn_next_page)
        left_panel.addLayout(nav_layout)
        
        # --- Drawing Mode ---
        mode_section = QLabel("🎨 Drawing Mode")
        mode_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(mode_section)
        
        mode_layout = QHBoxLayout()
        self.btn_mode_label = QPushButton("📦 Label")
        self.btn_mode_label.setCheckable(True)
        self.btn_mode_label.setChecked(True)
        self.btn_mode_label.clicked.connect(lambda: self.set_mode('label'))
        
        self.btn_mode_anchor = QPushButton("🎯 Anchor")
        self.btn_mode_anchor.setCheckable(True)
        self.btn_mode_anchor.clicked.connect(lambda: self.set_mode('anchor'))
        
        self.btn_mode_value = QPushButton("📝 Value")
        self.btn_mode_value.setCheckable(True)
        self.btn_mode_value.clicked.connect(lambda: self.set_mode('value'))
        
        mode_layout.addWidget(self.btn_mode_label)
        mode_layout.addWidget(self.btn_mode_anchor)
        mode_layout.addWidget(self.btn_mode_value)
        left_panel.addLayout(mode_layout)
        
        # Style mode buttons
        for btn in [self.btn_mode_label, self.btn_mode_anchor, self.btn_mode_value]:
            btn.setStyleSheet("""
                QPushButton { padding: 8px; border-radius: 4px; }
                QPushButton:checked { background: #4285F4; color: white; }
            """)
        
        # --- Template Management ---
        template_section = QLabel("💾 Template")
        template_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(template_section)
        
        self.template_name_input = QLineEdit()
        self.template_name_input.setPlaceholderText("Enter template name...")
        left_panel.addWidget(self.template_name_input)
        
        save_layout = QHBoxLayout()
        self.btn_save_template = QPushButton("💾 Save Template")
        self.btn_save_template.clicked.connect(self.save_template)
        self.btn_save_template.setStyleSheet("background: #34A853; color: white; padding: 8px;")
        save_layout.addWidget(self.btn_save_template)
        
        self.btn_test_extract = QPushButton("🧪 Test Extract")
        self.btn_test_extract.clicked.connect(self.test_extract_current)
        self.btn_test_extract.setStyleSheet("background: #FF9800; color: white; padding: 8px;")
        save_layout.addWidget(self.btn_test_extract)
        left_panel.addLayout(save_layout)
        
        # --- Load Template ---
        load_section = QLabel("📥 Load Template")
        load_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(load_section)
        
        self.template_combo = QComboBox()
        left_panel.addWidget(self.template_combo)
        
        # Load Template for Editing button
        self.btn_load_template = QPushButton("📂 Load Template for Edit")
        self.btn_load_template.clicked.connect(self.load_template_for_editing)
        self.btn_load_template.setStyleSheet("background: #9C27B0; color: white; padding: 8px;")
        left_panel.addWidget(self.btn_load_template)
        
        extract_layout = QHBoxLayout()
        self.btn_run_extraction = QPushButton("▶️ Run Extraction")
        self.btn_run_extraction.clicked.connect(self.run_extraction)
        self.btn_run_extraction.setStyleSheet("background: #4285F4; color: white; padding: 8px;")
        extract_layout.addWidget(self.btn_run_extraction)
        left_panel.addLayout(extract_layout)
        
        # --- Export Options ---
        export_layout = QHBoxLayout()
        self.btn_export_excel = QPushButton("📊 Export Excel")
        self.btn_export_excel.clicked.connect(self.export_excel)
        self.btn_export_backup = QPushButton("📄 Export Backup PDF")
        self.btn_export_backup.clicked.connect(self.export_backup_pdf)
        export_layout.addWidget(self.btn_export_excel)
        export_layout.addWidget(self.btn_export_backup)
        left_panel.addLayout(export_layout)
        
        left_panel.addStretch()
        
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        left_widget.setFixedWidth(300)
        main_layout.addWidget(left_widget)
        
        # =================== CENTER PANEL (Canvas) ===================
        center_layout = QVBoxLayout()
        
        # Zoom toolbar
        self.zoom_scale = 1.0  # Track current zoom level
        zoom_toolbar = QHBoxLayout()
        
        self.btn_zoom_fit = QPushButton("⊡ Fit")
        self.btn_zoom_fit.clicked.connect(self.zoom_fit)
        self.btn_zoom_fit.setStyleSheet("padding: 5px 10px;")
        self.btn_zoom_fit.setToolTip("Fit page to window")
        
        self.btn_zoom_out = QPushButton("➖ Zoom Out")
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        self.btn_zoom_out.setStyleSheet("padding: 5px 10px;")
        
        self.zoom_label = QLabel("100%")
        self.zoom_label.setStyleSheet("padding: 5px 15px; font-weight: bold;")
        
        self.btn_zoom_in = QPushButton("➕ Zoom In")
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        self.btn_zoom_in.setStyleSheet("padding: 5px 10px;")
        
        self.btn_zoom_100 = QPushButton("100%")
        self.btn_zoom_100.clicked.connect(self.zoom_reset)
        self.btn_zoom_100.setStyleSheet("padding: 5px 10px;")
        self.btn_zoom_100.setToolTip("Reset to 100%")
        
        zoom_toolbar.addWidget(self.btn_zoom_fit)
        zoom_toolbar.addWidget(self.btn_zoom_out)
        zoom_toolbar.addWidget(self.zoom_label)
        zoom_toolbar.addWidget(self.btn_zoom_in)
        zoom_toolbar.addWidget(self.btn_zoom_100)
        zoom_toolbar.addStretch()
        center_layout.addLayout(zoom_toolbar)
        
        # Canvas in scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(False)  # Changed for zoom support
        self.scroll_area.setStyleSheet("background: #2d2d2d;")
        
        self.canvas = OCRCanvasWidget()
        self.canvas.box_created.connect(self.on_box_created)
        self.canvas.box_selected.connect(self.on_box_selected)
        self.scroll_area.setWidget(self.canvas)
        
        center_layout.addWidget(self.scroll_area)
        main_layout.addLayout(center_layout, stretch=2)
        
        # =================== RIGHT PANEL (Box List & Results) ===================
        right_panel = QVBoxLayout()
        right_panel.setSpacing(8)
        
        # --- Box List ---
        box_section = QLabel("📋 Boxes on This Page")
        box_section.setStyleSheet("font-weight: bold;")
        right_panel.addWidget(box_section)
        
        self.box_tree = QListWidget()
        self.box_tree.setMaximumHeight(200)
        self.box_tree.itemClicked.connect(self.on_box_list_clicked)
        right_panel.addWidget(self.box_tree)
        
        box_btn_layout = QHBoxLayout()
        self.btn_delete_box = QPushButton("🗑️ Delete Box")
        self.btn_delete_box.clicked.connect(self.delete_selected_box)
        self.btn_clear_boxes = QPushButton("🧹 Clear All")
        self.btn_clear_boxes.clicked.connect(self.clear_all_boxes)
        box_btn_layout.addWidget(self.btn_delete_box)
        box_btn_layout.addWidget(self.btn_clear_boxes)
        right_panel.addLayout(box_btn_layout)
        
        # --- Extraction Results ---
        results_section = QLabel("📊 Extraction Results")
        results_section.setStyleSheet("font-weight: bold; margin-top: 15px;")
        right_panel.addWidget(results_section)
        
        self.result_table = QTableWidget(0, 3)
        self.result_table.setHorizontalHeaderLabels(["Label", "Anchor", "Value"])
        self.result_table.horizontalHeader().setStretchLastSection(True)
        right_panel.addWidget(self.result_table)
        
        right_panel.addStretch()
        
        right_widget = QWidget()
        right_widget.setLayout(right_panel)
        right_widget.setFixedWidth(350)
        main_layout.addWidget(right_widget)
    
    def add_pdfs(self):
        """Add multiple PDF files"""
        paths, _ = QFileDialog.getOpenFileNames(self, "Select PDFs", "", "PDF Files (*.pdf)")
        for path in paths:
            try:
                doc = fitz.open(path)
                filename = os.path.basename(path)
                self.loaded_pdfs.append((filename, doc, path))
                self.pdf_list.addItem(f"📄 {filename} ({len(doc)} pages)")
                
                # Cache page dimensions
                pdf_idx = len(self.loaded_pdfs) - 1
                for page_idx in range(len(doc)):
                    page = doc.load_page(page_idx)
                    self.page_dimensions[(pdf_idx, page_idx)] = (page.rect.width, page.rect.height)
                
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to open {path}: {e}")
        
        if self.loaded_pdfs and self.current_pdf_index < 0:
            self.current_pdf_index = 0
            self.current_page_index = 0
            self.pdf_list.setCurrentRow(0)
            self.render_current_page()
    
    def remove_pdf(self):
        """Remove selected PDF"""
        row = self.pdf_list.currentRow()
        if row >= 0:
            filename, doc, path = self.loaded_pdfs[row]
            doc.close()
            del self.loaded_pdfs[row]
            self.pdf_list.takeItem(row)
            
            # Clean up boxes for this PDF
            keys_to_remove = [k for k in self.page_boxes if k[0] == row]
            for k in keys_to_remove:
                del self.page_boxes[k]
            
            if not self.loaded_pdfs:
                self.current_pdf_index = -1
                self.canvas.set_image(None)
            elif row <= self.current_pdf_index:
                self.current_pdf_index = max(0, self.current_pdf_index - 1)
                self.render_current_page()
    
    def on_pdf_selected(self, row):
        """Handle PDF selection from list"""
        if row >= 0 and row != self.current_pdf_index:
            # Save current page boxes
            self.save_current_page_boxes()
            
            self.current_pdf_index = row
            self.current_page_index = 0
            self.render_current_page()
    
    def navigate_page(self, delta):
        """Navigate between pages"""
        if self.current_pdf_index < 0:
            return
        
        # Save current page boxes
        self.save_current_page_boxes()
        
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        new_page = self.current_page_index + delta
        
        if 0 <= new_page < len(doc):
            self.current_page_index = new_page
            self.render_current_page()
    
    def save_current_page_boxes(self):
        """Save boxes from canvas to page_boxes dict in PDF coordinates (scale=1.0)
        
        Boxes on canvas are drawn at current zoom scale. We convert them back to
        PDF coords so they can be correctly rescaled when zoom changes.
        """
        if self.current_pdf_index >= 0:
            key = (self.current_pdf_index, self.current_page_index)
            # Deep copy boxes and convert to PDF coordinates
            pdf_boxes = []
            for box in self.canvas.boxes:
                pdf_box = self._scale_box_to_pdf_coords(box)
                pdf_boxes.append(pdf_box)
            self.page_boxes[key] = pdf_boxes
    
    def _scale_box_to_pdf_coords(self, box):
        """Convert box from canvas coords (scaled) to PDF coords (scale=1.0)"""
        scale = self.zoom_scale
        pdf_rect = QRectF(
            box.rect.x() / scale,
            box.rect.y() / scale,
            box.rect.width() / scale,
            box.rect.height() / scale
        )
        pdf_box = OCRBox(pdf_rect, box.name, box.box_type, box.parent)
        pdf_box.id = box.id
        # Recursively convert children
        for child in box.children:
            pdf_child = self._scale_box_to_pdf_coords(child)
            pdf_child.parent = pdf_box
            pdf_box.children.append(pdf_child)
        return pdf_box
    
    def _scale_box_to_canvas_coords(self, box):
        """Convert box from PDF coords (scale=1.0) to canvas coords (scaled)"""
        scale = self.zoom_scale
        canvas_rect = QRectF(
            box.rect.x() * scale,
            box.rect.y() * scale,
            box.rect.width() * scale,
            box.rect.height() * scale
        )
        canvas_box = OCRBox(canvas_rect, box.name, box.box_type, box.parent)
        canvas_box.id = box.id
        # Recursively convert children
        for child in box.children:
            canvas_child = self._scale_box_to_canvas_coords(child)
            canvas_child.parent = canvas_box
            canvas_box.children.append(canvas_child)
        return canvas_box
    
    def render_current_page(self):
        """Render current page and load its boxes"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            return
        
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        page = doc.load_page(self.current_page_index)
        
        # Update page label
        self.lbl_page.setText(f"Page {self.current_page_index + 1}/{len(doc)}")
        
        # Track page rotation and dimensions
        key = (self.current_pdf_index, self.current_page_index)
        rotation = page.rotation  # 0, 90, 180, or 270
        self.page_rotations[key] = rotation
        
        # Store the page's original (unrotated) dimensions
        # page.rect gives the dimensions in the visual orientation (already rotated)
        self.page_dimensions[key] = (page.rect.width, page.rect.height)
        
        # Render at zoom scale
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom_scale, self.zoom_scale))
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(img)
        
        self.canvas.set_image(pixmap, scale_factor=self.zoom_scale)
        
        # Load boxes for this page (scale from PDF coords to canvas coords)
        if key in self.page_boxes:
            scaled_boxes = []
            for box in self.page_boxes[key]:
                scaled_box = self._scale_box_to_canvas_coords(box)
                scaled_boxes.append(scaled_box)
            self.canvas.set_boxes(scaled_boxes)
        else:
            self.canvas.set_boxes([])
        
        self.update_box_list()
        self.update_zoom_label()
    
    def update_zoom_label(self):
        """Update the zoom percentage label"""
        self.zoom_label.setText(f"{int(self.zoom_scale * 100)}%")
    
    def zoom_fit(self):
        """Fit the page to the scroll area"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            return
        
        # Get scroll area dimensions
        scroll_width = self.scroll_area.viewport().width() - 20
        scroll_height = self.scroll_area.viewport().height() - 20
        
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        page = doc.load_page(self.current_page_index)
        page_width = page.rect.width
        page_height = page.rect.height
        
        # Calculate scale to fit
        scale_x = scroll_width / page_width
        scale_y = scroll_height / page_height
        self.zoom_scale = min(scale_x, scale_y)
        
        self.apply_zoom()
    
    def zoom_in(self):
        """Zoom in by 25%"""
        self.zoom_scale = min(self.zoom_scale + 0.25, 5.0)
        self.apply_zoom()
    
    def zoom_out(self):
        """Zoom out by 25%"""
        self.zoom_scale = max(self.zoom_scale - 0.25, 0.25)
        self.apply_zoom()
    
    def zoom_reset(self):
        """Reset zoom to 100%"""
        self.zoom_scale = 1.0
        self.apply_zoom()
    
    def apply_zoom(self):
        """Apply the current zoom level to the canvas"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            return
        
        # Save current boxes to page_boxes (in PDF coords) before zoom changes
        self.save_current_page_boxes()
        
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        page = doc.load_page(self.current_page_index)
        
        # Render at zoom scale
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom_scale, self.zoom_scale))
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(img)
        
        self.canvas.set_image(pixmap, scale_factor=self.zoom_scale)
        self.update_zoom_label()
        
        # Reload boxes at new zoom scale
        key = (self.current_pdf_index, self.current_page_index)
        if key in self.page_boxes:
            scaled_boxes = []
            for box in self.page_boxes[key]:
                scaled_box = self._scale_box_to_canvas_coords(box)
                scaled_boxes.append(scaled_box)
            self.canvas.set_boxes(scaled_boxes)
        else:
            self.canvas.set_boxes([])
        self.canvas.update()
    
    def test_extract_current(self):
        """Test extraction on current page's boxes before saving template"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            QMessageBox.warning(self, "No PDF", "Please load a PDF first.")
            return
        
        self.save_current_page_boxes()
        
        key = (self.current_pdf_index, self.current_page_index)
        boxes = self.page_boxes.get(key, [])
        
        if not boxes:
            QMessageBox.warning(self, "No Boxes", "No boxes drawn on this page.")
            return
        
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        page = doc.load_page(self.current_page_index)
        
        # Get page rotation and dimensions
        rotation = self.page_rotations.get(key, 0)
        page_width, page_height = self.page_dimensions.get(key, (page.rect.width, page.rect.height))
        
        results = []
        
        for box in boxes:
            if box.box_type == 'label':
                # Find anchor and value children
                anchors = [b for b in box.children if b.box_type == 'anchor']
                values_boxes = [b for b in box.children if b.box_type == 'value']
                
                anchor_text = ""
                value_text = ""
                
                # Extract anchor text
                for anchor in anchors:
                    vis_x = anchor.rect.x() / self.zoom_scale
                    vis_y = anchor.rect.y() / self.zoom_scale
                    vis_w = anchor.rect.width() / self.zoom_scale
                    vis_h = anchor.rect.height() / self.zoom_scale
                    
                    # Transform to PDF coordinates
                    pdf_x, pdf_y, pdf_w, pdf_h = self.transform_visual_to_pdf_coords(
                        vis_x, vis_y, vis_w, vis_h, page_width, page_height, rotation
                    )
                    
                    rect = fitz.Rect(pdf_x, pdf_y, pdf_x + pdf_w, pdf_y + pdf_h)
                    text = page.get_text("text", clip=rect).strip()
                    if text:
                        anchor_text += text + " "
                
                # Extract value text
                for val in values_boxes:
                    vis_x = val.rect.x() / self.zoom_scale
                    vis_y = val.rect.y() / self.zoom_scale
                    vis_w = val.rect.width() / self.zoom_scale
                    vis_h = val.rect.height() / self.zoom_scale
                    
                    # Transform to PDF coordinates
                    pdf_x, pdf_y, pdf_w, pdf_h = self.transform_visual_to_pdf_coords(
                        vis_x, vis_y, vis_w, vis_h, page_width, page_height, rotation
                    )
                    
                    rect = fitz.Rect(pdf_x, pdf_y, pdf_x + pdf_w, pdf_y + pdf_h)
                    text = page.get_text("text", clip=rect).strip()
                    if text:
                        value_text += text + " "
                
                if anchor_text or value_text:
                    results.append({
                        'label': box.name,
                        'anchor': anchor_text.strip(),
                        'value': value_text.strip()
                    })
        
        # Show results in a message box and print to terminal
        if results:
            result_text = f"TEST EXTRACTION RESULTS (Page rotation: {rotation}°):\n\n"
            print(f"\n{'='*60}")
            print(f"[TEST EXTRACT] Page rotation: {rotation}°")
            print(f"{'='*60}")
            for r in results:
                result_text += f"📦 Label: {r['label']}\n"
                result_text += f"   🎯 Anchor: {r['anchor']}\n"
                result_text += f"   📝 Value: {r['value']}\n\n"
                print(f"Label: {r['label']}")
                print(f"  Anchor: {r['anchor']}")
                print(f"  Value: {r['value']}")
            print(f"{'='*60}\n")
            QMessageBox.information(self, "Test Extraction", result_text)
        else:
            QMessageBox.warning(self, "No Results", "No text extracted. Make sure anchor and value boxes are drawn inside label boxes.")
    
    def set_mode(self, mode):
        """Set drawing mode"""
        self.canvas.set_mode(mode)
        
        # Update button states
        self.btn_mode_label.setChecked(mode == 'label')
        self.btn_mode_anchor.setChecked(mode == 'anchor')
        self.btn_mode_value.setChecked(mode == 'value')
    
    def on_box_created(self, box):
        """Handle new box creation"""
        self.save_current_page_boxes()
        self.update_box_list()
    
    def on_box_selected(self, box):
        """Handle box selection"""
        if box and box.box_type == 'label':
            self.canvas.set_active_parent(box)
        self.update_box_list()
    
    def on_box_list_clicked(self, item):
        """Handle click on box list item"""
        # Find box by name
        text = item.text()
        for box in self.canvas.boxes:
            if box.name in text:
                self.canvas.selected_box = box
                if box.box_type == 'label':
                    self.canvas.set_active_parent(box)
                self.canvas.update()
                return
            for child in box.children:
                if child.name in text:
                    self.canvas.selected_box = child
                    self.canvas.update()
                    return
    
    def update_box_list(self):
        """Update the box list widget"""
        self.box_tree.clear()
        for box in self.canvas.boxes:
            self.box_tree.addItem(f"📦 [L] {box.name}")
            for child in box.children:
                prefix = "🎯" if child.box_type == 'anchor' else "📝"
                type_char = "A" if child.box_type == 'anchor' else "V"
                self.box_tree.addItem(f"    {prefix} [{type_char}] {child.name}")
    
    def delete_selected_box(self):
        """Delete selected box"""
        self.canvas.delete_selected_box()
        self.save_current_page_boxes()
        self.update_box_list()
    
    def clear_all_boxes(self):
        """Clear all boxes on current page"""
        reply = QMessageBox.question(self, "Clear All", 
            "Delete all boxes on this page?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.canvas.clear_boxes()
            self.save_current_page_boxes()
            self.update_box_list()
    
    def load_template_list(self):
        """Load template names into combo box"""
        self.template_combo.clear()
        session = SessionLocal()
        templates = session.query(OCRTemplate).all()
        for t in templates:
            self.template_combo.addItem(t.name, t.id)
        session.close()
    
    def load_template_for_editing(self):
        """Load an existing template's boxes onto the current PDF page for editing.
        
        This allows adding new labels/anchors to existing templates.
        The template works on ANY loaded PDF - boxes are position-based patterns.
        """
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            QMessageBox.warning(self, "No PDF", "Please load a PDF first.")
            return
        
        template_id = self.template_combo.currentData()
        template_name = self.template_combo.currentText()
        
        if not template_id:
            QMessageBox.warning(self, "No Template", "Please select a template to load.")
            return
        
        session = SessionLocal()
        template = session.query(OCRTemplate).filter(OCRTemplate.id == template_id).first()
        
        if not template:
            QMessageBox.warning(self, "Error", "Template not found.")
            session.close()
            return
        
        # Clear current boxes
        self.page_boxes.clear()
        self.canvas.clear_boxes()
        
        # Load boxes from template
        loaded_count = 0
        current_key = (self.current_pdf_index, self.current_page_index)
        current_page_rotation = self.page_rotations.get(current_key, 0)
        
        for ocr_page in template.pages:
            # Get stored page info
            stored_rotation = getattr(ocr_page, 'page_rotation', 0) or 0
            
            # Load label boxes for this page
            label_boxes = session.query(LabeledBox).filter(
                LabeledBox.page_id == ocr_page.id,
                LabeledBox.box_type == 'label'
            ).all()
            
            for db_label in label_boxes:
                # Create OCRBox from database - coords are stored in raw PDF space
                label_box = self._db_box_to_ocr_box(db_label, stored_rotation, current_page_rotation)
                
                # Add to page_boxes (in PDF coords, scale=1.0)
                if current_key not in self.page_boxes:
                    self.page_boxes[current_key] = []
                self.page_boxes[current_key].append(label_box)
                loaded_count += 1
        
        session.close()
        
        # Set template name for easy update
        self.template_name_input.setText(template_name)
        
        # Reload page to display boxes
        self.render_current_page()
        self.update_box_list()
        
        QMessageBox.information(self, "Loaded", 
            f"Loaded {loaded_count} label boxes from '{template_name}'.\n"
            f"Add new boxes and click 'Save Template' to update.")
    
    def _db_box_to_ocr_box(self, db_box, stored_rotation, current_rotation):
        """Convert a database LabeledBox to an OCRBox.
        
        Stored coordinates are in RAW PDF space. We load them directly since
        we're storing in PDF coords (scale=1.0) before display scaling.
        """
        # Create rect from stored coords (these are raw PDF coords)
        rect = QRectF(db_box.x, db_box.y, db_box.width, db_box.height)
        
        ocr_box = OCRBox(rect, db_box.name, db_box.box_type)
        ocr_box.id = db_box.id
        
        # Recursively load children
        for child_db in db_box.children:
            child_box = self._db_box_to_ocr_box(child_db, stored_rotation, current_rotation)
            child_box.parent = ocr_box
            ocr_box.children.append(child_box)
        
        return ocr_box
    
    def save_template(self):
        """Save current template to database"""
        name = self.template_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Warning", "Please enter a template name.")
            return
        
        # Save current page boxes first
        self.save_current_page_boxes()
        
        # Check if any boxes exist
        total_boxes = sum(len(boxes) for boxes in self.page_boxes.values())
        if total_boxes == 0:
            QMessageBox.warning(self, "Warning", "No boxes to save. Draw some boxes first.")
            return
        
        session = SessionLocal()
        
        # Check for existing template
        existing = session.query(OCRTemplate).filter(OCRTemplate.name == name).first()
        if existing:
            reply = QMessageBox.question(self, "Overwrite?", 
                f"Template '{name}' exists. Overwrite?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                session.close()
                return
            session.delete(existing)
            session.commit()
        
        # Create template
        template = OCRTemplate(name=name)
        session.add(template)
        session.commit()
        
        # Save pages and boxes
        order_idx = 0
        for (pdf_idx, page_idx), boxes in self.page_boxes.items():
            if not boxes:
                continue
            
            filename, doc, path = self.loaded_pdfs[pdf_idx]
            page_width, page_height = self.page_dimensions[(pdf_idx, page_idx)]
            page_rotation = self.page_rotations.get((pdf_idx, page_idx), 0)
            
            ocr_page = OCRPage(
                template_id=template.id,
                pdf_filename=filename,
                page_number=page_idx,
                page_width=page_width,
                page_height=page_height,
                page_rotation=page_rotation,
                order_index=order_idx
            )
            session.add(ocr_page)
            session.commit()
            order_idx += 1
            
            # Save boxes
            for box in boxes:
                self._save_box_to_db(session, ocr_page.id, box, None, pdf_idx, page_idx)
        
        session.commit()
        session.close()
        
        QMessageBox.information(self, "Success", f"Template '{name}' saved!")
        self.load_template_list()
    
    def transform_visual_to_pdf_coords(self, x, y, w, h, page_width, page_height, rotation):
        """
        Transform visual coordinates (what user sees/draws) to PDF internal coordinates.
        The visual display is already derotated by PyMuPDF, so we need to reverse that
        when working with the PDF's internal coordinate system.
        
        For a page rotated 90° anti-clockwise (rotation=90):
        - Visual shows it correctly rotated
        - User draws on the visual
        - We need to store coordinates that work when the rotation is removed
        """
        if rotation == 0:
            return x, y, w, h
        elif rotation == 90:
            # 90° anti-clockwise: visual (x, y) -> PDF (y, page_width - x - w)
            # Dimensions swap: w, h -> h, w
            new_x = y
            new_y = page_width - x - w
            new_w = h
            new_h = w
            return new_x, new_y, new_w, new_h
        elif rotation == 180:
            # 180°: visual (x, y) -> PDF (page_width - x - w, page_height - y - h)
            new_x = page_width - x - w
            new_y = page_height - y - h
            return new_x, new_y, w, h
        elif rotation == 270:
            # 270° (or 90° clockwise): visual (x, y) -> PDF (page_height - y - h, x)
            new_x = page_height - y - h
            new_y = x
            new_w = h
            new_h = w
            return new_x, new_y, new_w, new_h
        else:
            return x, y, w, h
    
    def _save_box_to_db(self, session, page_id, box, parent_id, pdf_idx=None, page_idx=None):
        """Recursively save box and its children, extracting anchor text from PDF
        
        IMPORTANT: 
        - We save TRANSFORMED (raw PDF) coordinates, not visual coordinates
        - This ensures offset calculation works correctly during extraction 
        - since search_for() returns raw PDF coordinates too
        """
        scale = self.canvas.scale_factor
        
        # Get visual coordinates (scaled back to PDF visual size)
        vis_x = box.rect.x() / scale
        vis_y = box.rect.y() / scale
        vis_w = box.rect.width() / scale
        vis_h = box.rect.height() / scale
        
        # Get page rotation and dimensions for coordinate transformation
        key = (pdf_idx, page_idx) if pdf_idx is not None and page_idx is not None else None
        rotation = self.page_rotations.get(key, 0) if key else 0
        page_dims = self.page_dimensions.get(key, (0, 0)) if key else (0, 0)
        page_width, page_height = page_dims
        
        # Transform visual coordinates to RAW PDF coordinates
        # This is crucial: we store in the same coordinate space that search_for uses
        pdf_x, pdf_y, pdf_w, pdf_h = self.transform_visual_to_pdf_coords(
            vis_x, vis_y, vis_w, vis_h, page_width, page_height, rotation
        )
        
        # Get the actual text for anchor boxes from the PDF
        box_name = box.name
        if box.box_type == 'anchor' and pdf_idx is not None and page_idx is not None:
            filename, doc, path = self.loaded_pdfs[pdf_idx]
            page = doc.load_page(page_idx)
            
            # Use transformed coordinates for text extraction
            rect = fitz.Rect(pdf_x, pdf_y, pdf_x + pdf_w, pdf_y + pdf_h)
            extracted_text = page.get_text("text", clip=rect).strip()
            
            print(f"[DEBUG] Anchor: visual=({vis_x:.1f},{vis_y:.1f}), raw=({pdf_x:.1f},{pdf_y:.1f}), rot={rotation}")
            
            if extracted_text:
                box_name = extracted_text
                print(f"[DEBUG] Anchor text: '{extracted_text}'")
            else:
                # Try with expanded rect
                expanded_rect = rect + (-5, -5, 5, 5)
                extracted_text = page.get_text("text", clip=expanded_rect).strip()
                if extracted_text:
                    box_name = extracted_text
                    print(f"[DEBUG] Anchor text (expanded): '{extracted_text}'")
                else:
                    print(f"[DEBUG] WARNING: No anchor text extracted")
        
        # Save RAW PDF coordinates - these match what search_for() returns
        # so we can calculate offset directly without rotation transformation
        db_box = LabeledBox(
            page_id=page_id,
            parent_box_id=parent_id,
            name=box_name,
            box_type=box.box_type,
            x=pdf_x,        # RAW PDF X 
            y=pdf_y,        # RAW PDF Y  
            width=pdf_w,    # RAW PDF Width
            height=pdf_h    # RAW PDF Height
        )
        session.add(db_box)
        session.commit()
        
        print(f"[DEBUG] Saved {box.box_type} at raw coords: ({pdf_x:.1f}, {pdf_y:.1f}) {pdf_w:.1f}x{pdf_h:.1f}")
        
        for child in box.children:
            self._save_box_to_db(session, page_id, child, db_box.id, pdf_idx, page_idx)

    
    def run_extraction(self):
        """Run extraction on multiple PDFs using selected template"""
        if self.template_combo.count() == 0:
            QMessageBox.warning(self, "No Template", "Please save or select a template first.")
            return
        
        # Select PDFs for extraction
        paths, _ = QFileDialog.getOpenFileNames(self, "Select PDFs for Extraction", "", "PDF Files (*.pdf)")
        if not paths:
            return
        
        template_id = self.template_combo.currentData()
        session = SessionLocal()
        template = session.query(OCRTemplate).filter(OCRTemplate.id == template_id).first()
        
        if not template:
            QMessageBox.warning(self, "Error", "Template not found.")
            session.close()
            return
        
        # Get all label boxes from template to build column headers
        all_labels = []
        label_info = {}  # label_id -> (label_name, anchors, values, page_dims)
        
        for ocr_page in template.pages:
            label_boxes = session.query(LabeledBox).filter(
                LabeledBox.page_id == ocr_page.id,
                LabeledBox.box_type == 'label'
            ).all()
            
            for label_box in label_boxes:
                if label_box.name not in [l['name'] for l in all_labels]:
                    anchors = [b for b in label_box.children if b.box_type == 'anchor']
                    values = [b for b in label_box.children if b.box_type == 'value']
                    
                    all_labels.append({
                        'name': label_box.name,
                        'id': label_box.id
                    })
                    label_info[label_box.id] = {
                        'name': label_box.name,
                        'anchors': anchors,
                        'values': values,
                        'page_width': ocr_page.page_width,
                        'page_height': ocr_page.page_height,
                        'page_rotation': getattr(ocr_page, 'page_rotation', 0) or 0
                    }
        
        if not all_labels:
            QMessageBox.warning(self, "No Labels", "Template has no label boxes defined.")
            session.close()
            return
        
        # Setup result table with dynamic columns: PDF, then Anchor + Value for each label
        self.result_table.clear()
        self.result_table.setRowCount(0)
        
        # Create column headers: PDF Filename, Label1_Anchor, Label1_Value, Label2_Anchor, Label2_Value, ...
        columns = ["PDF Filename"]
        for label in all_labels:
            columns.append(f"{label['name']}_Anchor")
            columns.append(f"{label['name']}_Value")
        
        self.result_table.setColumnCount(len(columns))
        self.result_table.setHorizontalHeaderLabels(columns)
        
        self.extraction_results = []
        self.extraction_screenshots = []  # Collect for backup PDF
        extracted_count = 0
        
        try:
            for pdf_path in paths:
                doc = fitz.open(pdf_path)
                pdf_filename = os.path.basename(pdf_path)
                
                # Extract data for this PDF - anchor and value per label
                row_data = {'PDF Filename': pdf_filename}
                
                for label in all_labels:
                    info = label_info[label['id']]
                    
                    # Find this label's value in the PDF
                    match = self._find_box_on_pages(
                        doc, None, info['anchors'], info['values'],
                        info['page_width'], info['page_height'], info['page_rotation']
                    )
                    
                    if match:
                        row_data[f"{label['name']}_Anchor"] = match['anchor_text']
                        row_data[f"{label['name']}_Value"] = match['value_text']
                        extracted_count += 1
                        
                        # Collect screenshot data for backup PDF
                        if 'value_rect' in match and match['value_rect']:
                            self.extraction_screenshots.append({
                                'pdf_path': pdf_path,
                                'page_idx': match['page'],
                                'value_rect': match['value_rect'],
                                'label_name': label['name'],
                                'value_text': match['value_text'],
                                'pdf_filename': pdf_filename
                            })
                    else:
                        row_data[f"{label['name']}_Anchor"] = ""
                        row_data[f"{label['name']}_Value"] = ""
                
                doc.close()
                
                # Add row to table
                row_idx = self.result_table.rowCount()
                self.result_table.insertRow(row_idx)
                
                for col_idx, col_name in enumerate(columns):
                    value = row_data.get(col_name, "")
                    self.result_table.setItem(row_idx, col_idx, QTableWidgetItem(value))
                
                self.extraction_results.append(row_data)
            
            # Resize columns to fit content
            self.result_table.resizeColumnsToContents()
            
            # Generate extraction backup PDF automatically
            if self.extraction_screenshots:
                self._generate_extraction_backup()
            
            QMessageBox.information(self, "Complete", 
                f"Extraction complete!\n{len(paths)} PDFs processed\n{extracted_count} values extracted")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Extraction failed: {e}")
            import traceback
            traceback.print_exc()
        finally:
            session.close()
    
    def _generate_extraction_backup(self):
        """Generate backup PDF with screenshots of extracted values from processed PDFs"""
        if not self.extraction_screenshots:
            return
        
        # Ask user where to save
        path, _ = QFileDialog.getSaveFileName(self, "Save Extraction Backup PDF", 
            "extraction_backup.pdf", "PDF Files (*.pdf)")
        if not path:
            return
        
        try:
            backup_doc = fitz.open()
            
            for data in self.extraction_screenshots:
                pdf_path = data['pdf_path']
                page_idx = data['page_idx']
                value_rect = data['value_rect']
                label_name = data['label_name']
                value_text = data['value_text']
                pdf_filename = data['pdf_filename']
                
                # Open source PDF and get page
                src_doc = fitz.open(pdf_path)
                src_page = src_doc.load_page(page_idx)
                
                # Expand rect a bit for context
                expanded_rect = value_rect + (-20, -20, 20, 20)
                expanded_rect = expanded_rect & src_page.rect  # Clip to page bounds
                
                # Get cropped screenshot (scale 2x for quality)
                mat = fitz.Matrix(2, 2)
                pix = src_page.get_pixmap(matrix=mat, clip=expanded_rect)
                
                # Create new page in backup doc
                # Page size: cropped image width + margins, enough height for image + text
                img_width = pix.width
                img_height = pix.height
                page_width = img_width + 40  # 20px margins
                page_height = img_height + 120  # Space for header/footer text
                
                backup_page = backup_doc.new_page(width=page_width, height=page_height)
                
                # Add header text
                header_text = f"Source: {pdf_filename} | Page: {page_idx + 1} | Label: {label_name}"
                text_point = fitz.Point(20, 25)
                backup_page.insert_text(text_point, header_text, fontsize=10)
                
                # Insert the cropped image
                img_rect = fitz.Rect(20, 40, 20 + img_width, 40 + img_height)
                backup_page.insert_image(img_rect, pixmap=pix)
                
                # Add extracted text at bottom
                text_y = 40 + img_height + 20
                value_preview = value_text[:100] + "..." if len(value_text) > 100 else value_text
                backup_page.insert_text(fitz.Point(20, text_y), f"Value: {value_preview}", fontsize=9)
                
                src_doc.close()
            
            backup_doc.save(path)
            backup_doc.close()
            
            print(f"[DEBUG] Extraction backup saved to: {path}")
            
        except Exception as e:
            print(f"[DEBUG] Error generating extraction backup: {e}")
            import traceback
            traceback.print_exc()
    
    def _find_box_on_pages(self, doc, label_box, anchors, values, base_width, base_height, template_rotation=0):
        """
        Find box content using MULTI-ANCHOR approach.
        
        When anchor text appears multiple times in a document, we use additional anchors
        to disambiguate and find the correct instance:
        1. Find all instances of the PRIMARY anchor (first anchor in list)
        2. For each instance, check if SECONDARY anchors are at expected relative positions
        3. If all anchors match (within tolerance), extract value at stored offset
        
        This allows extraction to work even when anchor text like "Name:" appears multiple times.
        """
        POSITION_TOLERANCE = 50  # pixels tolerance for anchor position matching
        
        if not anchors:
            return None
        
        # Get first anchor and value for primary matching
        first_anchor = anchors[0] if anchors else None
        first_value = values[0] if values else None
        
        if not first_anchor or not first_value:
            return None
        
        # Primary anchor text (what we search for)
        primary_anchor_text = (first_anchor.name or "").strip()
        if not primary_anchor_text:
            return None
        
        print(f"[DEBUG] Multi-anchor search: primary='{primary_anchor_text}'")
        
        # Build list of secondary anchors with their expected offsets from primary
        secondary_anchors = []
        for anchor in anchors[1:]:
            anchor_text = (anchor.name or "").strip()
            if anchor_text:
                secondary_anchors.append({
                    'text': anchor_text,
                    'expected_dx': anchor.x - first_anchor.x,
                    'expected_dy': anchor.y - first_anchor.y
                })
                print(f"[DEBUG] Secondary anchor: '{anchor_text}' at offset ({anchor.x - first_anchor.x:.1f}, {anchor.y - first_anchor.y:.1f})")
        
        # Calculate value offset from primary anchor
        value_dx = first_value.x - first_anchor.x
        value_dy = first_value.y - first_anchor.y
        value_w = first_value.width
        value_h = first_value.height
        
        print(f"[DEBUG] Value offset from primary: dx={value_dx:.1f}, dy={value_dy:.1f}, w={value_w:.1f}, h={value_h:.1f}")
        
        # Search all pages
        for page_idx in range(len(doc)):
            try:
                page = doc.load_page(page_idx)
                
                # Find ALL instances of primary anchor
                primary_instances = page.search_for(primary_anchor_text)
                
                # Fallback: search for first word if full text not found
                if not primary_instances:
                    words = primary_anchor_text.split()
                    if words:
                        primary_instances = page.search_for(words[0])
                
                if not primary_instances:
                    continue
                
                print(f"[DEBUG] Page {page_idx}: Found {len(primary_instances)} instance(s) of '{primary_anchor_text}'")
                
                # Check each primary instance
                for primary_rect in primary_instances:
                    print(f"[DEBUG] Checking primary at {primary_rect}")
                    
                    # If we have secondary anchors, verify they're at expected positions
                    all_secondary_match = True
                    
                    for secondary in secondary_anchors:
                        sec_text = secondary['text']
                        expected_dx = secondary['expected_dx']
                        expected_dy = secondary['expected_dy']
                        
                        # Find all instances of this secondary anchor
                        sec_instances = page.search_for(sec_text)
                        
                        # Check if any instance is at expected offset (within tolerance)
                        found_matching = False
                        for sec_rect in sec_instances:
                            actual_dx = sec_rect.x0 - primary_rect.x0
                            actual_dy = sec_rect.y0 - primary_rect.y0
                            
                            dx_diff = abs(actual_dx - expected_dx)
                            dy_diff = abs(actual_dy - expected_dy)
                            
                            if dx_diff <= POSITION_TOLERANCE and dy_diff <= POSITION_TOLERANCE:
                                print(f"[DEBUG] Secondary '{sec_text}' matched at offset diff ({dx_diff:.1f}, {dy_diff:.1f})")
                                found_matching = True
                                break
                        
                        if not found_matching:
                            print(f"[DEBUG] Secondary '{sec_text}' NOT found at expected offset")
                            all_secondary_match = False
                            break
                    
                    # If all secondary anchors matched (or there are none), extract value
                    if all_secondary_match:
                        print(f"[DEBUG] All anchors matched! Extracting value...")
                        
                        # Calculate value rect
                        value_rect = fitz.Rect(
                            primary_rect.x0 + value_dx,
                            primary_rect.y0 + value_dy,
                            primary_rect.x0 + value_dx + value_w,
                            primary_rect.y0 + value_dy + value_h
                        )
                        
                        value_rect = value_rect.normalize()
                        value_rect = value_rect & page.rect
                        
                        print(f"[DEBUG] Value rect: {value_rect}")
                        
                        value_text = ""
                        if not value_rect.is_empty:
                            # Try to get text from value region
                            text = page.get_text("text", clip=value_rect).strip()
                            
                            # If no text, try expanded rect
                            if not text:
                                expanded = (value_rect + (-10, -10, 10, 10)) & page.rect
                                text = page.get_text("text", clip=expanded).strip()
                            
                            # Fallback to text blocks
                            if not text:
                                blocks = page.get_text("blocks", clip=value_rect)
                                for block in blocks:
                                    if len(block) > 4:
                                        text += str(block[4]).strip() + " "
                                text = text.strip()
                            
                            if text:
                                value_text = text
                                print(f"[DEBUG] Extracted: '{value_text[:50]}'")
                        
                        if value_text:
                            return {
                                'page': page_idx,
                                'anchor_text': primary_anchor_text,
                                'value_text': value_text,
                                'value_rect': value_rect,  # For extraction backup screenshots
                                'target_rotation': page.rotation
                            }
                
            except Exception as e:
                print(f"[DEBUG] Error on page {page_idx}: {e}")
                import traceback
                traceback.print_exc()
                continue
        
        return None
    
    def export_excel(self):
        """Export results to Excel"""
        if not self.extraction_results:
            QMessageBox.warning(self, "No Data", "No extraction results to export.")
            return
        
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel Files (*.xlsx)")
        if path:
            df = pd.DataFrame(self.extraction_results)
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Success", f"Exported to {path}")
    
    def export_backup_pdf(self):
        """Export backup PDF with screenshots of all boxes"""
        self.save_current_page_boxes()
        
        total_boxes = sum(len(boxes) for boxes in self.page_boxes.values())
        if total_boxes == 0:
            QMessageBox.warning(self, "No Boxes", "No boxes to export.")
            return
        
        path, _ = QFileDialog.getSaveFileName(self, "Save Backup PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        
        try:
            # Create new PDF
            backup_doc = fitz.open()
            
            for (pdf_idx, page_idx), boxes in self.page_boxes.items():
                if not boxes:
                    continue
                
                filename, doc, orig_path = self.loaded_pdfs[pdf_idx]
                page = doc.load_page(page_idx)
                
                # Get page as image (scale 2x for quality)
                mat = fitz.Matrix(2, 2)
                pix = page.get_pixmap(matrix=mat)
                
                for box in boxes:
                    # Create a page for each label box
                    self._add_box_to_backup(backup_doc, pix, box, filename, page_idx, mat)
            
            backup_doc.save(path)
            backup_doc.close()
            
            QMessageBox.information(self, "Success", f"Backup PDF saved to {path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create backup: {e}")
            import traceback
            traceback.print_exc()
    
    def _add_box_to_backup(self, backup_doc, pix, box, filename, page_idx, mat):
        """Add a box screenshot to backup PDF"""
        scale = 2  # We use 2x scale for pixmap
        
        # Calculate box coords in pixmap space (need to account for canvas scaling)
        canvas_scale = self.canvas.scale_factor
        
        # Convert canvas coords to original PDF coords, then to pixmap coords
        x0 = int((box.rect.x() / canvas_scale) * scale)
        y0 = int((box.rect.y() / canvas_scale) * scale)
        x1 = int(((box.rect.x() + box.rect.width()) / canvas_scale) * scale)
        y1 = int(((box.rect.y() + box.rect.height()) / canvas_scale) * scale)
        
        # Add margin
        margin = 30
        x0 = max(0, x0 - margin)
        y0 = max(0, y0 - margin)
        x1 = min(pix.width, x1 + margin)
        y1 = min(pix.height, y1 + margin)
        
        # Create cropped image using PIL approach (more reliable)
        crop_rect = fitz.IRect(x0, y0, x1, y1)
        
        # Create new pixmap for the cropped area
        cropped = fitz.Pixmap(pix.colorspace, crop_rect, pix.alpha)
        cropped.copy(pix, crop_rect)
        
        # Create new page in backup doc
        width = x1 - x0 + 40
        height = y1 - y0 + 80  # Extra space for labels
        
        new_page = backup_doc.new_page(width=width, height=height)
        
        # Add title
        title = f"[{box.box_type.upper()}] {box.name}"
        subtitle = f"Source: {filename}, Page {page_idx + 1}"
        
        new_page.insert_text((10, 20), title, fontsize=14, fontname="helv")
        new_page.insert_text((10, 35), subtitle, fontsize=10, fontname="helv", color=(0.5, 0.5, 0.5))
        
        # Insert cropped image
        img_rect = fitz.Rect(10, 50, width - 10, height - 10)
        new_page.insert_image(img_rect, pixmap=cropped)
        
        # Add info about children
        if box.children:
            child_info = ", ".join([f"{c.box_type}: {c.name}" for c in box.children])
            new_page.insert_text((10, height - 15), f"Children: {child_info}", 
                                fontsize=8, fontname="helv", color=(0.4, 0.4, 0.4))


# ============================================================================
# SCHEDULER MODULE
# ============================================================================

class SchedulerModule(QWidget):
    def __init__(self):
        super().__init__()
        self.scheduler = BackgroundScheduler()
        self.scheduler.start()
        self.setup_ui()
        self.load_jobs_from_db()
        self.check_missed_jobs()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        title = QLabel("⏰ Scheduler")
        title.setObjectName("moduleTitle")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        layout.addWidget(title)
        
        btn_add = QPushButton("➕ Add Job")
        btn_add.clicked.connect(self.add_job_dialog)
        layout.addWidget(btn_add)
        
        self.job_table = QTableWidget(0, 5)
        self.job_table.setHorizontalHeaderLabels(["Name", "Type", "Next Run", "Status", "Actions"])
        self.job_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.job_table)
    
    def load_jobs_from_db(self):
        """Load all jobs from database and add to scheduler"""
        session = SessionLocal()
        jobs = session.query(Job).all()
        
        for job_db in jobs:
            if job_db.enabled:
                self.schedule_job(job_db)
        
        session.close()
        self.refresh_job_list()
    
    def check_missed_jobs(self):
        """Check for and execute missed jobs on startup"""
        session = SessionLocal()
        now = datetime.datetime.now()
        
        jobs = session.query(Job).filter(Job.enabled == True, Job.next_run != None).all()
        
        for job_db in jobs:
            if job_db.next_run < now:
                # Job was missed
                grace = datetime.timedelta(seconds=job_db.misfire_grace_time)
                if now - job_db.next_run <= grace:
                    print(f"Executing missed job: {job_db.name}")
                    self.execute_job(job_db)
        
        session.close()
    
    def schedule_job(self, job_db):
        """Add job to APScheduler based on database record"""
        job_id = f"job_{job_db.id}"
        
        try:
            if job_db.job_type == "one_time":
                trigger = DateTrigger(run_date=job_db.run_date)
            elif job_db.job_type == "recurring":
                if job_db.recurrence == "interval":
                    trigger = IntervalTrigger(seconds=job_db.interval_seconds)
                elif job_db.recurrence == "daily":
                    h, m = map(int, job_db.recurrence_time.split(":"))
                    trigger = CronTrigger(hour=h, minute=m)
                elif job_db.recurrence == "weekly":
                    h, m = map(int, job_db.recurrence_time.split(":"))
                    trigger = CronTrigger(day_of_week=job_db.day_of_week, hour=h, minute=m)
                elif job_db.recurrence == "monthly":
                    h, m = map(int, job_db.recurrence_time.split(":"))
                    trigger = CronTrigger(day=job_db.day_of_month, hour=h, minute=m)
                else:
                    return
            else:
                return
            
            self.scheduler.add_job(
                lambda: self.execute_job_by_id(job_db.id),
                trigger,
                id=job_id,
                name=job_db.name,
                misfire_grace_time=job_db.misfire_grace_time
            )
            
            # Update next_run in database
            job = self.scheduler.get_job(job_id)
            if job:
                session = SessionLocal()
                db_job = session.query(Job).get(job_db.id)
                db_job.next_run = job.next_run_time
                session.commit()
                session.close()
                
        except Exception as e:
            print(f"Error scheduling job {job_db.name}: {e}")
    
    def execute_job_by_id(self, job_id):
        """Execute job by database ID"""
        session = SessionLocal()
        job_db = session.query(Job).get(job_id)
        if job_db:
            self.execute_job(job_db)
        session.close()
    
    def execute_job(self, job_db):
        """Execute the job script"""
        try:
            result = subprocess.run(job_db.script_path, shell=True, capture_output=True, text=True)
            print(f"Job '{job_db.name}' executed. Return code: {result.returncode}")
            
            # Update last_run
            session = SessionLocal()
            db_job = session.query(Job).get(job_db.id)
            db_job.last_run = datetime.datetime.now()
            
            # For one-time jobs, disable after execution
            if job_db.job_type == "one_time":
                db_job.enabled = False
                # Remove from scheduler
                try:
                    self.scheduler.remove_job(f"job_{job_db.id}")
                except:
                    pass
            
            session.commit()
            session.close()
            self.refresh_job_list()
            
        except Exception as e:
            print(f"Job '{job_db.name}' failed: {e}")
    
    def add_job_dialog(self):
        """Enhanced dialog for adding jobs"""
        from PySide6.QtWidgets import QDateTimeEdit, QRadioButton, QButtonGroup, QCheckBox
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Job")
        dialog.resize(500, 600)
        layout = QVBoxLayout(dialog)
        
        # Job Name
        layout.addWidget(QLabel("Job Name:"))
        name_input = QLineEdit()
        layout.addWidget(name_input)
        
        # Script Path
        layout.addWidget(QLabel("Script Path:"))
        script_layout = QHBoxLayout()
        script_input = QLineEdit()
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(lambda: script_input.setText(
            QFileDialog.getOpenFileName(dialog, "Select Script")[0]))
        script_layout.addWidget(script_input)
        script_layout.addWidget(btn_browse)
        layout.addLayout(script_layout)
        
        # Job Type
        layout.addWidget(QLabel("Job Type:"))
        type_group = QButtonGroup(dialog)
        radio_onetime = QRadioButton("One-Time")
        radio_recurring = QRadioButton("Recurring")
        radio_onetime.setChecked(True)
        type_group.addButton(radio_onetime)
        type_group.addButton(radio_recurring)
        type_layout = QHBoxLayout()
        type_layout.addWidget(radio_onetime)
        type_layout.addWidget(radio_recurring)
        layout.addLayout(type_layout)
        
        # One-Time Section
        onetime_widget = QWidget()
        onetime_layout = QVBoxLayout(onetime_widget)
        onetime_layout.addWidget(QLabel("Run Date & Time:"))
        datetime_picker = QDateTimeEdit()
        datetime_picker.setDateTime(datetime.datetime.now() + datetime.timedelta(hours=1))
        datetime_picker.setDisplayFormat("yyyy-MM-dd HH:mm")
        onetime_layout.addWidget(datetime_picker)
        layout.addWidget(onetime_widget)
        
        # Recurring Section
        recurring_widget = QWidget()
        recurring_layout = QVBoxLayout(recurring_widget)
        
        recurring_layout.addWidget(QLabel("Recurrence Type:"))
        recurrence_combo = QComboBox()
        recurrence_combo.addItems(["Interval", "Daily", "Weekly", "Monthly"])
        recurring_layout.addWidget(recurrence_combo)
        
        # Interval settings
        interval_widget = QWidget()
        interval_layout = QHBoxLayout(interval_widget)
        interval_layout.addWidget(QLabel("Every:"))
        interval_spin = QSpinBox()
        interval_spin.setRange(1, 86400)
        interval_spin.setValue(1)
        interval_layout.addWidget(interval_spin)
        interval_unit = QComboBox()
        interval_unit.addItems(["Seconds", "Minutes", "Hours"])
        interval_unit.setCurrentText("Hours")
        interval_layout.addWidget(interval_unit)
        recurring_layout.addWidget(interval_widget)
        
        # Time picker for daily/weekly/monthly
        time_widget = QWidget()
        time_layout = QHBoxLayout(time_widget)
        time_layout.addWidget(QLabel("Time:"))
        time_picker = QLineEdit()
        time_picker.setText("09:00")
        time_picker.setPlaceholderText("HH:MM")
        time_layout.addWidget(time_picker)
        recurring_layout.addWidget(time_widget)
        
        # Weekly: Day selection
        weekly_widget = QWidget()
        weekly_layout = QVBoxLayout(weekly_widget)
        weekly_layout.addWidget(QLabel("Days of Week:"))
        day_checks = []
        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for i, day in enumerate(day_names):
            cb = QCheckBox(day)
            cb.setProperty("day_index", i)
            day_checks.append(cb)
            weekly_layout.addWidget(cb)
        recurring_layout.addWidget(weekly_widget)
        
        # Monthly: Day of month
        monthly_widget = QWidget()
        monthly_layout = QHBoxLayout(monthly_widget)
        monthly_layout.addWidget(QLabel("Day of Month:"))
        day_spin = QSpinBox()
        day_spin.setRange(1, 31)
        day_spin.setValue(1)
        monthly_layout.addWidget(day_spin)
        monthly_layout.addWidget(day_spin)
        
        # Business Day Checkbox (Placeholder for now)
        business_day_cb = QCheckBox("Business Day Only (Mon-Fri)")
        monthly_layout.addWidget(business_day_cb)
        
        recurring_layout.addWidget(monthly_widget)
        
        # Summary Label
        summary_label = QLabel("Summary: Runs once at specified time.")
        summary_label.setStyleSheet("color: #666; font-style: italic; margin-top: 10px;")
        summary_label.setWordWrap(True)
        layout.addWidget(summary_label)
        
        # Show/hide based on recurrence type
        def update_recurrence_widgets():
            rec_type = recurrence_combo.currentText()
            interval_widget.setVisible(rec_type == "Interval")
            time_widget.setVisible(rec_type in ["Daily", "Weekly", "Monthly"])
            weekly_widget.setVisible(rec_type == "Weekly")
            monthly_widget.setVisible(rec_type == "Monthly")
            update_summary()
            
        def update_summary():
            if radio_onetime.isChecked():
                summary_label.setText(f"Summary: Runs once on {datetime_picker.dateTime().toString('yyyy-MM-dd HH:mm')}")
                return
                
            rec_type = recurrence_combo.currentText()
            if rec_type == "Interval":
                summary_label.setText(f"Summary: Runs every {interval_spin.value()} {interval_unit.currentText().lower()}")
            elif rec_type == "Daily":
                summary_label.setText(f"Summary: Runs every day at {time_picker.text()}")
            elif rec_type == "Weekly":
                days = [cb.text() for cb in day_checks if cb.isChecked()]
                day_str = ", ".join(days) if days else "selected days"
                summary_label.setText(f"Summary: Runs every {day_str} at {time_picker.text()}")
            elif rec_type == "Monthly":
                day = day_spin.value()
                suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
                biz_str = " (Business Day)" if business_day_cb.isChecked() else ""
                summary_label.setText(f"Summary: Runs on the {day}{suffix}{biz_str} of every month at {time_picker.text()}")

        # Connect signals to update summary
        radio_onetime.toggled.connect(update_summary)
        datetime_picker.dateTimeChanged.connect(update_summary)
        recurrence_combo.currentTextChanged.connect(update_recurrence_widgets)
        interval_spin.valueChanged.connect(update_summary)
        interval_unit.currentTextChanged.connect(update_summary)
        time_picker.textChanged.connect(update_summary)
        day_spin.valueChanged.connect(update_summary)
        business_day_cb.stateChanged.connect(update_summary)
        for cb in day_checks:
            cb.stateChanged.connect(update_summary)
        
        update_recurrence_widgets()
        
        layout.addWidget(recurring_widget)
        recurring_widget.setVisible(False)
        
        # Toggle visibility based on job type
        def update_job_type_widgets():
            is_onetime = radio_onetime.isChecked()
            onetime_widget.setVisible(is_onetime)
            recurring_widget.setVisible(not is_onetime)
        
        radio_onetime.toggled.connect(update_job_type_widgets)
        
        # Misfire Grace Time
        layout.addWidget(QLabel("Misfire Grace Time (minutes):"))
        grace_spin = QSpinBox()
        grace_spin.setRange(1, 1440)
        grace_spin.setValue(5)
        layout.addWidget(grace_spin)
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec() == QDialog.Accepted:
            self.save_job(
                name_input.text(),
                script_input.text(),
                radio_onetime.isChecked(),
                datetime_picker.dateTime().toPython(),
                recurrence_combo.currentText(),
                interval_spin.value(),
                interval_unit.currentText(),
                time_picker.text(),
                day_checks,
                day_spin.value(),
                grace_spin.value()
            )
    
    def save_job(self, name, script, is_onetime, run_datetime, rec_type, 
                 interval_val, interval_unit, rec_time, day_checks, day_of_month, grace_min):
        """Save job to database and schedule it"""
        if not name or not script:
            QMessageBox.warning(self, "Warning", "Name and script path are required")
            return
        
        session = SessionLocal()
        
        job_db = Job()
        job_db.name = name
        job_db.script_path = script
        job_db.misfire_grace_time = grace_min * 60
        
        if is_onetime:
            job_db.job_type = "one_time"
            job_db.run_date = run_datetime
            job_db.next_run = run_datetime
        else:
            job_db.job_type = "recurring"
            job_db.recurrence = rec_type.lower()
            
            if rec_type == "Interval":
                multiplier = {"Seconds": 1, "Minutes": 60, "Hours": 3600}[interval_unit]
                job_db.interval_seconds = interval_val * multiplier
            elif rec_type in ["Daily", "Weekly", "Monthly"]:
                job_db.recurrence_time = rec_time
                
                if rec_type == "Weekly":
                    selected_days = [str(cb.property("day_index")) for cb in day_checks if cb.isChecked()]
                    job_db.day_of_week = ",".join(selected_days)
                elif rec_type == "Monthly":
                    job_db.day_of_month = day_of_month
        
        session.add(job_db)
        session.commit()
        
        # Schedule the job
        self.schedule_job(job_db)
        
        session.close()
        self.refresh_job_list()
        QMessageBox.information(self, "Success", "Job added successfully!")
    
    def refresh_job_list(self):
        """Refresh the job table"""
        session = SessionLocal()
        jobs = session.query(Job).all()
        
        self.job_table.setRowCount(len(jobs))
        
        for row, job in enumerate(jobs):
            self.job_table.setItem(row, 0, QTableWidgetItem(job.name))
            
            job_type_str = "One-Time" if job.job_type == "one_time" else f"Recurring ({job.recurrence})"
            self.job_table.setItem(row, 1, QTableWidgetItem(job_type_str))
            
            next_run_str = job.next_run.strftime("%Y-%m-%d %H:%M") if job.next_run else "N/A"
            self.job_table.setItem(row, 2, QTableWidgetItem(next_run_str))
            
            status_str = "Enabled" if job.enabled else "Disabled"
            self.job_table.setItem(row, 3, QTableWidgetItem(status_str))
            
            # Actions
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(0, 0, 0, 0)
            
            btn_toggle = QPushButton("Disable" if job.enabled else "Enable")
            btn_toggle.clicked.connect(lambda checked, j=job: self.toggle_job(j.id))
            btn_delete = QPushButton("Delete")
            btn_delete.clicked.connect(lambda checked, j=job: self.delete_job(j.id))
            
            actions_layout.addWidget(btn_toggle)
            actions_layout.addWidget(btn_delete)
            
            self.job_table.setCellWidget(row, 4, actions_widget)
        
        session.close()
    
    def toggle_job(self, job_id):
        """Enable or disable a job"""
        session = SessionLocal()
        job = session.query(Job).get(job_id)
        
        if job:
            job.enabled = not job.enabled
            session.commit()
            
            scheduler_job_id = f"job_{job_id}"
            if job.enabled:
                self.schedule_job(job)
            else:
                try:
                    self.scheduler.remove_job(scheduler_job_id)
                except:
                    pass
        
        session.close()
        self.refresh_job_list()
    
    def delete_job(self, job_id):
        """Delete a job"""
        reply = QMessageBox.question(self, "Confirm Delete", 
                                     "Are you sure you want to delete this job?",
                                     QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            session = SessionLocal()
            job = session.query(Job).get(job_id)
            
            if job:
                # Remove from scheduler
                try:
                    self.scheduler.remove_job(f"job_{job_id}")
                except:
                    pass
                
                session.delete(job)
                session.commit()
            
            session.close()
            self.refresh_job_list()

# ============================================================================
# MAIL DRAFTER MODULE
# ============================================================================

class MailDrafterModule(QWidget):
    def __init__(self, pdf_editor_module):
        super().__init__()
        self.pdf_editor = pdf_editor_module
        self.setup_ui()
    
    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Left Panel: Form
        form_panel = QWidget()
        form_layout = QVBoxLayout(form_panel)
        
        title = QLabel("📧 Mail Drafter")
        title.setObjectName("moduleTitle")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        form_layout.addWidget(title)
        
        # Template controls
        template_row = QHBoxLayout()
        self.template_combo = QComboBox()
        self.template_combo.addItem("-- Select Template --")
        self.template_combo.currentIndexChanged.connect(self.load_template)
        template_row.addWidget(self.template_combo)
        btn_save_template = QPushButton("💾 Save as Template")
        btn_save_template.clicked.connect(self.save_template)
        template_row.addWidget(btn_save_template)
        form_layout.addLayout(template_row)
        
        form_layout.addWidget(QLabel("From (Send on Behalf):"))
        self.from_input = QLineEdit()
        self.from_input.setPlaceholderText("Optional: shared.mailbox@company.com")
        form_layout.addWidget(self.from_input)
        
        form_layout.addWidget(QLabel("To:"))
        self.to_input = QLineEdit()
        form_layout.addWidget(self.to_input)
        
        form_layout.addWidget(QLabel("CC:"))
        self.cc_input = QLineEdit()
        self.cc_input.setPlaceholderText("Optional: cc1@email.com; cc2@email.com")
        form_layout.addWidget(self.cc_input)
        
        form_layout.addWidget(QLabel("Subject:"))
        self.subject_input = QLineEdit()
        form_layout.addWidget(self.subject_input)
        
        form_layout.addWidget(QLabel("Body:"))
        self.body_input = QTextEdit()
        form_layout.addWidget(self.body_input)
        
        btn_draft = QPushButton("📝 Generate Draft & Preview")
        btn_draft.setStyleSheet("background-color: #3b82f6; color: white; padding: 10px; font-weight: bold;")
        btn_draft.clicked.connect(self.generate_draft)
        form_layout.addWidget(btn_draft)
        
        layout.addWidget(form_panel, stretch=2)
        
        # Right Panel: Attachments
        attach_panel = QWidget()
        attach_layout = QVBoxLayout(attach_panel)
        attach_layout.addWidget(QLabel("<h3>Select Attachments</h3>"))
        attach_layout.addWidget(QLabel("Check open PDFs to attach:"))
        
        self.attach_list = QListWidget()
        attach_layout.addWidget(self.attach_list)
        
        btn_refresh = QPushButton("🔄 Refresh List")
        btn_refresh.clicked.connect(self.refresh_attachments)
        attach_layout.addWidget(btn_refresh)
        
        layout.addWidget(attach_panel, stretch=1)
        
        self.refresh_attachments()
        self.load_templates()
    
    def refresh_attachments(self):
        self.attach_list.clear()
        from PySide6.QtWidgets import QListWidgetItem # Import locally to avoid NameError
        docks = self.pdf_editor.docks
        for i, dock in enumerate(docks):
            tab_name = dock.windowTitle()
            item = QListWidgetItem(tab_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            item.setData(Qt.UserRole, i) # Store dock index
            self.attach_list.addItem(item)

    def load_templates(self):
        """Load saved mail templates from disk"""
        self.template_combo.clear()
        self.template_combo.addItem("-- Select Template --")
        template_dir = os.path.join(os.getcwd(), "MailTemplates")
        if os.path.exists(template_dir):
            for f in os.listdir(template_dir):
                if f.endswith(".json"):
                    self.template_combo.addItem(f.replace(".json", ""))

    def save_template(self):
        """Save current form as a template"""
        import json
        name, ok = QInputDialog.getText(self, "Save Template", "Template Name:")
        if ok and name:
            template_dir = os.path.join(os.getcwd(), "MailTemplates")
            os.makedirs(template_dir, exist_ok=True)
            data = {
                "from": self.from_input.text(),
                "to": self.to_input.text(),
                "cc": self.cc_input.text(),
                "subject": self.subject_input.text(),
                "body": self.body_input.toPlainText()
            }
            with open(os.path.join(template_dir, f"{name}.json"), "w") as f:
                json.dump(data, f)
            self.load_templates()
            QMessageBox.information(self, "Success", f"Template '{name}' saved!")

    def load_template(self, index):
        """Load a template into the form"""
        import json
        if index <= 0: return
        template_name = self.template_combo.currentText()
        template_path = os.path.join(os.getcwd(), "MailTemplates", f"{template_name}.json")
        if os.path.exists(template_path):
            with open(template_path, "r") as f:
                data = json.load(f)
            self.from_input.setText(data.get("from", ""))
            self.to_input.setText(data.get("to", ""))
            self.cc_input.setText(data.get("cc", ""))
            self.subject_input.setText(data.get("subject", ""))
            self.body_input.setPlainText(data.get("body", ""))
    
    def generate_draft(self):
        try:
            import win32com.client
            import datetime
            
            subject = self.subject_input.text().strip()
            if not subject:
                QMessageBox.warning(self, "Warning", "Subject is required")
                return
            
            # 1. Create Folder Structure
            today = datetime.date.today().strftime("%Y-%m-%d")
            safe_subject = "".join([c for c in subject if c.isalnum() or c in (' ', '-', '_')]).strip()
            folder_path = os.path.join(os.getcwd(), "MailDrafts", today, safe_subject)
            os.makedirs(folder_path, exist_ok=True)
            
            # 2. Save Attachments
            attachments = []
            docks = self.pdf_editor.docks
            for i in range(self.attach_list.count()):
                item = self.attach_list.item(i)
                if item.checkState() == Qt.Checked:
                    dock_idx = item.data(Qt.UserRole)
                    if 0 <= dock_idx < len(docks):
                        dock = docks[dock_idx]
                        tab = dock.widget()
                        if tab and tab.doc:
                            filename = dock.windowTitle()
                            if not filename.lower().endswith(".pdf"):
                                filename += ".pdf"
                            save_path = os.path.join(folder_path, filename)
                            tab.doc.save(save_path)
                            attachments.append(save_path)
            
            # 3. Create Outlook Item
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0) # 0 = olMailItem
            
            mail.Display() # Required to load signature
            signature = mail.HTMLBody
            
            mail.To = self.to_input.text()
            mail.Subject = subject
            
            # CC recipients
            cc_text = self.cc_input.text().strip()
            if cc_text:
                mail.CC = cc_text
            
            # Send on Behalf requires the account to have permissions
            from_addr = self.from_input.text().strip()
            if from_addr:
                try:
                    mail.SentOnBehalfOfName = from_addr
                except Exception as e:
                    print(f"Could not set SentOnBehalfOfName: {e}")
            
            # Preserve signature by appending to body
            user_body = self.body_input.toPlainText().replace("\n", "<br>")
            mail.HTMLBody = f"<p>{user_body}</p><br>" + signature
            
            # Add Attachments
            for path in attachments:
                mail.Attachments.Add(path)
            
            # 4. Save Draft to Folder
            draft_path = os.path.join(folder_path, "Draft.msg")
            mail.SaveAs(draft_path)
            
            # 5. Save to Outlook Drafts
            mail.Save()
            
            QMessageBox.information(self, "Success", f"Draft generated!\nSaved to: {folder_path}")
            
        except ImportError:
            QMessageBox.critical(self, "Error", "pywin32 not installed. Please run: pip install pywin32")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create draft: {e}")

