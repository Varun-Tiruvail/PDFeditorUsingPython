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
                               QRubberBand, QMenu)
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
        self.btn_redact = self.create_btn("🚫 Redact Auto", self.redact_page_numbers)
        self.btn_redact_custom = self.create_btn("🎯 Redact Custom", self.redact_custom_location)
        self.btn_pagenum = self.create_btn("🔢 Add Page #", self.add_page_numbers)
        self.btn_header = self.create_btn("📝 Header/Footer", self.add_header_footer)
        self.btn_advanced = self.create_btn("🔧 Advanced Tools", self.show_advanced_menu)
        
        for btn in [self.btn_open, self.btn_save, self.btn_close_all, self.btn_ppt, self.btn_compress, self.btn_merge, self.btn_split, 
                   self.btn_redact, self.btn_redact_custom, self.btn_pagenum, self.btn_header, self.btn_advanced]:
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

    def redact_page_numbers(self):
        tab = self.current_tab()
        if not tab: return
        
        try:
            doc = tab.doc
            count = 0
            # Enhanced patterns to catch more page number formats including semi-bold
            patterns = [
                r"^\d+$",                          # Just number: 1, 2, 3
                r"^Page\s*\d+$",                   # Page 1, Page1
                r"^\d+\s*of\s*\d+$",               # 1 of 10
                r"^Page\s*\d+\s*of\s*\d+$",        # Page 1 of 10
                r"^-\s*\d+\s*-$",                  # - 1 -
                r"^\[\d+\]$",                      # [1]
                r"^\(\d+\)$",                      # (1)
                r"^p\.?\s*\d+$",                   # p.1, p 1
            ]
            
            for page in doc:
                rect = page.rect
                w, h = rect.width, rect.height
                
                # Define regions: Bottom Center (middle 33%) and Bottom Right (right 33%)
                # Bottom 10% height
                regions = [
                    fitz.Rect(w * 0.33, h * 0.9, w * 0.66, h), # Bottom Center
                    fitz.Rect(w * 0.66, h * 0.9, w, h)         # Bottom Right
                ]
                
                for region in regions:
                    blocks = page.get_text("dict", clip=region)["blocks"]
                    for b in blocks:
                        for l in b["lines"]:
                            for s in l["spans"]:
                                text = s["text"].strip()
                                for pat in patterns:
                                    if re.match(pat, text, re.IGNORECASE):
                                        page.add_redact_annot(fitz.Rect(s["bbox"]), fill=(1, 1, 1))
                                        count += 1
                                        break
                page.apply_redactions()
            
            tab.render() # Refresh view
            QMessageBox.information(self, "Success", f"Redacted {count} locations in Bottom Center/Right.")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def redact_custom_location(self):
        self.redact_mode = "standard"
        tab = self.current_tab()
        if not tab:
            QMessageBox.warning(self, "No PDF", "Please open a PDF first.")
            return
        
        tab.label.set_selection_mode(True)
        tab.label.setCursor(Qt.CrossCursor)
        QMessageBox.information(self, "Custom Redaction", "Draw a box around the area you wish to redact.")

    def prepare_rasterize_redaction(self):
        """Start coordinate selection for rasterization redaction"""
        self.redact_mode = "rasterize"
        tab = self.current_tab()
        if not tab: return
        
        tab.label.set_selection_mode(True)
        tab.label.setCursor(Qt.CrossCursor)
        QMessageBox.information(self, "Select Redaction Area", "Draw a box around the area (e.g. page number) to redact on ALL pages during rasterization.")

    def apply_custom_redaction(self, tab, ui_rect):
        tab.label.set_selection_mode(False)
        tab.label.setCursor(Qt.ArrowCursor)
        
        if ui_rect.width() < 5 or ui_rect.height() < 5:
            return

        pixmap = tab.label.pixmap()
        if not pixmap or pixmap.isNull():
            return

        try:
            # 1. Get Source Page Info
            # Note: pixmap dimensions match the VISUAL size of the page (after rotation) * scale
            p_width = pixmap.width()
            p_height = pixmap.height()
            
            # Map UI coordinates to Pixmap coordinates
            pixmap_rect = pixmap.rect()
            label_rect = tab.label.rect()
            
            # Center offset
            offset_x = (label_rect.width() - pixmap_rect.width()) / 2
            offset_y = (label_rect.height() - pixmap_rect.height()) / 2
            
            # Get Selection coordinates relative to the Pixmap (Visual Page)
            vis_x0 = (ui_rect.left() - offset_x)
            vis_y0 = (ui_rect.top() - offset_y)
            vis_x1 = (ui_rect.right() - offset_x)
            vis_y1 = (ui_rect.bottom() - offset_y)
            
            # Normalize these (0.0 to 1.0) relative to visual page size
            # This makes us independent of zoom (scale) AND independent of absolute page size (if we want relative placement)
            n_x0 = vis_x0 / p_width
            n_y0 = vis_y0 / p_height
            n_x1 = vis_x1 / p_width
            n_y1 = vis_y1 / p_height
            
            # Clamp
            n_x0 = max(0.0, min(n_x0, 1.0))
            n_y0 = max(0.0, min(n_y0, 1.0))
            n_x1 = max(0.0, min(n_x1, 1.0))
            n_y1 = max(0.0, min(n_y1, 1.0))

            # --- BRANCH BASED ON MODE ---
            if getattr(self, "redact_mode", "standard") == "rasterize":
                # For rasterization, we just need to pass the normalized rect or handle it there.
                # Currently rasterize expects exact geometry relative to Bottom/Right.
                # Let's calculate geometry for the current page and pass it.
                # Rasterization will assume all pages become images of this visual orientation.
                # We need to de-normalize for the underlying PDF logic if we use it there, 
                # but rasterizer creates new pages.
                
                # Let's calc visual rect in PDF points (unscaled) for current page
                page = tab.doc.load_page(tab.current_page)
                # Visual dimensions in points:
                if page.rotation in (0, 180):
                    vis_w_pts, vis_h_pts = page.rect.width, page.rect.height
                else:
                    vis_w_pts, vis_h_pts = page.rect.height, page.rect.width
                
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

            # Standard Mode
            reply = QMessageBox.question(self, "Confirm Redaction", 
                                       "Redact this area on all pages?",
                                       QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            
            if reply == QMessageBox.Cancel: return
            
            pages_to_process = range(len(tab.doc)) if reply == QMessageBox.Yes else [tab.current_page]
            
            for pg_idx in pages_to_process:
                pg = tab.doc.load_page(pg_idx)
                rot = pg.rotation
                
                # 1. Determine Visual Dimensions of this page in Points
                # Internal (Unrotated) Dimensions
                w_int = pg.rect.width
                h_int = pg.rect.height
                
                if rot in (90, 270):
                    w_vis = h_int
                    h_vis = w_int
                else:
                    w_vis = w_int
                    h_vis = h_int
                
                # 2. Map Normalized Coordinates to Visual Points for this page
                # This ensures "Relative Visual Position" is preserved
                vx0 = n_x0 * w_vis
                vy0 = n_y0 * h_vis
                vx1 = n_x1 * w_vis
                vy1 = n_y1 * h_vis
                
                # 3. Transform Visual Rect (vx0, vy0, vx1, vy1) to Internal Rect (ix0, iy0, ix1, iy1)
                # Apply Inverse Rotation Logic
                
                if rot == 0:
                    rect = fitz.Rect(vx0, vy0, vx1, vy1)
                    
                elif rot == 90:
                    # Vis x -> Int y (start from top?)
                    # x_int = y_vis
                    # y_int = w_vis - x_vis - (width of rect? no, x_vis is right edge?)
                    # Let's map corners:
                    # TL_vis (vx0, vy0) -> (vy0, w_vis - vx0) ? No.
                    # 90 deg CW: Top Edge -> Right Edge.
                    # Vis (x, 0) -> Int (H_int, x) ?? No.
                    # Let's use the verified logic:
                    # x_int = y_vis
                    # y_int = h_int - x_vis  (Note: h_int == w_vis)
                    
                    # We have a Rect defined by 2 points. We must map both points (TL and BR) 
                    # Use points to avoid confusion with min/max
                    p1 = (vx0, vy0)
                    p2 = (vx1, vy1)
                    
                    # Transform Function for 90 deg (Counter-Clockwise relative to content? No page rotation is CW)
                    # Content at (x,y) appears at rot(x,y).
                    # We perceive (vx, vy). We want (ix, iy).
                    # ix = vy
                    # iy = w_vis - vx (which is h_int - vx)
                    
                    ix0, iy0 = vy0, w_vis - vx0
                    ix1, iy1 = vy1, w_vis - vx1
                    rect = fitz.Rect(ix0, iy0, ix1, iy1).normalize()
                    
                elif rot == 180:
                    # ix = w_vis - vx  (w_vis == w_int)
                    # iy = h_vis - vy  (h_vis == h_int)
                    ix0, iy0 = w_vis - vx0, h_vis - vy0
                    ix1, iy1 = w_vis - vx1, h_vis - vy1
                    rect = fitz.Rect(ix0, iy0, ix1, iy1).normalize()
                    
                elif rot == 270:
                    # 270 CW. Top Edge -> Left Edge.
                    # Vis (0, y) -> Int (y, 0) ? No.
                    # Logic:
                    # ix = h_vis - vy  (h_vis == w_int)
                    # iy = vx
                    ix0, iy0 = h_vis - vy0, vx0
                    ix1, iy1 = h_vis - vy1, vx1
                    rect = fitz.Rect(ix0, iy0, ix1, iy1).normalize()
                
                else:
                    # Fallback for odd rotations
                    rect = fitz.Rect(vx0, vy0, vx1, vy1)

                pg.add_redact_annot(rect, fill=(1, 1, 1))
                pg.apply_redactions()
            
            tab.render()
            QMessageBox.information(self, "Success", "Redaction applied.")
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
            
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Page Numbers")
        layout = QVBoxLayout(dialog)
        
        layout.addWidget(QLabel("Format:"))
        fmt_combo = QComboBox()
        fmt_combo.addItems(["Page n of n", "n"])
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
        layout.addWidget(pos_combo)
        
        layout.addWidget(QLabel("Font Size:"))
        size_spin = QSpinBox()
        size_spin.setRange(6, 72)
        size_spin.setValue(10)
        layout.addWidget(size_spin)
        
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
                
                current_seq_num = 1
                for i in range(len(doc)):
                    pg_index = i + 1
                    
                    if pg_index in skipped:
                        continue
                    
                    if pg_index not in omitted:
                        # Load page properly to avoid stale reference
                        page = doc.load_page(i)
                        
                        if fmt == "n":
                            text = f"{current_seq_num}"
                        else:
                            text = f"Page {current_seq_num} of {total_eligible}"
                            
                        rect = page.rect
                        pos_idx = pos_combo.currentIndex()
                        
                        if pos_idx == 0: pt = fitz.Point(rect.width/2 - 30, rect.height - 20)
                        elif pos_idx == 1: pt = fitz.Point(rect.width - 80, rect.height - 20)
                        elif pos_idx == 2: pt = fitz.Point(20, rect.height - 20)
                        elif pos_idx == 3: pt = fitz.Point(rect.width/2 - 30, 30)
                        else: pt = fitz.Point(rect.width - 80, 30)
                            
                        page.insert_text(pt, text, fontname="times-roman", fontsize=font_size, color=(0, 0, 0))
                    
                    current_seq_num += 1
                
                tab.render()
                QMessageBox.information(self, "Success", "Page numbers added! Preview updated.")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def add_header_footer(self):
        tab = self.current_tab()
        if not tab: return
            
        dialog = QDialog(self)
        dialog.setWindowTitle("Add/Remove Header/Footer")
        dialog.resize(450, 350)
        layout = QVBoxLayout(dialog)
        
        # Remove Button at top
        btn_remove = QPushButton("🗑️ Remove All Headers/Footers")
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
        layout.addWidget(type_combo)
        
        layout.addWidget(QLabel("Alignment:"))
        align_combo = QComboBox()
        align_combo.addItems(["Center", "Left", "Right"])
        layout.addWidget(align_combo)
        
        # Font Selection
        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("Font:"))
        font_combo = QComboBox()
        font_combo.addItems([
            "Times New Roman",
            "Times-Roman", 
            "Helvetica",
            "Courier",
            "Arial"
        ])
        font_combo.setCurrentText("Times New Roman")  # Default
        font_layout.addWidget(font_combo)
        layout.addLayout(font_layout)
        
        # Styling
        style_layout = QHBoxLayout()
        
        style_layout.addWidget(QLabel("Size:"))
        size_spin = QSpinBox()
        size_spin.setRange(8, 72)
        size_spin.setValue(12)
        style_layout.addWidget(size_spin)
        
        style_layout.addWidget(QLabel("Color:"))
        color_combo = QComboBox()
        color_combo.addItems(["Black", "Red", "Blue", "Green", "Gray"])
        style_layout.addWidget(color_combo)
        
        layout.addLayout(style_layout)
        
        # Preset Logic
        def load_draft():
            text_input.setText("DRAFT")
            type_combo.setCurrentText("Header")
            align_combo.setCurrentText("Center")
            font_combo.setCurrentText("Times New Roman")
            size_spin.setValue(26)
            color_combo.setCurrentText("Red")
        
        btn_draft.clicked.connect(load_draft)
        
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
                font_name = font_combo.currentText()
                
                # Map to PyMuPDF font names
                font_map = {
                    "Times New Roman": "times-roman",
                    "Times-Roman": "times-roman",
                    "Helvetica": "helv",
                    "Courier": "cour",
                    "Arial": "helv"  # Arial maps to Helvetica
                }
                fontname = font_map.get(font_name, "times-roman")
                
                # Map color names to RGB tuples
                colors = {
                    "black": (0, 0, 0),
                    "red": (1, 0, 0),
                    "blue": (0, 0, 1),
                    "green": (0, 0.5, 0),
                    "gray": (0.5, 0.5, 0.5)
                }
                color = colors.get(color_name, (0, 0, 0))
                
                for page in doc:
                    rect = page.rect
                    y = 30 if is_header else rect.height - 20
                    
                    # Calculate X based on text length (approx)
                    text_width = len(text) * (size * 0.5) 
                    
                    if align == "Center": x = (rect.width - text_width) / 2
                    elif align == "Left": x = 20
                    else: x = rect.width - 20 - text_width
                    
                    page.insert_text(fitz.Point(x, y), text, fontname=fontname, fontsize=size, color=color)
                
                tab.render()
                QMessageBox.information(self, "Success", "Header/Footer added! Preview updated.")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
    
    def remove_header_footer(self, tab, parent_dialog):
        """Remove header/footer text matching common patterns (page numbers, dates, etc.)"""
        try:
            doc = tab.doc
            removed_count = 0
            
            # Patterns that identify header/footer content
            hf_patterns = [
                r"^\d+$",                          # Just number
                r"^Page\s*\d+",                    # Page 1...
                r"^\d+\s*of\s*\d+$",               # 1 of 10
                r"^-\s*\d+\s*-$",                  # - 1 -
                r"^\[\d+\]$",                      # [1]
                r"^\(\d+\)$",                      # (1)
                r"^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$",  # Dates
                r"^(Draft|Confidential|Private)$",  # Common watermarks
            ]
            
            for page in doc:
                rect = page.rect
                # Define header and footer regions (top 60px and bottom 60px)
                header_rect = fitz.Rect(0, 0, rect.width, 60)
                footer_rect = fitz.Rect(0, rect.height - 60, rect.width, rect.height)
                
                for region in [header_rect, footer_rect]:
                    blocks = page.get_text("dict", clip=region)["blocks"]
                    for block in blocks:
                        if "lines" in block:
                            for line in block["lines"]:
                                for span in line["spans"]:
                                    text = span["text"].strip()
                                    # Only remove if it matches a header/footer pattern
                                    for pat in hf_patterns:
                                        if re.match(pat, text, re.IGNORECASE):
                                            bbox = fitz.Rect(span["bbox"])
                                            page.add_redact_annot(bbox, fill=(1, 1, 1))
                                            removed_count += 1
                                            break
                page.apply_redactions()
            
            tab.render()
            parent_dialog.accept()
            QMessageBox.information(self, "Success", f"Removed {removed_count} header/footer items!")
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
# OCR TRAINER MODULE
# ============================================================================

class BoundingBox:
    def __init__(self, rect, name):
        self.rect = rect  # QRectF
        self.name = name

class OCRTrainerModule(QWidget):
    def __init__(self):
        super().__init__()
        self.current_pdf = None
        self.current_image = None
        self.boxes = []
        self.setup_ui()
    
    def setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Left Panel
        left_panel = QVBoxLayout()
        
        title = QLabel("🔍 OCR Trainer")
        title.setObjectName("moduleTitle")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        left_panel.addWidget(title)
        
        btn_upload = QPushButton("📤 Upload PDF")
        btn_upload.clicked.connect(self.upload_sample)
        left_panel.addWidget(btn_upload)
        
        self.template_name = QLineEdit()
        self.template_name.setPlaceholderText("Template Name")
        left_panel.addWidget(self.template_name)
        
        btn_save = QPushButton("💾 Save Template")
        btn_save.clicked.connect(self.save_template)
        left_panel.addWidget(btn_save)
        
        lbl = QLabel("📥 Extract:")
        left_panel.addWidget(lbl)
        
        self.template_combo = QComboBox()
        self.load_templates()
        left_panel.addWidget(self.template_combo)
        
        btn_extract = QPushButton("▶️ Run Extraction")
        btn_extract.clicked.connect(self.run_extraction)
        left_panel.addWidget(btn_extract)
        
        self.result_table = QTableWidget(0, 2)
        self.result_table.setHorizontalHeaderLabels(["Field", "Value"])
        left_panel.addWidget(self.result_table)
        
        btn_export = QPushButton("📊 Export to Excel")
        btn_export.clicked.connect(self.export_excel)
        left_panel.addWidget(btn_export)
        
        left_panel.addStretch()
        
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        left_widget.setFixedWidth(280)
        
        layout.addWidget(left_widget)
        
        # Right Panel - Canvas
        self.canvas = CanvasWidget()
        layout.addWidget(self.canvas)
    
    def upload_sample(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if path:
            try:
                doc = fitz.open(path)
                page = doc.load_page(0)
                
                # Store ACTUAL page dimensions (not zoomed)
                self.actual_page_width = page.rect.width
                self.actual_page_height = page.rect.height
                
                # Render at 2x for better display
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                self.current_image = QPixmap.fromImage(img)
                self.canvas.set_image(self.current_image, scale_factor=2.0)
                self.current_pdf = path
                doc.close()
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))
    
    def save_template(self):
        name = self.template_name.text().strip()
        if not name or not self.canvas.boxes:
            QMessageBox.warning(self, "Warning", "Enter name and draw boxes")
            return
        
        session = SessionLocal()
        
        # Check if template name already exists
        existing = session.query(Template).filter(Template.name == name).first()
        if existing:
            reply = QMessageBox.question(self, "Template Exists", 
                                        f"Template '{name}' already exists. Overwrite?",
                                        QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                session.close()
                return
            else:
                # Delete existing template (will cascade delete fields)
                session.delete(existing)
                session.commit()
        
        # Use ACTUAL page dimensions, not zoomed display dimensions
        template = Template(name=name, 
                          base_width=self.actual_page_width, 
                          base_height=self.actual_page_height)
        session.add(template)
        session.commit()
        
        print("=" * 50)
        print(f"SAVING TEMPLATE: {name}")
        print(f"Base dimensions: {self.actual_page_width:.2f} x {self.actual_page_height:.2f}")
        print(f"Scale factor: {self.canvas.scale_factor}")
        print(f"Number of boxes: {len(self.canvas.boxes)}")
        print("-" * 50)
        
        # Scale box coordinates back to original PDF size
        for box in self.canvas.boxes:
            scaled_x = box.rect.x() / self.canvas.scale_factor
            scaled_y = box.rect.y() / self.canvas.scale_factor
            scaled_w = box.rect.width() / self.canvas.scale_factor
            scaled_h = box.rect.height() / self.canvas.scale_factor
            
            print(f"Box: {box.name}")
            print(f"  Display coords: ({box.rect.x():.2f}, {box.rect.y():.2f}, {box.rect.width():.2f}, {box.rect.height():.2f})")
            print(f"  Saved coords: ({scaled_x:.2f}, {scaled_y:.2f}, {scaled_w:.2f}, {scaled_h:.2f})")
            
            field = Field(template_id=template.id, name=box.name,
                        x=scaled_x, y=scaled_y, 
                        width=scaled_w, height=scaled_h)
            session.add(field)
        
        session.commit()
        session.close()
        
        print("=" * 50)
        
        QMessageBox.information(self, "Success", "Template saved!")
        self.load_templates()
    
    def load_templates(self):
        self.template_combo.clear()
        session = SessionLocal()
        templates = session.query(Template).all()
        for t in templates:
            self.template_combo.addItem(t.name, t.id)
        session.close()
    
    def run_extraction(self):
        if self.template_combo.count() == 0:
            return
        
        path, _ = QFileDialog.getOpenFileName(self, "Select PDF to Extract", "", "PDF Files (*.pdf)")
        if not path:
            return
        
        template_id = self.template_combo.currentData()
        session = SessionLocal()
        template = session.query(Template).filter(Template.id == template_id).first()
        
        try:
            doc = fitz.open(path)
            page = doc.load_page(0)
            page_rect = page.rect
            
            # Print debug info
            print("=" * 50)
            print(f"EXTRACTION DEBUG")
            print(f"Template: {template.name}")
            print(f"Template base dimensions: {template.base_width:.2f} x {template.base_height:.2f}")
            print(f"PDF page dimensions: {page_rect.width:.2f} x {page_rect.height:.2f}")
            
            scale_x = page_rect.width / template.base_width
            scale_y = page_rect.height / template.base_height
            
            print(f"Scale factors: X={scale_x:.4f}, Y={scale_y:.4f}")
            print(f"Number of fields: {len(template.fields)}")
            print("-" * 50)
            
            self.result_table.setRowCount(len(template.fields))
            
            for i, field in enumerate(template.fields):
                # Calculate scaled coordinates
                x0 = field.x * scale_x
                y0 = field.y * scale_y
                x1 = (field.x + field.width) * scale_x
                y1 = (field.y + field.height) * scale_y
                
                # Add small padding (2px) to handle minor shifts
                padding = 2
                rect = fitz.Rect(x0 - padding, y0 - padding, x1 + padding, y1 + padding)
                
                print(f"Field: {field.name}")
                print(f"  Stored coords: ({field.x:.2f}, {field.y:.2f}, {field.width:.2f}, {field.height:.2f})")
                print(f"  Scaled rect (w/ padding): ({rect.x0:.2f}, {rect.y0:.2f}) -> ({rect.x1:.2f}, {rect.y1:.2f})")
                
                # Try to extract text
                text = page.get_text("text", clip=rect).strip()
                
                # If that doesn't work, try textbox method
                if not text:
                    text = page.get_textbox(rect).strip()
                
                print(f"  Raw extracted: '{text}'")
                
                # SMART EXTRACTION:
                # If the text starts with the field name (e.g. Field="Name", Text="Name: Varun"),
                # strip the field name to get just the value.
                import re
                # Pattern: Start of string, Field Name (case insensitive), optional colon/hyphen, whitespace
                pattern = f"^{re.escape(field.name)}[:\\-\\s]*"
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    cleaned_text = re.sub(pattern, "", text, count=1, flags=re.IGNORECASE).strip()
                    if cleaned_text:
                        print(f"  Smart Cleaned: '{text}' -> '{cleaned_text}'")
                        text = cleaned_text
                
                print(f"  Final Value: '{text}'")
                print()
                
                self.result_table.setItem(i, 0, QTableWidgetItem(field.name))
                self.result_table.setItem(i, 1, QTableWidgetItem(text))
            
            # Create a visual preview with rectangles drawn
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            preview_pixmap = QPixmap.fromImage(img)
            
            # Draw extraction rectangles on the preview using QPainter
            from PySide6.QtGui import QPainter
            painter = QPainter(preview_pixmap)
            pen = QPen(QColor(255, 0, 0), 3)
            painter.setPen(pen)
            
            for field in template.fields:
                x0 = field.x * scale_x * 2
                y0 = field.y * scale_y * 2
                w = field.width * scale_x * 2
                h = field.height * scale_y * 2
                painter.drawRect(QRectF(x0, y0, w, h))
            
            painter.end()
            
            # Create a simple preview window
            preview = QLabel()
            preview.setPixmap(preview_pixmap)
            preview.setWindowTitle("Extraction Preview (Red boxes show extraction areas)")
            preview.show()
            preview.setStyleSheet("background: black;")
            
            # Store reference to keep window alive
            self.preview_window = preview
            
            doc.close()
            
            print("=" * 50)
            QMessageBox.information(self, "Success", f"Extracted {len(template.fields)} fields!\nCheck the preview window to see extraction areas.")
            
        except Exception as e:
            print(f"ERROR: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error", str(e))
        finally:
            session.close()
    
    def export_excel(self):
        if self.result_table.rowCount() == 0:
            return
        
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel Files (*.xlsx)")
        if path:
            data = []
            for i in range(self.result_table.rowCount()):
                data.append([
                    self.result_table.item(i, 0).text(),
                    self.result_table.item(i, 1).text()
                ])
            df = pd.DataFrame(data, columns=["Field", "Value"])
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Success", "Exported to Excel!")

class CanvasWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.pixmap = None
        self.boxes = []
        self.start_point = None
        self.current_rect = None
        self.scale_factor = 1.0
        self.setMinimumSize(400, 400)
    
    def set_image(self, pixmap, scale_factor=1.0):
        self.pixmap = pixmap
        self.boxes = []
        self.scale_factor = scale_factor
        self.setFixedSize(pixmap.size())
        self.update()
    
    def paintEvent(self, event):
        from PySide6.QtGui import QPainter
        if not self.pixmap:
            return
        
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self.pixmap)
        
        pen = QPen(QColor(255, 0, 0), 2)
        painter.setPen(pen)
        
        for box in self.boxes:
            painter.drawRect(box.rect.toRect())
            painter.drawText(box.rect.topLeft().toPoint(), box.name)
        
        if self.current_rect:
            pen.setColor(QColor(0, 0, 255))
            painter.setPen(pen)
            painter.drawRect(self.current_rect.toRect())
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.pixmap:
            self.start_point = event.position()
    
    def mouseMoveEvent(self, event):
        if self.start_point:
            self.current_rect = QRectF(self.start_point, event.position()).normalized()
            self.update()
    
    def mouseReleaseEvent(self, event):
        if self.current_rect:
            from PySide6.QtWidgets import QInputDialog
            name, ok = QInputDialog.getText(self, "Field Name", "Enter field name:")
            if ok and name:
                self.boxes.append(BoundingBox(self.current_rect, name))
            self.current_rect = None
            self.start_point = None
            self.update()

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

