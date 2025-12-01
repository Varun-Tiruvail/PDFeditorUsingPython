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
                               QGraphicsRectItem, QTabWidget)
from PySide6.QtCore import Qt, QPointF, QRectF, Signal, QThread
from PySide6.QtGui import QPixmap, QImage, QPen, QColor, QBrush
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
# PDF EDITOR MODULE
# ============================================================================

class PDFTab(QWidget):
    def __init__(self, doc, path=None):
        super().__init__()
        self.doc = doc
        self.path = path
        self.current_page = 0
        self.scale = 1.5
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Navigation Toolbar
        nav_layout = QHBoxLayout()
        nav_layout.setAlignment(Qt.AlignCenter)
        
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
        
        layout.addLayout(nav_layout)
        
        # Scroll Area
        self.scroll = QScrollArea()
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)
        self.scroll.setWidget(self.label)
        self.scroll.setWidgetResizable(True)
        layout.addWidget(self.scroll)
        
        self.render()

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

class PDFEditorModule(QWidget):
    def __init__(self):
        super().__init__()
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
        
        self.btn_open = self.create_btn("📂 Open", self.open_pdf)
        self.btn_save = self.create_btn("💾 Save", self.save_pdf)
        self.btn_compress = self.create_btn("🗜️ Compress", self.compress_pdf)
        self.btn_merge = self.create_btn("📑 Merge", self.merge_pdfs)
        self.btn_split = self.create_btn("✂️ Split", self.split_pdf)
        self.btn_redact = self.create_btn("🚫 Redact Page #", self.redact_page_numbers)
        self.btn_pagenum = self.create_btn("🔢 Add Page #", self.add_page_numbers)
        self.btn_header = self.create_btn("📝 Header/Footer", self.add_header_footer)
        
        for btn in [self.btn_open, self.btn_save, self.btn_compress, self.btn_merge, self.btn_split, 
                   self.btn_redact, self.btn_pagenum, self.btn_header]:
            toolbar.addWidget(btn)
        toolbar.addStretch()
        layout.addLayout(toolbar)
        
        # Tabs for Multiple PDFs
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        layout.addWidget(self.tabs)

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
        return self.tabs.currentWidget()

    def close_tab(self, index):
        widget = self.tabs.widget(index)
        if widget:
            widget.deleteLater()
        self.tabs.removeTab(index)

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if path:
            try:
                doc = fitz.open(path)
                tab = PDFTab(doc, path)
                self.tabs.addTab(tab, os.path.basename(path))
                self.tabs.setCurrentWidget(tab)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open PDF: {e}")

    def save_pdf(self):
        tab = self.current_tab()
        if not tab: return
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
        if path:
            try:
                tab.doc.save(path)
                QMessageBox.information(self, "Success", "PDF saved successfully!")
                # Update tab name
                idx = self.tabs.indexOf(tab)
                self.tabs.setTabText(idx, os.path.basename(path))
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

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
                self.tabs.addTab(new_tab, os.path.basename(path))
                self.tabs.setCurrentWidget(new_tab)
                QMessageBox.information(self, "Success", "Compressed PDF opened in new tab!")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def merge_pdfs(self):
        # Custom dialog for reordering
        dialog = QDialog(self)
        dialog.setWindowTitle("Merge PDFs - Drag to Reorder")
        dialog.resize(500, 400)
        layout = QVBoxLayout(dialog)
        
        list_widget = QListWidget()
        list_widget.setDragDropMode(QListWidget.InternalMove)
        layout.addWidget(list_widget)
        
        btn_add = QPushButton("Add PDFs")
        def add_files():
            files, _ = QFileDialog.getOpenFileNames(self, "Select PDFs", "", "PDF Files (*.pdf)")
            for f in files:
                list_widget.addItem(f)
        btn_add.clicked.connect(add_files)
        layout.addWidget(btn_add)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        if dialog.exec() == QDialog.Accepted and list_widget.count() > 0:
            try:
                merged = fitz.open()
                for i in range(list_widget.count()):
                    f = list_widget.item(i).text()
                    merged.insert_pdf(fitz.open(f))
                
                # Open merged in new tab (in-memory)
                tab = PDFTab(merged, "Merged.pdf")
                self.tabs.addTab(tab, "Merged.pdf")
                self.tabs.setCurrentWidget(tab)
                QMessageBox.information(self, "Success", "Merged PDF opened in new tab! Click Save to keep it.")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def split_pdf(self):
        tab = self.current_tab()
        if not tab: return
        
        # For simplicity, just split first 3 pages as a demo, or ask user?
        # Let's just do a simple split 1-3 for now as per previous logic, but open in new tab
        try:
            new_doc = fitz.open()
            new_doc.insert_pdf(tab.doc, from_page=0, to_page=min(2, len(tab.doc)-1))
            
            new_tab = PDFTab(new_doc, "Split.pdf")
            self.tabs.addTab(new_tab, "Split.pdf")
            self.tabs.setCurrentWidget(new_tab)
            QMessageBox.information(self, "Success", "Split PDF (Pages 1-3) opened in new tab!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def redact_page_numbers(self):
        tab = self.current_tab()
        if not tab: return
        
        try:
            doc = tab.doc
            count = 0
            patterns = [
                r"^\d+$", r"^Page\s+\d+$", r"^\d+\s+of\s+\d+$", r"^Page\s+\d+\s+of\s+\d+$"
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
        
        layout.addWidget(QLabel("Exclude Pages (e.g. 1, 3-5):"))
        exclude_input = QLineEdit()
        layout.addWidget(exclude_input)
        
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
        
        if dialog.exec() == QDialog.Accepted:
            try:
                doc = tab.doc
                exclude_str = exclude_input.text().strip()
                excluded = set()
                if exclude_str:
                    for part in exclude_str.split(','):
                        if '-' in part:
                            start, end = map(int, part.split('-'))
                            excluded.update(range(start, end + 1))
                        else:
                            excluded.add(int(part))
                
                total = len(doc)
                fmt = fmt_combo.currentText()
                font_size = size_spin.value()
                
                for i, page in enumerate(doc):
                    pg_num = i + 1
                    if pg_num in excluded: continue
                    
                    if fmt == "n":
                        text = f"{pg_num}"
                    else:
                        text = f"Page {pg_num} of {total}"
                        
                    rect = page.rect
                    pos_idx = pos_combo.currentIndex()
                    
                    if pos_idx == 0: pt = fitz.Point(rect.width/2 - 30, rect.height - 20)
                    elif pos_idx == 1: pt = fitz.Point(rect.width - 80, rect.height - 20)
                    elif pos_idx == 2: pt = fitz.Point(20, rect.height - 20)
                    elif pos_idx == 3: pt = fitz.Point(rect.width/2 - 30, 30)
                    else: pt = fitz.Point(rect.width - 80, 30)
                        
                    page.insert_text(pt, text, fontsize=font_size, color=(0, 0, 0))
                
                tab.render()
                QMessageBox.information(self, "Success", "Page numbers added! Preview updated.")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def add_header_footer(self):
        tab = self.current_tab()
        if not tab: return
            
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Header/Footer")
        layout = QVBoxLayout(dialog)
        
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
                    
                    page.insert_text(fitz.Point(x, y), text, fontsize=size, color=color)
                
                tab.render()
                QMessageBox.information(self, "Success", "Header/Footer added! Preview updated.")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

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
        interval_spin.setValue(60)
        interval_layout.addWidget(interval_spin)
        interval_unit = QComboBox()
        interval_unit.addItems(["Seconds", "Minutes", "Hours"])
        interval_unit.setCurrentText("Minutes")
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
        recurring_layout.addWidget(monthly_widget)
        
        # Show/hide based on recurrence type
        def update_recurrence_widgets():
            rec_type = recurrence_combo.currentText()
            interval_widget.setVisible(rec_type == "Interval")
            time_widget.setVisible(rec_type in ["Daily", "Weekly", "Monthly"])
            weekly_widget.setVisible(rec_type == "Weekly")
            monthly_widget.setVisible(rec_type == "Monthly")
        
        recurrence_combo.currentTextChanged.connect(update_recurrence_widgets)
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

