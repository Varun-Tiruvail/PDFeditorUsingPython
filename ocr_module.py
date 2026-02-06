"""
OCR Image Trainer Module - For Scanned PDFs/Images
Uses EasyOCR (pure Python) instead of native PDF text extraction.
Supports: Multi-anchor, Multi-PDF templates, Extraction backup PDF
NO EXTERNAL EXECUTABLES REQUIRED - EasyOCR uses PyTorch internally
"""
import os
import fitz  # PyMuPDF
from PIL import Image
import io
import numpy as np

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QScrollArea, QListWidget, QLineEdit, QMessageBox, QFileDialog,
    QComboBox, QTableWidget, QTableWidgetItem, QInputDialog, QSplitter
)
from PySide6.QtCore import Qt, Signal, QRectF
from PySide6.QtGui import QImage, QPixmap, QPainter, QPen, QBrush, QColor

# Import shared components from modules
from modules import (
    OCRBox, OCRCanvasWidget, SessionLocal, 
    OCRTemplate, OCRPage, LabeledBox
)
import pandas as pd

# Try to import EasyOCR (pure Python, no external executable needed)
try:
    import easyocr
    EASYOCR_AVAILABLE = True
    # Initialize reader once (lazy loading to improve startup time)
    _ocr_reader = None
except ImportError:
    EASYOCR_AVAILABLE = False
    _ocr_reader = None
    print("[WARNING] easyocr not installed. Run: pip install easyocr")


# Position tolerance for multi-anchor matching (in pixels at 300 DPI)
POSITION_TOLERANCE = 50

# Default OCR DPI
OCR_DPI = 300


def get_ocr_reader():
    """Get or create EasyOCR reader (singleton pattern for efficiency)"""
    global _ocr_reader
    if _ocr_reader is None and EASYOCR_AVAILABLE:
        print("[INFO] Initializing EasyOCR... (first run downloads models)")
        _ocr_reader = easyocr.Reader(['en'], gpu=False)  # CPU mode for compatibility
    return _ocr_reader


class OCRImageTrainerModule(QWidget):
    """OCR Trainer for scanned images/PDFs using Tesseract OCR.
    
    Key difference from regular OCR Trainer:
    - Converts PDF pages to images first
    - Uses Tesseract OCR to extract text with bounding boxes
    - Stores coordinates in image space (DPI-normalized)
    """
    
    def __init__(self):
        super().__init__()
        
        # PDF management
        self.loaded_pdfs = []  # List of (filename, fitz.Document, path)
        self.current_pdf_index = -1
        self.current_page_index = 0
        
        # Page images (cached PIL images at OCR DPI)
        self.page_images = {}  # (pdf_idx, page_idx) -> PIL Image
        self.page_ocr_results = {}  # (pdf_idx, page_idx) -> OCR results
        
        # Box management
        self.page_boxes = {}  # (pdf_idx, page_idx) -> [OCRBox, ...]
        self.page_dimensions = {}  # (pdf_idx, page_idx) -> (width, height)
        
        # Zoom
        self.zoom_scale = 1.0
        
        # Extraction results
        self.extraction_results = []
        self.extraction_screenshots = []
        
        self.setup_ui()
        self.load_template_list()
    
    def setup_ui(self):
        """Setup the user interface"""
        main_layout = QHBoxLayout(self)
        
        # =================== LEFT PANEL (Controls) ===================
        left_panel = QVBoxLayout()
        
        # --- OCR Status ---
        if EASYOCR_AVAILABLE:
            status_label = QLabel("✅ EasyOCR Ready (No External Install)")
            status_label.setStyleSheet("color: green; font-weight: bold;")
        else:
            status_label = QLabel("❌ EasyOCR Not Found - pip install easyocr")
            status_label.setStyleSheet("color: red; font-weight: bold;")
        left_panel.addWidget(status_label)
        
        # --- PDF Management ---
        pdf_section = QLabel("📄 PDFs (Scanned)")
        pdf_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(pdf_section)
        
        self.pdf_list = QListWidget()
        self.pdf_list.setMaximumHeight(100)
        self.pdf_list.currentRowChanged.connect(self.on_pdf_selected)
        left_panel.addWidget(self.pdf_list)
        
        pdf_buttons = QHBoxLayout()
        self.btn_add_pdf = QPushButton("➕ Add PDFs")
        self.btn_add_pdf.clicked.connect(self.add_pdfs)
        self.btn_remove_pdf = QPushButton("➖ Remove")
        self.btn_remove_pdf.clicked.connect(self.remove_pdf)
        pdf_buttons.addWidget(self.btn_add_pdf)
        pdf_buttons.addWidget(self.btn_remove_pdf)
        left_panel.addLayout(pdf_buttons)
        
        # --- Page Navigation ---
        nav_section = QLabel("📑 Navigation")
        nav_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(nav_section)
        
        nav_layout = QHBoxLayout()
        self.btn_prev = QPushButton("◀ Prev")
        self.btn_prev.clicked.connect(lambda: self.navigate_page(-1))
        self.lbl_page = QLabel("Page 0/0")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        self.btn_next = QPushButton("Next ▶")
        self.btn_next.clicked.connect(lambda: self.navigate_page(1))
        nav_layout.addWidget(self.btn_prev)
        nav_layout.addWidget(self.lbl_page)
        nav_layout.addWidget(self.btn_next)
        left_panel.addLayout(nav_layout)
        
        # --- OCR Controls ---
        ocr_section = QLabel("🔍 OCR")
        ocr_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(ocr_section)
        
        self.btn_run_ocr = QPushButton("🔤 Run OCR on Page")
        self.btn_run_ocr.clicked.connect(self.run_ocr_on_page)
        self.btn_run_ocr.setStyleSheet("background: #FF5722; color: white; padding: 8px;")
        left_panel.addWidget(self.btn_run_ocr)
        
        self.ocr_status_label = QLabel("OCR: Not run")
        left_panel.addWidget(self.ocr_status_label)
        
        # --- Zoom Controls ---
        zoom_layout = QHBoxLayout()
        self.btn_zoom_out = QPushButton("🔍-")
        self.btn_zoom_out.clicked.connect(self.zoom_out)
        self.zoom_label = QLabel("100%")
        self.zoom_label.setAlignment(Qt.AlignCenter)
        self.btn_zoom_in = QPushButton("🔍+")
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        self.btn_zoom_fit = QPushButton("Fit")
        self.btn_zoom_fit.clicked.connect(self.zoom_fit)
        zoom_layout.addWidget(self.btn_zoom_out)
        zoom_layout.addWidget(self.zoom_label)
        zoom_layout.addWidget(self.btn_zoom_in)
        zoom_layout.addWidget(self.btn_zoom_fit)
        left_panel.addLayout(zoom_layout)
        
        # --- Box Drawing Mode ---
        mode_section = QLabel("✏️ Drawing Mode")
        mode_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(mode_section)
        
        mode_layout = QHBoxLayout()
        self.btn_mode_label = QPushButton("📦 Label")
        self.btn_mode_label.clicked.connect(lambda: self.set_mode('label'))
        self.btn_mode_label.setStyleSheet("background: #4285F4; color: white;")
        self.btn_mode_anchor = QPushButton("⚓ Anchor")
        self.btn_mode_anchor.clicked.connect(lambda: self.set_mode('anchor'))
        self.btn_mode_anchor.setStyleSheet("background: #34A853; color: white;")
        self.btn_mode_value = QPushButton("💎 Value")
        self.btn_mode_value.clicked.connect(lambda: self.set_mode('value'))
        self.btn_mode_value.setStyleSheet("background: #EA4335; color: white;")
        mode_layout.addWidget(self.btn_mode_label)
        mode_layout.addWidget(self.btn_mode_anchor)
        mode_layout.addWidget(self.btn_mode_value)
        left_panel.addLayout(mode_layout)
        
        # --- Box List ---
        box_section = QLabel("📋 Boxes")
        box_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(box_section)
        
        self.box_list = QListWidget()
        self.box_list.setMaximumHeight(120)
        left_panel.addWidget(self.box_list)
        
        box_buttons = QHBoxLayout()
        self.btn_delete_box = QPushButton("🗑️ Delete")
        self.btn_delete_box.clicked.connect(self.delete_selected_box)
        self.btn_clear_boxes = QPushButton("🧹 Clear All")
        self.btn_clear_boxes.clicked.connect(self.clear_all_boxes)
        box_buttons.addWidget(self.btn_delete_box)
        box_buttons.addWidget(self.btn_clear_boxes)
        left_panel.addLayout(box_buttons)
        
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
        self.btn_test_extract = QPushButton("🧪 Test OCR")
        self.btn_test_extract.clicked.connect(self.test_extract_current)
        self.btn_test_extract.setStyleSheet("background: #FF9800; color: white; padding: 8px;")
        save_layout.addWidget(self.btn_save_template)
        save_layout.addWidget(self.btn_test_extract)
        left_panel.addLayout(save_layout)
        
        # --- Load Template ---
        load_section = QLabel("📥 Load Template")
        load_section.setStyleSheet("font-weight: bold; margin-top: 10px;")
        left_panel.addWidget(load_section)
        
        self.template_combo = QComboBox()
        left_panel.addWidget(self.template_combo)
        
        # Load Template for Editing
        self.btn_load_template = QPushButton("📂 Load Template for Edit")
        self.btn_load_template.clicked.connect(self.load_template_for_editing)
        self.btn_load_template.setStyleSheet("background: #9C27B0; color: white; padding: 8px;")
        left_panel.addWidget(self.btn_load_template)
        
        extract_layout = QHBoxLayout()
        self.btn_run_extraction = QPushButton("▶️ Run OCR Extraction")
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
        
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMinimumWidth(600)
        
        self.canvas = OCRCanvasWidget()
        self.canvas.box_created.connect(self.on_box_created)
        self.canvas.box_selected.connect(self.on_box_selected)
        self.scroll_area.setWidget(self.canvas)
        
        center_layout.addWidget(self.scroll_area)
        main_layout.addLayout(center_layout, stretch=2)
        
        # =================== RIGHT PANEL (Results Table) ===================
        right_panel = QVBoxLayout()
        
        results_label = QLabel("📊 Extraction Results")
        results_label.setStyleSheet("font-weight: bold;")
        right_panel.addWidget(results_label)
        
        self.result_table = QTableWidget()
        self.result_table.setMinimumWidth(400)
        right_panel.addWidget(self.result_table)
        
        right_widget = QWidget()
        right_widget.setLayout(right_panel)
        main_layout.addWidget(right_widget)
    
    # =================== PDF Management ===================
    
    def add_pdfs(self):
        """Add multiple PDF files"""
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF Files", "", "PDF Files (*.pdf)"
        )
        for path in paths:
            try:
                doc = fitz.open(path)
                filename = os.path.basename(path)
                self.loaded_pdfs.append((filename, doc, path))
                self.pdf_list.addItem(filename)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load {path}: {e}")
        
        if self.loaded_pdfs and self.current_pdf_index < 0:
            self.pdf_list.setCurrentRow(0)
    
    def remove_pdf(self):
        """Remove selected PDF"""
        row = self.pdf_list.currentRow()
        if row >= 0:
            filename, doc, path = self.loaded_pdfs[row]
            doc.close()
            del self.loaded_pdfs[row]
            self.pdf_list.takeItem(row)
            
            # Clear cached data for this PDF
            keys_to_remove = [k for k in self.page_images.keys() if k[0] == row]
            for k in keys_to_remove:
                del self.page_images[k]
                if k in self.page_ocr_results:
                    del self.page_ocr_results[k]
                if k in self.page_boxes:
                    del self.page_boxes[k]
            
            if self.loaded_pdfs:
                self.current_pdf_index = min(row, len(self.loaded_pdfs) - 1)
                self.pdf_list.setCurrentRow(self.current_pdf_index)
            else:
                self.current_pdf_index = -1
                self.canvas.set_image(None)
    
    def on_pdf_selected(self, row):
        """Handle PDF selection from list"""
        if row >= 0:
            self.save_current_page_boxes()
            self.current_pdf_index = row
            self.current_page_index = 0
            self.render_current_page()
    
    def navigate_page(self, delta):
        """Navigate between pages"""
        if self.current_pdf_index < 0:
            return
        
        self.save_current_page_boxes()
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        new_page = self.current_page_index + delta
        
        if 0 <= new_page < len(doc):
            self.current_page_index = new_page
            self.render_current_page()
    
    # =================== Page Rendering & OCR ===================
    
    def _page_to_image(self, page, dpi=OCR_DPI):
        """Convert PDF page to PIL Image for OCR"""
        mat = fitz.Matrix(dpi / 72, dpi / 72)  # 72 is base PDF DPI
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img
    
    def run_ocr_on_page(self):
        """Run EasyOCR on current page - no external executable needed"""
        if not EASYOCR_AVAILABLE:
            QMessageBox.warning(self, "OCR Not Available", 
                "EasyOCR is not installed. Run: pip install easyocr")
            return
        
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            return
        
        key = (self.current_pdf_index, self.current_page_index)
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        page = doc.load_page(self.current_page_index)
        
        # Convert to image at OCR DPI
        self.ocr_status_label.setText("OCR: Processing...")
        QMessageBox.information(self, "Running OCR", 
            "Running EasyOCR on page. First run downloads models (~100MB).\nThis may take a moment...")
        
        try:
            pil_image = self._page_to_image(page)
            self.page_images[key] = pil_image
            
            # Convert PIL image to numpy array for EasyOCR
            img_array = np.array(pil_image)
            
            # Get EasyOCR reader and run OCR
            reader = get_ocr_reader()
            if reader is None:
                raise Exception("EasyOCR failed to initialize")
            
            # EasyOCR returns: [(bbox, text, confidence), ...]
            # bbox is [[x1,y1], [x2,y1], [x2,y2], [x1,y2]]
            ocr_results = reader.readtext(img_array)
            
            # Process OCR results into our format
            results = []
            for bbox, text, conf in ocr_results:
                text = text.strip()
                if not text:
                    continue
                
                # Convert bbox to left, top, width, height
                # bbox is [[x1,y1], [x2,y1], [x2,y2], [x1,y2]]
                x_coords = [p[0] for p in bbox]
                y_coords = [p[1] for p in bbox]
                left = min(x_coords)
                top = min(y_coords)
                width = max(x_coords) - left
                height = max(y_coords) - top
                
                results.append({
                    'text': text,
                    'left': left,
                    'top': top,
                    'width': width,
                    'height': height,
                    'conf': int(conf * 100)  # Convert 0-1 to 0-100
                })
            
            self.page_ocr_results[key] = results
            self.ocr_status_label.setText(f"OCR: Found {len(results)} text regions")
            
        except Exception as e:
            self.ocr_status_label.setText(f"OCR: Error - {e}")
            QMessageBox.critical(self, "OCR Error", f"OCR failed: {e}")
    
    def render_current_page(self):
        """Render current page and load its boxes"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            return
        
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        page = doc.load_page(self.current_page_index)
        
        # Update page label
        self.lbl_page.setText(f"Page {self.current_page_index + 1}/{len(doc)}")
        
        key = (self.current_pdf_index, self.current_page_index)
        
        # Store page dimensions at OCR DPI
        dpi_scale = OCR_DPI / 72
        self.page_dimensions[key] = (
            page.rect.width * dpi_scale,
            page.rect.height * dpi_scale
        )
        
        # Render at zoom scale (relative to screen display, not OCR DPI)
        display_mat = fitz.Matrix(self.zoom_scale, self.zoom_scale)
        pix = page.get_pixmap(matrix=display_mat)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(img)
        
        self.canvas.set_image(pixmap, scale_factor=self.zoom_scale)
        
        # Load boxes for this page (scale from OCR coords to display coords)
        if key in self.page_boxes:
            scaled_boxes = []
            for box in self.page_boxes[key]:
                scaled_box = self._scale_box_to_canvas_coords(box)
                scaled_boxes.append(scaled_box)
            self.canvas.set_boxes(scaled_boxes)
        else:
            self.canvas.set_boxes([])
        
        # Update OCR status
        if key in self.page_ocr_results:
            self.ocr_status_label.setText(f"OCR: {len(self.page_ocr_results[key])} words")
        else:
            self.ocr_status_label.setText("OCR: Not run")
        
        self.update_box_list()
        self.update_zoom_label()
    
    # =================== Box Scaling (OCR DPI <-> Display) ===================
    
    def save_current_page_boxes(self):
        """Save boxes from canvas to page_boxes dict in OCR coordinates"""
        if self.current_pdf_index >= 0:
            key = (self.current_pdf_index, self.current_page_index)
            # Convert canvas boxes to OCR coordinates
            ocr_boxes = []
            for box in self.canvas.boxes:
                ocr_box = self._scale_box_to_ocr_coords(box)
                ocr_boxes.append(ocr_box)
            self.page_boxes[key] = ocr_boxes
    
    def _scale_box_to_ocr_coords(self, box):
        """Convert box from canvas coords (display scale) to OCR coords"""
        # Canvas is at zoom_scale, OCR is at OCR_DPI/72 scale
        # Convert: canvas_coord / zoom_scale * (OCR_DPI / 72)
        dpi_scale = OCR_DPI / 72
        conversion = dpi_scale / self.zoom_scale
        
        ocr_rect = QRectF(
            box.rect.x() * conversion,
            box.rect.y() * conversion,
            box.rect.width() * conversion,
            box.rect.height() * conversion
        )
        ocr_box = OCRBox(ocr_rect, box.name, box.box_type, box.parent)
        ocr_box.id = box.id
        
        # Recursively convert children
        for child in box.children:
            ocr_child = self._scale_box_to_ocr_coords(child)
            ocr_child.parent = ocr_box
            ocr_box.children.append(ocr_child)
        
        return ocr_box
    
    def _scale_box_to_canvas_coords(self, box):
        """Convert box from OCR coords to canvas coords (display scale)"""
        dpi_scale = OCR_DPI / 72
        conversion = self.zoom_scale / dpi_scale
        
        canvas_rect = QRectF(
            box.rect.x() * conversion,
            box.rect.y() * conversion,
            box.rect.width() * conversion,
            box.rect.height() * conversion
        )
        canvas_box = OCRBox(canvas_rect, box.name, box.box_type, box.parent)
        canvas_box.id = box.id
        
        # Recursively convert children
        for child in box.children:
            canvas_child = self._scale_box_to_canvas_coords(child)
            canvas_child.parent = canvas_box
            canvas_box.children.append(canvas_child)
        
        return canvas_box
    
    # =================== Zoom Controls ===================
    
    def update_zoom_label(self):
        """Update the zoom percentage label"""
        self.zoom_label.setText(f"{int(self.zoom_scale * 100)}%")
    
    def zoom_fit(self):
        """Fit the page to the scroll area"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            return
        
        scroll_width = self.scroll_area.viewport().width() - 20
        scroll_height = self.scroll_area.viewport().height() - 20
        
        filename, doc, path = self.loaded_pdfs[self.current_pdf_index]
        page = doc.load_page(self.current_page_index)
        
        scale_x = scroll_width / page.rect.width
        scale_y = scroll_height / page.rect.height
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
    
    def apply_zoom(self):
        """Apply the current zoom level to the canvas"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            return
        
        # Save current boxes in OCR coords before zoom changes
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
    
    # =================== Box Management ===================
    
    def set_mode(self, mode):
        """Set drawing mode"""
        self.canvas.set_mode(mode)
    
    def on_box_created(self, box):
        """Handle new box creation"""
        self.update_box_list()
    
    def on_box_selected(self, box):
        """Handle box selection"""
        if box:
            self.canvas.set_active_parent(box)
    
    def update_box_list(self):
        """Update the box list display"""
        self.box_list.clear()
        for box in self.canvas.boxes:
            self.box_list.addItem(f"📦 {box.name}")
            for child in box.children:
                prefix = "  ⚓" if child.box_type == 'anchor' else "  💎"
                self.box_list.addItem(f"{prefix} {child.name}")
    
    def delete_selected_box(self):
        """Delete selected box"""
        self.canvas.delete_selected_box()
        self.update_box_list()
    
    def clear_all_boxes(self):
        """Clear all boxes on current page"""
        reply = QMessageBox.question(self, "Clear All", 
            "Delete all boxes on this page?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.canvas.clear_boxes()
            self.save_current_page_boxes()
            self.update_box_list()
    
    # =================== Template Management ===================
    
    def load_template_list(self):
        """Load template names into combo box"""
        self.template_combo.clear()
        session = SessionLocal()
        templates = session.query(OCRTemplate).all()
        for t in templates:
            self.template_combo.addItem(t.name, t.id)
        session.close()
    
    def save_template(self):
        """Save current template to database"""
        name = self.template_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "Warning", "Please enter a template name.")
            return
        
        # Add OCR suffix if not present
        if not name.endswith("_OCR"):
            name = name + "_OCR"
            self.template_name_input.setText(name)
        
        self.save_current_page_boxes()
        
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
            
            # Store page info (OCR DPI coordinates)
            ocr_page = OCRPage(
                template_id=template.id,
                pdf_filename=filename,
                page_number=page_idx,
                page_width=page_width,
                page_height=page_height,
                page_rotation=0,  # OCR uses image space, no rotation
                order_index=order_idx
            )
            session.add(ocr_page)
            session.commit()
            order_idx += 1
            
            # Save boxes (in OCR coordinates)
            for box in boxes:
                self._save_box_to_db(session, ocr_page.id, box, None)
        
        session.commit()
        session.close()
        
        QMessageBox.information(self, "Success", f"Template '{name}' saved!")
        self.load_template_list()
    
    def _save_box_to_db(self, session, page_id, box, parent_id):
        """Save a box and its children to database"""
        db_box = LabeledBox(
            page_id=page_id,
            parent_box_id=parent_id,
            name=box.name,
            box_type=box.box_type,
            x=box.rect.x(),
            y=box.rect.y(),
            width=box.rect.width(),
            height=box.rect.height()
        )
        session.add(db_box)
        session.commit()
        
        # Save children
        for child in box.children:
            self._save_box_to_db(session, page_id, child, db_box.id)
    
    def load_template_for_editing(self):
        """Load an existing template's boxes onto the current PDF page for editing"""
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
        
        for ocr_page in template.pages:
            label_boxes = session.query(LabeledBox).filter(
                LabeledBox.page_id == ocr_page.id,
                LabeledBox.box_type == 'label'
            ).all()
            
            for db_label in label_boxes:
                label_box = self._db_box_to_ocr_box(db_label)
                
                if current_key not in self.page_boxes:
                    self.page_boxes[current_key] = []
                self.page_boxes[current_key].append(label_box)
                loaded_count += 1
        
        session.close()
        
        self.template_name_input.setText(template_name)
        self.render_current_page()
        self.update_box_list()
        
        QMessageBox.information(self, "Loaded", 
            f"Loaded {loaded_count} label boxes from '{template_name}'.\n"
            f"Add new boxes and click 'Save Template' to update.")
    
    def _db_box_to_ocr_box(self, db_box):
        """Convert a database LabeledBox to an OCRBox"""
        rect = QRectF(db_box.x, db_box.y, db_box.width, db_box.height)
        ocr_box = OCRBox(rect, db_box.name, db_box.box_type)
        ocr_box.id = db_box.id
        
        for child_db in db_box.children:
            child_box = self._db_box_to_ocr_box(child_db)
            child_box.parent = ocr_box
            ocr_box.children.append(child_box)
        
        return ocr_box
    
    # =================== OCR-Based Extraction ===================
    
    def test_extract_current(self):
        """Test OCR extraction on current page's boxes"""
        if self.current_pdf_index < 0 or not self.loaded_pdfs:
            QMessageBox.warning(self, "No PDF", "Please load a PDF first.")
            return
        
        key = (self.current_pdf_index, self.current_page_index)
        
        # Check if OCR has been run
        if key not in self.page_ocr_results:
            QMessageBox.warning(self, "OCR Required", 
                "Please run OCR on this page first (click 'Run OCR on Page').")
            return
        
        self.save_current_page_boxes()
        boxes = self.page_boxes.get(key, [])
        
        if not boxes:
            QMessageBox.warning(self, "No Boxes", "No boxes drawn on this page.")
            return
        
        ocr_results = self.page_ocr_results[key]
        results = []
        
        for box in boxes:
            if box.box_type == 'label':
                anchors = [b for b in box.children if b.box_type == 'anchor']
                values = [b for b in box.children if b.box_type == 'value']
                
                anchor_text = ""
                value_text = ""
                
                # Extract anchor text from OCR results
                for anchor in anchors:
                    text = self._get_text_in_rect(ocr_results, anchor.rect)
                    if text:
                        anchor_text += text + " "
                
                # Extract value text from OCR results
                for value in values:
                    text = self._get_text_in_rect(ocr_results, value.rect)
                    if text:
                        value_text += text + " "
                
                results.append({
                    'label': box.name,
                    'anchor': anchor_text.strip(),
                    'value': value_text.strip()
                })
        
        # Show results
        msg = "OCR Extraction Results:\n\n"
        for r in results:
            msg += f"📦 {r['label']}:\n"
            msg += f"   ⚓ Anchor: {r['anchor']}\n"
            msg += f"   💎 Value: {r['value']}\n\n"
        
        QMessageBox.information(self, "OCR Test Results", msg)
    
    def _get_text_in_rect(self, ocr_results, rect):
        """Get OCR text that falls within a given rectangle (in OCR coordinates)"""
        texts = []
        for item in ocr_results:
            # Check if OCR word overlaps with rect
            word_rect = QRectF(
                item['left'], item['top'],
                item['width'], item['height']
            )
            if rect.intersects(word_rect):
                texts.append(item['text'])
        return ' '.join(texts)
    
    def run_extraction(self):
        """Run OCR extraction on selected PDFs using loaded template"""
        template_id = self.template_combo.currentData()
        if not template_id:
            QMessageBox.warning(self, "No Template", "Please select a template first.")
            return
        
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select PDFs for OCR Extraction", "", "PDF Files (*.pdf)"
        )
        if not paths:
            return
        
        if not TESSERACT_AVAILABLE:
            QMessageBox.warning(self, "OCR Not Available", 
                "Tesseract OCR is not installed.")
            return
        
        session = SessionLocal()
        template = session.query(OCRTemplate).filter(OCRTemplate.id == template_id).first()
        
        if not template:
            QMessageBox.warning(self, "Error", "Template not found.")
            session.close()
            return
        
        # Get all labels and their anchor/value info
        all_labels = []
        label_info = {}
        
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
                        'page_height': ocr_page.page_height
                    }
        
        if not all_labels:
            QMessageBox.warning(self, "No Labels", "Template has no label boxes defined.")
            session.close()
            return
        
        # Setup result table
        self.result_table.clear()
        self.result_table.setRowCount(0)
        
        columns = ["PDF Filename"]
        for label in all_labels:
            columns.append(f"{label['name']}_Anchor")
            columns.append(f"{label['name']}_Value")
        
        self.result_table.setColumnCount(len(columns))
        self.result_table.setHorizontalHeaderLabels(columns)
        
        self.extraction_results = []
        self.extraction_screenshots = []
        extracted_count = 0
        
        try:
            for pdf_path in paths:
                doc = fitz.open(pdf_path)
                pdf_filename = os.path.basename(pdf_path)
                
                row_data = {'PDF Filename': pdf_filename}
                
                for label in all_labels:
                    info = label_info[label['id']]
                    
                    # Find this label's value in the PDF using OCR
                    match = self._find_box_with_ocr(
                        doc, info['anchors'], info['values'],
                        info['page_width'], info['page_height'],
                        pdf_path
                    )
                    
                    if match:
                        row_data[f"{label['name']}_Anchor"] = match['anchor_text']
                        row_data[f"{label['name']}_Value"] = match['value_text']
                        extracted_count += 1
                        
                        # Collect screenshot data
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
            
            self.result_table.resizeColumnsToContents()
            
            # Generate backup PDF
            if self.extraction_screenshots:
                self._generate_extraction_backup()
            
            QMessageBox.information(self, "Complete", 
                f"OCR Extraction complete!\n{len(paths)} PDFs processed\n{extracted_count} values extracted")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"OCR Extraction failed: {e}")
            import traceback
            traceback.print_exc()
        finally:
            session.close()
    
    def _find_box_with_ocr(self, doc, anchors, values, page_width, page_height, pdf_path):
        """Find box content using OCR multi-anchor approach"""
        if not anchors or not values:
            return None
        
        first_anchor = anchors[0]
        first_value = values[0]
        
        # Get primary anchor text from stored box
        primary_anchor_rect = QRectF(first_anchor.x, first_anchor.y, 
                                      first_anchor.width, first_anchor.height)
        
        # Calculate value offset from primary anchor
        value_dx = first_value.x - first_anchor.x
        value_dy = first_value.y - first_anchor.y
        value_w = first_value.width
        value_h = first_value.height
        
        # Build secondary anchors list
        secondary_anchors = []
        for anchor in anchors[1:]:
            secondary_anchors.append({
                'rect': QRectF(anchor.x, anchor.y, anchor.width, anchor.height),
                'expected_dx': anchor.x - first_anchor.x,
                'expected_dy': anchor.y - first_anchor.y
            })
        
        # Search all pages with OCR
        for page_idx in range(len(doc)):
            try:
                page = doc.load_page(page_idx)
                
                # Convert page to image and run OCR with EasyOCR
                pil_image = self._page_to_image(page)
                img_array = np.array(pil_image)
                
                reader = get_ocr_reader()
                if reader is None:
                    continue
                
                ocr_results = reader.readtext(img_array)
                
                # Build word list with positions from EasyOCR results
                words = []
                for bbox, text, conf in ocr_results:
                    text = text.strip()
                    if not text:
                        continue
                    x_coords = [p[0] for p in bbox]
                    y_coords = [p[1] for p in bbox]
                    words.append({
                        'text': text,
                        'left': min(x_coords),
                        'top': min(y_coords),
                        'width': max(x_coords) - min(x_coords),
                        'height': max(y_coords) - min(y_coords)
                    })
                
                # Try to match anchor pattern at primary anchor's relative position
                # For now, use direct rect matching (same position)
                value_rect = QRectF(
                    first_value.x,
                    first_value.y,
                    first_value.width,
                    first_value.height
                )
                
                # Get text in value region
                value_text = self._get_text_in_rect_from_words(words, value_rect)
                
                # Get anchor text for verification
                anchor_text = self._get_text_in_rect_from_words(words, primary_anchor_rect)
                
                if value_text:
                    return {
                        'page': page_idx,
                        'anchor_text': anchor_text,
                        'value_text': value_text,
                        'value_rect': fitz.Rect(
                            value_rect.x() * 72 / OCR_DPI,
                            value_rect.y() * 72 / OCR_DPI,
                            (value_rect.x() + value_rect.width()) * 72 / OCR_DPI,
                            (value_rect.y() + value_rect.height()) * 72 / OCR_DPI
                        )
                    }
                
            except Exception as e:
                print(f"[DEBUG] OCR error on page {page_idx}: {e}")
                continue
        
        return None
    
    def _get_text_in_rect_from_words(self, words, rect):
        """Get text from word list that falls within rect (OCR coordinates)"""
        texts = []
        for word in words:
            word_rect = QRectF(word['left'], word['top'], word['width'], word['height'])
            if rect.intersects(word_rect):
                texts.append(word['text'])
        return ' '.join(texts)
    
    # =================== Export Functions ===================
    
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
            backup_doc = fitz.open()
            
            for (pdf_idx, page_idx), boxes in self.page_boxes.items():
                if not boxes:
                    continue
                
                filename, doc, orig_path = self.loaded_pdfs[pdf_idx]
                page = doc.load_page(page_idx)
                
                mat = fitz.Matrix(2, 2)
                pix = page.get_pixmap(matrix=mat)
                
                for box in boxes:
                    self._add_box_to_backup(backup_doc, pix, box, filename, page_idx, mat)
            
            backup_doc.save(path)
            backup_doc.close()
            
            QMessageBox.information(self, "Success", f"Backup PDF saved to {path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save backup: {e}")
    
    def _add_box_to_backup(self, backup_doc, pix, box, filename, page_idx, mat):
        """Add a box screenshot to backup PDF"""
        if box.box_type != 'label':
            return
        
        # Calculate scaled rect
        dpi_scale = OCR_DPI / 72
        scale = mat.a / dpi_scale  # Convert from OCR coords to screenshot coords
        
        rect = fitz.Rect(
            box.rect.x() * scale,
            box.rect.y() * scale,
            (box.rect.x() + box.rect.width()) * scale,
            (box.rect.y() + box.rect.height()) * scale
        )
        
        rect = rect.normalize() & fitz.Rect(0, 0, pix.width, pix.height)
        
        if rect.is_empty:
            return
        
        # Crop pixmap
        cropped = fitz.Pixmap(pix, fitz.IRect(rect))
        
        # Create new page
        page_width = max(cropped.width + 40, 200)
        page_height = cropped.height + 100
        backup_page = backup_doc.new_page(width=page_width, height=page_height)
        
        # Add header
        header = f"{filename} | Page {page_idx + 1} | {box.name}"
        backup_page.insert_text(fitz.Point(20, 25), header, fontsize=10)
        
        # Insert cropped image
        img_rect = fitz.Rect(20, 40, 20 + cropped.width, 40 + cropped.height)
        backup_page.insert_image(img_rect, pixmap=cropped)
    
    def _generate_extraction_backup(self):
        """Generate backup PDF with screenshots of extracted values"""
        if not self.extraction_screenshots:
            return
        
        path, _ = QFileDialog.getSaveFileName(self, "Save OCR Extraction Backup PDF", 
            "ocr_extraction_backup.pdf", "PDF Files (*.pdf)")
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
                
                src_doc = fitz.open(pdf_path)
                src_page = src_doc.load_page(page_idx)
                
                expanded_rect = value_rect + (-20, -20, 20, 20)
                expanded_rect = expanded_rect & src_page.rect
                
                mat = fitz.Matrix(2, 2)
                pix = src_page.get_pixmap(matrix=mat, clip=expanded_rect)
                
                img_width = pix.width
                img_height = pix.height
                page_width = img_width + 40
                page_height = img_height + 120
                
                backup_page = backup_doc.new_page(width=page_width, height=page_height)
                
                header_text = f"Source: {pdf_filename} | Page: {page_idx + 1} | Label: {label_name}"
                backup_page.insert_text(fitz.Point(20, 25), header_text, fontsize=10)
                
                img_rect = fitz.Rect(20, 40, 20 + img_width, 40 + img_height)
                backup_page.insert_image(img_rect, pixmap=pix)
                
                text_y = 40 + img_height + 20
                value_preview = value_text[:100] + "..." if len(value_text) > 100 else value_text
                backup_page.insert_text(fitz.Point(20, text_y), f"Value: {value_preview}", fontsize=9)
                
                src_doc.close()
            
            backup_doc.save(path)
            backup_doc.close()
            
            print(f"[DEBUG] OCR extraction backup saved to: {path}")
            
        except Exception as e:
            print(f"[DEBUG] Error generating OCR extraction backup: {e}")
            import traceback
            traceback.print_exc()
