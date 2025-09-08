import sys
import fitz
import pytesseract
import pandas
import json
from PIL import Image

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                             QLabel, QSplitter, QAction, QFileDialog,
                             QVBoxLayout, QPushButton, QScrollArea, QTextEdit,
                             QStackedWidget, QSpacerItem, QSizePolicy)
from PyQt5.QtGui import QPixmap, QImage, QPainter, QColor, QTextCursor, QFont
from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal, pyqtSlot, QRect

# =====================================================================
#  Dark Theme Stylesheet and Helper Function (Unchanged)
# =====================================================================
DARK_STYLESHEET = """
QWidget {
    background-color: #2b2b2b;
    color: #f0f0f0;
    border: none;
}
QMainWindow {
    background-color: #2b2b2b;
}
QMenuBar {
    background-color: #3c3c3c;
    color: #f0f0f0;
}
QMenuBar::item:selected {
    background-color: #555555;
}
QMenu {
    background-color: #3c3c3c;
    color: #f0f0f0;
}
QMenu::item:selected {
    background-color: #555555;
}
QPushButton {
    background-color: #555555;
    border: 1px solid #666666;
    padding: 5px;
    border-radius: 3px;
    min-width: 30px;
}
QPushButton:hover {
    background-color: #6a6a6a;
}
QPushButton:pressed {
    background-color: #787878;
}
QPushButton:disabled {
    background-color: #444444;
    color: #888888;
}
QTextEdit, QScrollArea {
    background-color: #3c3c3c;
    border: 1px solid #555555;
    border-radius: 3px;
}
QSplitter::handle {
    background-color: #555555;
}
QLabel {
    color: #f0f0f0;
}
QPushButton#WelcomeButton {
    padding: 15px;
    font-size: 16px;
}
"""

def is_rtl_char(char):
    return '\u0590' <= char <= '\u05FF'

# =====================================================================
#  OCR Worker Thread (MODIFIED to normalize coordinates)
# =====================================================================
class OCRWorker(QObject):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    # ## MODIFIED ##: It now needs to know the zoom factor used for the OCR image.
    def __init__(self, page_pixmap, zoom_factor):
        super().__init__()
        self.page_pixmap = page_pixmap
        self.zoom_factor = zoom_factor

    @pyqtSlot()
    def run(self):
        print(f"OCR Worker: Starting OCR at zoom {self.zoom_factor}...")
        try:
            pil_image = Image.frombytes("RGB", (self.page_pixmap.width, self.page_pixmap.height), self.page_pixmap.samples)
            data = pytesseract.image_to_data(pil_image, lang='heb', output_type=pytesseract.Output.DATAFRAME)
            data = data.dropna(subset=['text'])
            data = data[data.conf > 30]

            full_text = ""
            # This list will hold the NORMALIZED (1.0x zoom) coordinates
            normalized_char_bboxes = []
            
            for index, row in data.iterrows():
                word_text = str(row['text'])
                if word_text.strip():
                    full_text += word_text + " "
                    x, y, w, h = row['left'], row['top'], row['width'], row['height']
                    
                    word_char_bboxes_scaled = [] # Temp list for scaled bboxes
                    if len(word_text) > 0:
                        char_width = w / len(word_text)
                        for i, char in enumerate(word_text):
                            char_x = x + (i * char_width)
                            word_char_bboxes_scaled.append([char_x, y, char_x + char_width, y + h])

                    if any(is_rtl_char(c) for c in word_text):
                        word_char_bboxes_scaled.reverse()

                    # ## NEW LOGIC ##: Normalize every coordinate before adding it to the final list.
                    for bbox in word_char_bboxes_scaled:
                        normalized_bbox = [coord / self.zoom_factor for coord in bbox]
                        normalized_char_bboxes.append(normalized_bbox)
                    
                    normalized_char_bboxes.append(None)
            
            if normalized_char_bboxes: normalized_char_bboxes.pop()

            self.finished.emit({'text': full_text.strip(), 'char_bboxes': normalized_char_bboxes})

        except pytesseract.TesseractNotFoundError:
            self.error.emit("Tesseract Error: 'tesseract' not found.")
        except Exception as e:
            self.error.emit(f"An unexpected error occurred during OCR: {e}")

# =====================================================================
#  PdfViewerWidget and InteractiveTextEdit are UNCHANGED
# =====================================================================
class PdfViewerWidget(QLabel):
    request_scroll = pyqtSignal(QRect)
    # ... (code is exactly the same as before) ...
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_pixmap = None
        self.highlight_rect = None

    def set_pixmap(self, pixmap):
        self.current_pixmap = pixmap
        self.setPixmap(self.current_pixmap)
        self.highlight_rect = None
        self.update()

    @pyqtSlot(list)
    def highlight_char(self, bbox):
        if bbox:
            # This widget now receives SCALED coordinates, so it just needs to draw them.
            self.highlight_rect = QRect(int(bbox[0]), int(bbox[1]), int(bbox[2]-bbox[0]), int(bbox[3]-bbox[1]))
            self.request_scroll.emit(self.highlight_rect)
        else:
            self.highlight_rect = None
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.highlight_rect and self.current_pixmap:
            painter = QPainter(self)
            painter.setBrush(QColor(0, 150, 255, 100))
            painter.setPen(Qt.NoPen)
            painter.drawRect(self.highlight_rect)

class InteractiveTextEdit(QTextEdit):
    char_hovered = pyqtSignal(list)
    # ... (code is exactly the same as before) ...
    def __init__(self, parent=None):
        super().__init__(parent)
        self.char_bboxes = []
        self.cursorPositionChanged.connect(self.on_cursor_position_changed)

    def set_char_bboxes(self, bboxes):
        # This now receives NORMALIZED coordinates.
        self.char_bboxes = bboxes

    @pyqtSlot()
    def on_cursor_position_changed(self):
        self.update_highlight()

    def update_highlight(self):
        cursor = self.textCursor()
        pos = cursor.position()
        
        if cursor.hasSelection():
            pos = cursor.selectionEnd() if cursor.position() == cursor.selectionEnd() else cursor.selectionStart()

        if 0 <= pos < len(self.char_bboxes):
            bbox = self.char_bboxes[pos]
            # This emits the NORMALIZED coordinates.
            self.char_hovered.emit(bbox if bbox else [])
        else:
            self.char_hovered.emit([])

# =====================================================================
#  Main Application Window (MODIFIED to handle coordinate scaling)
# =====================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        # ... (init variables are unchanged) ...
        super().__init__()
        self.setWindowTitle("Interactive Local PDF OCR Tool")
        self.setGeometry(100, 100, 1200, 800)
        self.doc = None
        self.current_pdf_path = None
        self.current_page_number = 0
        self.zoom_factor = 2.0
        self.font_size = 14
        self.ocr_data_cache = {}
        self.current_fitz_pixmap = None
        self.ocr_thread = None
        self.ocr_worker = None
        self.setup_ui()
        self.setup_menu()

    def setup_ui(self):
        # ... (UI setup is unchanged) ...
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QHBoxLayout(self.central_widget)
        self.splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(self.splitter)
        self.pdf_stack = QStackedWidget()
        welcome_widget = QWidget()
        welcome_layout = QVBoxLayout(welcome_widget)
        welcome_layout.setAlignment(Qt.AlignCenter)
        title_label = QLabel("Private OCR Tool")
        title_label.setFont(QFont("Arial", 24))
        title_label.setAlignment(Qt.AlignCenter)
        open_pdf_button = QPushButton("Open a New PDF")
        open_pdf_button.setObjectName("WelcomeButton")
        open_pdf_button.clicked.connect(self.open_pdf_file)
        load_project_button = QPushButton("Load Existing Project")
        load_project_button.setObjectName("WelcomeButton")
        load_project_button.clicked.connect(self.load_project)
        welcome_layout.addWidget(title_label)
        welcome_layout.addSpacing(20)
        welcome_layout.addWidget(open_pdf_button)
        welcome_layout.addWidget(load_project_button)
        pdf_viewer_container = QWidget()
        pdf_viewer_layout = QVBoxLayout(pdf_viewer_container)
        pdf_viewer_layout.setContentsMargins(0, 0, 0, 0)
        self.pdf_viewer = PdfViewerWidget()
        self.pdf_viewer.setAlignment(Qt.AlignCenter)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.pdf_viewer)
        pdf_viewer_layout.addWidget(self.scroll_area)
        controls_layout = QHBoxLayout()
        zoom_out_button = QPushButton("-")
        zoom_out_button.clicked.connect(self.zoom_out)
        controls_layout.addWidget(zoom_out_button)
        zoom_in_button = QPushButton("+")
        zoom_in_button.clicked.connect(self.zoom_in)
        controls_layout.addWidget(zoom_in_button)
        controls_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.prev_button = QPushButton("< Previous")
        self.prev_button.clicked.connect(self.go_to_previous_page)
        controls_layout.addWidget(self.prev_button)
        self.page_number_label = QLabel("Page: N/A")
        controls_layout.addWidget(self.page_number_label)
        self.next_button = QPushButton("Next >")
        self.next_button.clicked.connect(self.go_to_next_page)
        controls_layout.addWidget(self.next_button)
        controls_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        pdf_viewer_layout.addLayout(controls_layout)
        self.pdf_stack.addWidget(welcome_widget)
        self.pdf_stack.addWidget(pdf_viewer_container)
        text_pane_container = QWidget()
        text_pane_layout = QVBoxLayout(text_pane_container)
        text_pane_layout.setContentsMargins(0,0,0,0)
        self.text_editor = InteractiveTextEdit("Open a PDF or load a project to begin.")
        self.text_editor.setFontPointSize(self.font_size)
        text_pane_layout.addWidget(self.text_editor)
        text_controls_layout = QHBoxLayout()
        self.run_ocr_button = QPushButton("Run OCR on Current Page")
        text_controls_layout.addWidget(self.run_ocr_button)
        text_controls_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        font_decrease_button = QPushButton("A-")
        font_decrease_button.clicked.connect(self.decrease_font_size)
        text_controls_layout.addWidget(font_decrease_button)
        font_increase_button = QPushButton("A+")
        font_increase_button.clicked.connect(self.increase_font_size)
        text_controls_layout.addWidget(font_increase_button)
        text_pane_layout.addLayout(text_controls_layout)
        self.run_ocr_button.clicked.connect(self.start_ocr_process)
        self.splitter.addWidget(self.pdf_stack)
        self.splitter.addWidget(text_pane_container)
        self.splitter.setSizes([700, 500])
        
        # ## MODIFIED ##: The signal from the text editor is now connected to a new handler in MainWindow.
        self.text_editor.char_hovered.connect(self.handle_highlight_request)
        self.pdf_viewer.request_scroll.connect(self.auto_scroll_pdf_view)

        self.update_navigation_controls()
        
    # ## NEW ##: This slot is the core of the fix.
    # It catches the normalized coordinates and scales them before highlighting.
    @pyqtSlot(list)
    def handle_highlight_request(self, normalized_bbox):
        if not normalized_bbox:
            # If the bbox is empty, just clear the highlight.
            scaled_bbox = []
        else:
            # Scale the normalized coordinates by the current zoom factor.
            scaled_bbox = [coord * self.zoom_factor for coord in normalized_bbox]
        
        # Send the correctly scaled coordinates to the PDF viewer for drawing.
        self.pdf_viewer.highlight_char(scaled_bbox)

    def zoom_in(self):
        self.zoom_factor += 0.2
        print(f"Zooming in. New factor: {self.zoom_factor:.1f}")
        # The re-render will now create a new, larger pixmap.
        self.display_page(self.current_page_number, re_render=True)

    def zoom_out(self):
        self.zoom_factor = max(0.2, self.zoom_factor - 0.2)
        print(f"Zooming out. New factor: {self.zoom_factor:.1f}")
        self.display_page(self.current_page_number, re_render=True)
    
    # ## MODIFIED ##: When zooming, we must re-render the page but NOT re-run OCR.
    def display_page(self, page_number, re_render=False):
        if not self.doc or not (0 <= page_number < len(self.doc)): return
        
        self.current_page_number = page_number
        page = self.doc.load_page(page_number)
        
        # The key change is that re-rendering now ONLY affects the view.
        # It does not change the underlying OCR data, which is now zoom-independent.
        mat = fitz.Matrix(self.zoom_factor, self.zoom_factor)
        pix = page.get_pixmap(matrix=mat)

        image_format = QImage.Format_RGB888
        q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, image_format)
        q_pixmap = QPixmap.fromImage(q_image)
        self.pdf_viewer.set_pixmap(q_pixmap)

        if str(page_number) in self.ocr_data_cache:
            page_data = self.ocr_data_cache[str(page_number)]
            self.text_editor.setText(page_data['edited_text'])
            self.text_editor.set_char_bboxes(page_data['char_bboxes'])
        else:
            self.text_editor.setText("Click 'Run OCR' to extract text from this page.")
            self.text_editor.set_char_bboxes([])
        
        self.update_navigation_controls()

    # ## MODIFIED ##: We must now pass the current zoom factor to the OCR worker.
    def start_ocr_process(self):
        # We need a pixmap to get its dimensions for the OCR worker.
        # But we create a fresh one here to ensure we use the current zoom factor.
        if self.doc:
            page = self.doc.load_page(self.current_page_number)
            mat = fitz.Matrix(self.zoom_factor, self.zoom_factor)
            pix_for_ocr = page.get_pixmap(matrix=mat)

            self.run_ocr_button.setEnabled(False)
            self.text_editor.setText("OCR in progress...")
            self.ocr_thread = QThread()
            # Pass the pixmap AND the zoom factor to the worker.
            self.ocr_worker = OCRWorker(pix_for_ocr, self.zoom_factor)
            self.ocr_worker.moveToThread(self.ocr_thread)
            self.ocr_thread.started.connect(self.ocr_worker.run)
            self.ocr_worker.finished.connect(self.handle_ocr_results)
            self.ocr_worker.error.connect(self.handle_ocr_error)
            self.ocr_worker.finished.connect(self.ocr_thread.quit)
            self.ocr_worker.finished.connect(self.ocr_worker.deleteLater)
            self.ocr_thread.finished.connect(self.ocr_thread.deleteLater)
            self.ocr_thread.start()

    # The rest of the MainWindow methods are unchanged...
    def setup_menu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu('&File')
        open_action = QAction('&Open PDF', self)
        open_action.triggered.connect(self.open_pdf_file)
        file_menu.addAction(open_action)
        save_action = QAction('&Save Project', self)
        save_action.triggered.connect(self.save_project)
        file_menu.addAction(save_action)
        load_action = QAction('&Load Project', self)
        load_action.triggered.connect(self.load_project)
        file_menu.addAction(load_action)
        file_menu.addSeparator()
        exit_action = QAction('&Exit', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def open_pdf_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Open PDF File", "", "PDF Files (*.pdf)")
        if filepath:
            self.load_pdf(filepath)
    
    def load_pdf(self, filepath, is_project_load=False):
        if self.doc: self.doc.close()
        if not is_project_load:
            self.ocr_data_cache.clear()
        try:
            self.doc = fitz.open(filepath)
            self.current_pdf_path = filepath
            self.current_page_number = 0
            self.pdf_stack.setCurrentIndex(1)
            self.display_page(self.current_page_number)
        except Exception as e:
            self.pdf_stack.setCurrentIndex(0)
            print(f"Failed to load PDF: {e}")
            self.doc = None
        finally:
            self.update_navigation_controls()

    @pyqtSlot(QRect)
    def auto_scroll_pdf_view(self, rect):
        self.scroll_area.ensureVisible(rect.x(), rect.y(), xMargin=50, yMargin=50)

    @pyqtSlot(dict)
    def handle_ocr_results(self, result_dict):
        print("Main thread: Received NORMALIZED char-level OCR results.")
        self.text_editor.setText(result_dict['text'])
        self.text_editor.set_char_bboxes(result_dict['char_bboxes'])
        self.ocr_data_cache[str(self.current_page_number)] = {
            'char_bboxes': result_dict['char_bboxes'],
            'edited_text': result_dict['text']
        }
        self.run_ocr_button.setEnabled(True)
    
    def save_project(self):
        if not self.current_pdf_path:
            print("No PDF loaded, nothing to save.")
            return
        if str(self.current_page_number) in self.ocr_data_cache:
            self.ocr_data_cache[str(self.current_page_number)]['edited_text'] = self.text_editor.toPlainText()
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "JSON Files (*.json)")
        if save_path:
            project_data = {
                'pdf_path': self.current_pdf_path,
                'ocr_data': self.ocr_data_cache
            }
            try:
                with open(save_path, 'w', encoding='utf-8') as f:
                    json.dump(project_data, f, ensure_ascii=False, indent=4)
                print(f"Project saved to {save_path}")
            except Exception as e:
                print(f"Error saving project: {e}")

    def load_project(self):
        load_path, _ = QFileDialog.getOpenFileName(self, "Load Project", "", "JSON Files (*.json)")
        if load_path:
            try:
                with open(load_path, 'r', encoding='utf-8') as f:
                    project_data = json.load(f)
                self.ocr_data_cache = project_data['ocr_data']
                self.load_pdf(project_data['pdf_path'], is_project_load=True)
                print(f"Project loaded from {load_path}")
            except Exception as e:
                print(f"Error loading project: {e}")

    def go_to_next_page(self):
        if str(self.current_page_number) in self.ocr_data_cache:
            self.ocr_data_cache[str(self.current_page_number)]['edited_text'] = self.text_editor.toPlainText()
        if self.doc and self.current_page_number < len(self.doc) - 1:
            self.current_page_number += 1
            self.display_page(self.current_page_number)

    def go_to_previous_page(self):
        if str(self.current_page_number) in self.ocr_data_cache:
            self.ocr_data_cache[str(self.current_page_number)]['edited_text'] = self.text_editor.toPlainText()
        if self.doc and self.current_page_number > 0:
            self.current_page_number -= 1
            self.display_page(self.current_page_number)
            
    def handle_ocr_error(self, error_message):
        print(f"Main thread: Received OCR error: {error_message}")
        self.text_editor.setText(error_message)
        self.run_ocr_button.setEnabled(True)

    def update_navigation_controls(self):
        doc_is_loaded = self.doc is not None
        self.prev_button.setEnabled(doc_is_loaded and self.current_page_number > 0)
        self.next_button.setEnabled(doc_is_loaded and self.current_page_number < len(self.doc) - 1)
        self.run_ocr_button.setEnabled(doc_is_loaded)
        if doc_is_loaded:
            self.page_number_label.setText(f"Page {self.current_page_number + 1} / {len(self.doc)}")
        else:
            self.page_number_label.setText("Page: N/A")
    def increase_font_size(self):
        self.font_size += 1
        self.text_editor.setFontPointSize(self.font_size)

    def decrease_font_size(self):
        self.font_size = max(8, self.font_size - 1)
        self.text_editor.setFontPointSize(self.font_size)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
