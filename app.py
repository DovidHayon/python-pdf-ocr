import sys
import fitz
import pytesseract
import pandas
import json
import docx
from PIL import Image

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                             QLabel, QSplitter, QAction, QFileDialog,
                             QVBoxLayout, QPushButton, QScrollArea, QTextEdit,
                             QStackedWidget, QSpacerItem, QSizePolicy, QProgressBar, QMessageBox)
from PyQt5.QtGui import QPixmap, QImage, QPainter, QColor, QTextCursor, QFont
from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal, pyqtSlot, QRect, QEvent

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
QProgressBar {
    border: 1px solid #555555;
    border-radius: 3px;
    text-align: center;
    color: #f0f0f0;
}
QProgressBar::chunk {
    background-color: #009688;
    border-radius: 3px;
}
"""

def is_rtl_char(char):
    return '\u0590' <= char <= '\u05FF'

# =====================================================================
#  OCR Workers (Single and Batch) - Unchanged
# =====================================================================
class OCRWorker(QObject):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    def __init__(self, page_pixmap, zoom_factor):
        super().__init__()
        self.page_pixmap = page_pixmap
        self.zoom_factor = zoom_factor
    @pyqtSlot()
    def run(self):
        try:
            pil_image = Image.frombytes("RGB", (self.page_pixmap.width, self.page_pixmap.height), self.page_pixmap.samples)
            data = pytesseract.image_to_data(pil_image, lang='heb', output_type=pytesseract.Output.DATAFRAME, timeout=30)
            data = data.dropna(subset=['text']); data = data[data.conf > 30]
            full_text = ""; word_data = []; last_block, last_par, last_line = -1, -1, -1
            for index, row in data.iterrows():
                if last_block != -1:
                    block, par, line = row['block_num'], row['par_num'], row['line_num']
                    separator = " ";
                    if block != last_block or par != last_par: separator = '\n\n'
                    full_text += separator
                    for _ in separator: word_data.append(None)
                last_block, last_par, last_line = row['block_num'], row['par_num'], row['line_num']
                word_text = str(row['text']); full_text += word_text
                x, y, w, h = row['left'], row['top'], row['width'], row['height']
                normalized_word_bbox = [c / self.zoom_factor for c in [x, y, x + w, y + h]]; char_bboxes = []
                if len(word_text) > 0:
                    char_width = w / len(word_text)
                    for i, char in enumerate(word_text):
                        char_x = x + (i * char_width); char_bboxes.append([c / self.zoom_factor for c in [char_x, y, char_x + char_width, y + h]])
                if any(is_rtl_char(c) for c in word_text): char_bboxes.reverse()
                for char_bbox in char_bboxes: word_data.append({'word_bbox': normalized_word_bbox, 'char_bbox': char_bbox})
            self.finished.emit({'text': full_text, 'word_data': word_data})
        except Exception as e: self.error.emit(f"An unexpected OCR error occurred: {e}")

class OCRAllWorker(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    progress_updated = pyqtSignal(int, int, dict)
    def __init__(self, pdf_data, ocr_zoom_level=2.0):
        super().__init__(); self._is_canceled = False; self.pdf_data = pdf_data; self.ocr_zoom_level = ocr_zoom_level
    @pyqtSlot()
    def run(self):
        doc = fitz.open(stream=self.pdf_data, filetype="pdf")
        total_pages = len(doc)
        for i in range(total_pages):
            if self._is_canceled: break
            page = doc.load_page(i); mat = fitz.Matrix(self.ocr_zoom_level, self.ocr_zoom_level); pix = page.get_pixmap(matrix=mat)
            try:
                pil_image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                data = pytesseract.image_to_data(pil_image, lang='heb', output_type=pytesseract.Output.DATAFRAME, timeout=30)
                data = data.dropna(subset=['text']); data = data[data.conf > 30]
                full_text = ""; word_data = []; last_block, last_par, last_line = -1, -1, -1
                for index, row in data.iterrows():
                    if last_block != -1:
                        block, par, line = row['block_num'], row['par_num'], row['line_num']
                        separator = " ";
                        if block != last_block or par != last_par: separator = '\n\n'
                        full_text += separator
                        for _ in separator: word_data.append(None)
                    last_block, last_par, last_line = row['block_num'], row['par_num'], row['line_num']
                    word_text = str(row['text']); full_text += word_text
                    x, y, w, h = row['left'], row['top'], row['width'], row['height']
                    normalized_word_bbox = [c / self.ocr_zoom_level for c in [x, y, x + w, y + h]]; char_bboxes = []
                    if len(word_text) > 0:
                        char_width = w / len(word_text)
                        for i, char in enumerate(word_text):
                            char_x = x + (i * char_width); char_bboxes.append([c / self.ocr_zoom_level for c in [char_x, y, char_x + char_width, y + h]])
                    if any(is_rtl_char(c) for c in word_text): char_bboxes.reverse()
                    for char_bbox in char_bboxes: word_data.append({'word_bbox': normalized_word_bbox, 'char_bbox': char_bbox})
                page_data = {'word_data': word_data, 'edited_text': full_text}
                self.progress_updated.emit(i + 1, total_pages, page_data)
            except Exception as e:
                self.error.emit(f"Error on page {i+1}: {e}"); break
        doc.close()
        self.finished.emit()
    def cancel(self): self._is_canceled = True
        
# =====================================================================
#  InteractiveTextEdit (MODIFIED with final arrow key fix)
# =====================================================================
class InteractiveTextEdit(QTextEdit):
    elements_hovered = pyqtSignal(list, list)
    text_changed_by_user = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setLineWrapMode(QTextEdit.WidgetWidth)
        self.word_data = []
        self.cursorPositionChanged.connect(self.on_cursor_position_changed)
        self.textChanged.connect(self.on_text_changed)

    def wheelEvent(self, event):
        if event.modifiers() == Qt.ControlModifier:
            if event.angleDelta().y() > 0: self.zoomIn()
            else: self.zoomOut()
            event.accept()
        else:
            super().wheelEvent(event)

    def keyPressEvent(self, event):
        cursor = self.textCursor()
        if event.key() == Qt.Key_Left:
            cursor.movePosition(QTextCursor.NextCharacter)
            self.setTextCursor(cursor)
            event.accept()
        elif event.key() == Qt.Key_Right:
            cursor.movePosition(QTextCursor.PreviousCharacter)
            self.setTextCursor(cursor)
            event.accept()
        else:
            super().keyPressEvent(event)


    def set_word_data(self, data): self.word_data = data
    @pyqtSlot()
    def on_cursor_position_changed(self): self.update_highlight()
    def on_text_changed(self): self.text_changed_by_user.emit()
    def update_highlight(self):
        cursor = self.textCursor()
        pos = cursor.position()
        if cursor.hasSelection():
            pos = cursor.selectionEnd() if cursor.position() == cursor.selectionEnd() else cursor.selectionStart()
        if 0 <= pos < len(self.word_data):
            data_item = self.word_data[pos]
            if data_item: self.elements_hovered.emit(data_item['word_bbox'], data_item['char_bbox'])
            else: self.elements_hovered.emit([], [])
        else: self.elements_hovered.emit([], [])

# =====================================================================
#  PdfViewerWidget and PdfScrollArea are UNCHANGED
# =====================================================================
class PdfViewerWidget(QLabel):
    request_scroll = pyqtSignal(QRect)
    def __init__(self, parent=None):
        super().__init__(parent); self.current_pixmap = None; self.word_highlight_rect = None; self.char_highlight_rect = None
    def set_pixmap(self, pixmap):
        self.current_pixmap = pixmap; self.setPixmap(self.current_pixmap); self.word_highlight_rect = None; self.char_highlight_rect = None; self.update()
    @pyqtSlot(list, list)
    def highlight_elements(self, word_bbox, char_bbox):
        if word_bbox: self.word_highlight_rect = QRect(int(word_bbox[0]), int(word_bbox[1]), int(word_bbox[2]-word_bbox[0]), int(word_bbox[3]-word_bbox[1]))
        else: self.word_highlight_rect = None
        if char_bbox:
            self.char_highlight_rect = QRect(int(char_bbox[0]), int(char_bbox[1]), int(char_bbox[2]-char_bbox[0]), int(char_bbox[3]-char_bbox[1])); self.request_scroll.emit(self.char_highlight_rect)
        else: self.char_highlight_rect = None
        self.update()
    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.current_pixmap: return
        painter = QPainter(self)
        if self.word_highlight_rect:
            painter.setBrush(QColor(255, 255, 0, 80)); painter.setPen(Qt.NoPen); painter.drawRect(self.word_highlight_rect)
        if self.char_highlight_rect:
            painter.setBrush(QColor(0, 150, 255, 100)); painter.setPen(Qt.NoPen); painter.drawRect(self.char_highlight_rect)

class PdfScrollArea(QScrollArea):
    zoom_requested = pyqtSignal(int)
    def wheelEvent(self, event):
        if event.modifiers() == Qt.AltModifier:
            if event.angleDelta().y() > 0: self.zoom_requested.emit(1)
            else: self.zoom_requested.emit(-1)
            event.accept()
        else: super().wheelEvent(event)

# =====================================================================
#  Main Application Window (MODIFIED for final bug fixes)
# =====================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Interactive Local PDF OCR Tool"); self.setGeometry(100, 100, 1200, 800)
        self.doc = None; self.current_pdf_path = None; self.current_page_number = 0
        self.zoom_factor = 2.0; self.font_size = 14; self.ocr_data_cache = {}; self.is_dirty = False
        self.ocr_thread = None; self.ocr_worker = None; self.ocr_all_thread = None; self.ocr_all_worker = None
        self.setup_ui(); self.setup_menu()

    def set_dirty_flag(self): self.is_dirty = True

    def setup_ui(self):
        self.central_widget = QWidget(); self.setCentralWidget(self.central_widget)
        main_layout = QHBoxLayout(self.central_widget); self.splitter = QSplitter(Qt.Horizontal); main_layout.addWidget(self.splitter)
        self.pdf_stack = QStackedWidget()
        welcome_widget = QWidget(); welcome_layout = QVBoxLayout(welcome_widget); welcome_layout.setAlignment(Qt.AlignCenter)
        title_label = QLabel("Private OCR Tool"); title_label.setFont(QFont("Arial", 24)); title_label.setAlignment(Qt.AlignCenter)
        open_pdf_button = QPushButton("Open a New PDF"); open_pdf_button.setObjectName("WelcomeButton"); open_pdf_button.clicked.connect(self.open_pdf_file)
        load_project_button = QPushButton("Load Existing Project"); load_project_button.setObjectName("WelcomeButton"); load_project_button.clicked.connect(self.load_project)
        welcome_layout.addWidget(title_label); welcome_layout.addSpacing(20); welcome_layout.addWidget(open_pdf_button); welcome_layout.addWidget(load_project_button)
        pdf_viewer_container = QWidget(); pdf_viewer_layout = QVBoxLayout(pdf_viewer_container); pdf_viewer_layout.setContentsMargins(0, 0, 0, 0)
        self.pdf_viewer = PdfViewerWidget(); self.pdf_viewer.setAlignment(Qt.AlignCenter)
        self.scroll_area = PdfScrollArea(); self.scroll_area.setWidgetResizable(True); self.scroll_area.setWidget(self.pdf_viewer)
        pdf_viewer_layout.addWidget(self.scroll_area)
        controls_layout = QHBoxLayout()
        zoom_out_button = QPushButton("-"); zoom_out_button.clicked.connect(self.zoom_out); controls_layout.addWidget(zoom_out_button)
        zoom_in_button = QPushButton("+"); zoom_in_button.clicked.connect(self.zoom_in); controls_layout.addWidget(zoom_in_button)
        controls_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.prev_button = QPushButton("< Previous"); self.prev_button.clicked.connect(self.go_to_previous_page); controls_layout.addWidget(self.prev_button)
        self.page_number_label = QLabel("Page: N/A"); controls_layout.addWidget(self.page_number_label)
        self.next_button = QPushButton("Next >"); self.next_button.clicked.connect(self.go_to_next_page); controls_layout.addWidget(self.next_button)
        controls_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        pdf_viewer_layout.addLayout(controls_layout)
        self.pdf_stack.addWidget(welcome_widget); self.pdf_stack.addWidget(pdf_viewer_container)
        text_pane_container = QWidget(); text_pane_layout = QVBoxLayout(text_pane_container); text_pane_layout.setContentsMargins(0,0,0,0)
        self.text_editor = InteractiveTextEdit("Open a PDF or load a project to begin."); self.text_editor.setFontPointSize(self.font_size)
        self.text_editor.text_changed_by_user.connect(self.set_dirty_flag)
        text_pane_layout.addWidget(self.text_editor)
        ocr_controls_layout = QHBoxLayout()
        self.run_ocr_button = QPushButton("Run OCR on Current Page"); ocr_controls_layout.addWidget(self.run_ocr_button)
        self.run_ocr_all_button = QPushButton("Run OCR on All Pages"); ocr_controls_layout.addWidget(self.run_ocr_all_button)
        self.export_to_word_button = QPushButton("Export to Word"); ocr_controls_layout.addWidget(self.export_to_word_button)
        self.cancel_ocr_all_button = QPushButton("Cancel"); ocr_controls_layout.addWidget(self.cancel_ocr_all_button); self.cancel_ocr_all_button.hide()
        text_pane_layout.addLayout(ocr_controls_layout)
        self.ocr_status_label = QLabel(""); self.ocr_status_label.setAlignment(Qt.AlignCenter); text_pane_layout.addWidget(self.ocr_status_label)
        self.ocr_progress_bar = QProgressBar(); text_pane_layout.addWidget(self.ocr_progress_bar); self.ocr_progress_bar.hide()
        font_controls_layout = QHBoxLayout(); font_controls_layout.addItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        font_decrease_button = QPushButton("A-"); font_decrease_button.clicked.connect(self.decrease_font_size); font_controls_layout.addWidget(font_decrease_button)
        font_increase_button = QPushButton("A+"); font_increase_button.clicked.connect(self.increase_font_size); font_controls_layout.addWidget(font_increase_button)
        self.toggle_layout_button = QPushButton("Toggle Layout"); self.toggle_layout_button.clicked.connect(self.toggle_layout); font_controls_layout.addWidget(self.toggle_layout_button)
        text_pane_layout.addLayout(font_controls_layout)
        self.run_ocr_button.clicked.connect(self.start_ocr_process)
        self.run_ocr_all_button.clicked.connect(self.start_ocr_all_process)
        self.export_to_word_button.clicked.connect(self.export_to_word)
        self.cancel_ocr_all_button.clicked.connect(self.cancel_ocr_all)
        self.splitter.addWidget(self.pdf_stack); self.splitter.addWidget(text_pane_container); self.splitter.setSizes([700, 500])
        self.text_editor.elements_hovered.connect(self.handle_highlight_request)
        self.pdf_viewer.request_scroll.connect(self.auto_scroll_pdf_view)
        self.scroll_area.zoom_requested.connect(self.handle_scroll_zoom)
        self.update_navigation_controls()
    
    def closeEvent(self, event):
        if self.is_dirty:
            reply = QMessageBox.question(self, 'Unsaved Changes',
                                           "You have unsaved changes. Do you want to save them before exiting?",
                                           QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                           QMessageBox.Save)
            if reply == QMessageBox.Save:
                self.save_project()
                event.accept()
            elif reply == QMessageBox.Discard:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def keyPressEvent(self, event):
        if event.modifiers() == Qt.ControlModifier:
            if event.key() == Qt.Key_Plus or event.key() == Qt.Key_Equal: self.zoom_in(); event.accept(); return
            elif event.key() == Qt.Key_Minus: self.zoom_out(); event.accept(); return
        super().keyPressEvent(event)
        
    @pyqtSlot(int)
    def handle_scroll_zoom(self, direction):
        if direction > 0: self.zoom_in()
        else: self.zoom_out()

    @pyqtSlot(list, list)
    def handle_highlight_request(self, normalized_word_bbox, normalized_char_bbox):
        scaled_word_bbox = [c * self.zoom_factor for c in normalized_word_bbox] if normalized_word_bbox else []
        scaled_char_bbox = [c * self.zoom_factor for c in normalized_char_bbox] if normalized_char_bbox else []
        self.pdf_viewer.highlight_elements(scaled_word_bbox, scaled_char_bbox)

    def zoom_in(self):
        if not self.doc: return
        self.perform_zoom(0.2)
    def zoom_out(self):
        if not self.doc: return
        self.perform_zoom(-0.2)
    def perform_zoom(self, delta):
        saved_cursor_pos = self.text_editor.textCursor().position(); saved_scroll_val = self.text_editor.verticalScrollBar().value()
        self.zoom_factor = max(0.2, self.zoom_factor + delta); print(f"Zoom changed. New factor: {self.zoom_factor:.1f}")
        self.display_page(self.current_page_number)
        cursor = self.text_editor.textCursor(); cursor.setPosition(saved_cursor_pos); self.text_editor.setTextCursor(cursor)
        self.text_editor.verticalScrollBar().setValue(saved_scroll_val)
    
    def increase_font_size(self): self.font_size += 1; self.text_editor.setFontPointSize(self.font_size)
    def decrease_font_size(self): self.font_size = max(8, self.font_size - 1); self.text_editor.setFontPointSize(self.font_size)

    def toggle_layout(self):
        if self.splitter.orientation() == Qt.Horizontal:
            self.splitter.setOrientation(Qt.Vertical)
        else:
            self.splitter.setOrientation(Qt.Horizontal)

    @pyqtSlot(QRect)
    def auto_scroll_pdf_view(self, rect):
        viewport_rect = self.scroll_area.viewport().rect(); viewport_center = viewport_rect.center()
        new_x = rect.center().x() - viewport_center.x(); new_y = rect.center().y() - viewport_center.y()
        self.scroll_area.horizontalScrollBar().setValue(new_x); self.scroll_area.verticalScrollBar().setValue(new_y)

    def display_page(self, page_number):
        if not self.doc or not (0 <= page_number < len(self.doc)): return
        self.current_page_number = page_number; page = self.doc.load_page(page_number)
        mat = fitz.Matrix(self.zoom_factor, self.zoom_factor); pix = page.get_pixmap(matrix=mat)
        q_image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888); q_pixmap = QPixmap.fromImage(q_image)
        self.pdf_viewer.set_pixmap(q_pixmap)
        if str(page_number) in self.ocr_data_cache:
            page_data = self.ocr_data_cache[str(page_number)]
            self.text_editor.setText(page_data['edited_text']); self.text_editor.set_word_data(page_data['word_data'])
        else:
            self.text_editor.setText("Click 'Run OCR' to extract text from this page."); self.text_editor.set_word_data([])
        self.update_navigation_controls()
    
    def start_ocr_process(self):
        if not self.doc: return
        page = self.doc.load_page(self.current_page_number); ocr_zoom_level = 2.0; mat = fitz.Matrix(ocr_zoom_level, ocr_zoom_level); pix_for_ocr = page.get_pixmap(matrix=mat)
        self.run_ocr_button.setEnabled(False); self.text_editor.setText("OCR in progress...")
        self.ocr_thread = QThread(); self.ocr_worker = OCRWorker(pix_for_ocr, ocr_zoom_level); self.ocr_worker.moveToThread(self.ocr_thread)
        self.ocr_thread.started.connect(self.ocr_worker.run); self.ocr_worker.finished.connect(self.handle_ocr_results)
        self.ocr_worker.error.connect(self.handle_ocr_error); self.ocr_worker.finished.connect(self.ocr_thread.quit)
        self.ocr_worker.finished.connect(self.ocr_worker.deleteLater); self.ocr_thread.finished.connect(self.ocr_thread.deleteLater)
        self.ocr_thread.start()
    
    def start_ocr_all_process(self):
        if not self.doc: return
        if self.ocr_all_thread and self.ocr_all_thread.isRunning():
            self.ocr_status_label.setText("Batch OCR is already running.")
            return

        self.set_ocr_all_ui_state(is_running=True)
        with open(self.current_pdf_path, 'rb') as f:
            pdf_data = f.read()
        self.ocr_all_thread = QThread(); self.ocr_all_worker = OCRAllWorker(pdf_data); self.ocr_all_worker.moveToThread(self.ocr_all_thread)
        self.ocr_all_thread.started.connect(self.ocr_all_worker.run); self.ocr_all_worker.progress_updated.connect(self.handle_ocr_all_progress)
        self.ocr_all_worker.finished.connect(self.handle_ocr_all_finished); self.ocr_all_worker.error.connect(self.handle_ocr_error)
        self.ocr_all_thread.start()

    def cancel_ocr_all(self):
        if self.ocr_all_worker: self.ocr_all_worker.cancel(); self.ocr_status_label.setText("Canceling...")

    @pyqtSlot(int, int, dict)
    def handle_ocr_all_progress(self, page_num, total_pages, page_data):
        self.ocr_progress_bar.setValue(page_num); self.ocr_progress_bar.setMaximum(total_pages); self.ocr_status_label.setText(f"Processing page {page_num} of {total_pages}...")
        self.ocr_data_cache[str(page_num - 1)] = page_data

    def handle_ocr_all_finished(self):
        self.set_ocr_all_ui_state(is_running=False)
        self.ocr_status_label.setText("Batch OCR finished.")
        
        # Clean up the thread and worker
        if self.ocr_all_thread:
            self.ocr_all_thread.quit()
            self.ocr_all_thread.wait()
            self.ocr_all_thread.deleteLater()
            self.ocr_all_thread = None
        if self.ocr_all_worker:
            self.ocr_all_worker.deleteLater()
            self.ocr_all_worker = None

        # Refresh the display of the current page
        self.display_page(self.current_page_number)

    def set_ocr_all_ui_state(self, is_running):
        self.run_ocr_button.setDisabled(is_running); self.run_ocr_all_button.setDisabled(is_running)
        if is_running:
            self.ocr_progress_bar.setValue(0); self.ocr_progress_bar.show(); self.ocr_status_label.setText("Starting batch OCR..."); self.cancel_ocr_all_button.show()
        else:
            self.ocr_progress_bar.hide(); self.cancel_ocr_all_button.hide()
    
    def setup_menu(self):
        menubar = self.menuBar(); file_menu = menubar.addMenu('&File')
        open_action = QAction('&Open PDF', self); open_action.triggered.connect(self.open_pdf_file); file_menu.addAction(open_action)
        save_action = QAction('&Save Project', self); save_action.triggered.connect(self.save_project); file_menu.addAction(save_action)
        load_action = QAction('&Load Project', self); load_action.triggered.connect(self.load_project); file_menu.addAction(load_action)
        file_menu.addSeparator(); exit_action = QAction('&Exit', self); exit_action.triggered.connect(self.close); file_menu.addAction(exit_action)
    def open_pdf_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Open PDF File", "", "PDF Files (*.pdf)");
        if filepath: self.load_pdf(filepath)
    def load_pdf(self, filepath, is_project_load=False):
        if self.doc: self.doc.close()
        if not is_project_load: self.ocr_data_cache.clear()
        try:
            self.doc = fitz.open(filepath); self.current_pdf_path = filepath; self.current_page_number = 0
            self.pdf_stack.setCurrentIndex(1); self.display_page(self.current_page_number)
        except Exception as e:
            self.pdf_stack.setCurrentIndex(0); print(f"Failed to load PDF: {e}"); self.doc = None
        finally: self.update_navigation_controls()
    @pyqtSlot(dict)
    def handle_ocr_results(self, result_dict):
        self.text_editor.setText(result_dict['text']); self.text_editor.set_word_data(result_dict['word_data'])
        self.ocr_data_cache[str(self.current_page_number)] = {'word_data': result_dict['word_data'], 'edited_text': result_dict['text']}
        self.run_ocr_button.setEnabled(True)
    def save_project(self):
        if not self.current_pdf_path: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "JSON Files (*.json)")
        if save_path:
            if str(self.current_page_number) in self.ocr_data_cache:
                self.ocr_data_cache[str(self.current_page_number)]['edited_text'] = self.text_editor.toPlainText()
            project_data = {'pdf_path': self.current_pdf_path, 'ocr_data': self.ocr_data_cache}
            try:
                with open(save_path, 'w', encoding='utf-8') as f: json.dump(project_data, f, ensure_ascii=False, indent=4)
                print(f"Project saved to {save_path}")
                self.is_dirty = False
            except Exception as e: print(f"Error saving project: {e}")

    def export_to_word(self):
        if not self.ocr_data_cache:
            self.ocr_status_label.setText("No OCR data to export.")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Export to Word", "", "Word Documents (*.docx)")
        if save_path:
            try:
                doc = docx.Document()
                # Sort the pages by page number
                sorted_pages = sorted(self.ocr_data_cache.items(), key=lambda item: int(item[0]))
                for page_num, page_data in sorted_pages:
                    doc.add_paragraph(page_data['edited_text'])
                    doc.add_page_break()

                doc.save(save_path)
                self.ocr_status_label.setText(f"Exported to {save_path}")
            except Exception as e:
                self.ocr_status_label.setText(f"Error exporting to Word: {e}")
    def load_project(self):
        load_path, _ = QFileDialog.getOpenFileName(self, "Load Project", "", "JSON Files (*.json)")
        if load_path:
            try:
                with open(load_path, 'r', encoding='utf-8') as f: project_data = json.load(f)
                self.ocr_data_cache = project_data['ocr_data']; self.load_pdf(project_data['pdf_path'], is_project_load=True)
                print(f"Project loaded from {load_path}")
            except Exception as e: print(f"Error loading project: {e}")
    def go_to_next_page(self):
        if str(self.current_page_number) in self.ocr_data_cache: self.ocr_data_cache[str(self.current_page_number)]['edited_text'] = self.text_editor.toPlainText()
        if self.doc and self.current_page_number < len(self.doc) - 1:
            self.current_page_number += 1; self.display_page(self.current_page_number)
    def go_to_previous_page(self):
        if str(self.current_page_number) in self.ocr_data_cache: self.ocr_data_cache[str(self.current_page_number)]['edited_text'] = self.text_editor.toPlainText()
        if self.doc and self.current_page_number > 0:
            self.current_page_number -= 1; self.display_page(self.current_page_number)
    def handle_ocr_error(self, error_message):
        self.ocr_status_label.setText(f"Error: {error_message}"); self.set_ocr_all_ui_state(is_running=False)
        self.run_ocr_button.setEnabled(True)
    def update_navigation_controls(self):
        doc_is_loaded = self.doc is not None
        self.prev_button.setEnabled(doc_is_loaded and self.current_page_number > 0)
        self.next_button.setEnabled(doc_is_loaded and self.current_page_number < len(self.doc) - 1)
        self.run_ocr_button.setEnabled(doc_is_loaded); self.run_ocr_all_button.setEnabled(doc_is_loaded)
        if doc_is_loaded: self.page_number_label.setText(f"Page {self.current_page_number + 1} / {len(self.doc)}")
        else: self.page_number_label.setText("Page: N/A")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
