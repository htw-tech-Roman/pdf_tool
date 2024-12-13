import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QLabel, QVBoxLayout, QPushButton, QGraphicsView, QGraphicsScene, QGraphicsRectItem, QListWidget, QHBoxLayout, QListWidgetItem, QWidget, QHBoxLayout, QProgressBar, QMessageBox
)
from PyQt5.QtGui import QPixmap, QColor, QImage, QCursor, QPen, QFont
from PyQt5.QtCore import Qt, QRectF, QThread, pyqtSignal
from fitz import Document, Matrix, Rect
import pandas as pd

class PDFProcessingThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, template_path, pdf_paths, output_path):
        super().__init__()
        self.template_path = template_path
        self.pdf_paths = pdf_paths
        self.output_path = output_path

    def run(self):
        try:
            with open(self.template_path, 'r') as f:
                template = json.load(f)

            if not isinstance(template, list):
                raise ValueError("Template file must contain a list of pages.")

            for page in template:
                if not isinstance(page, dict) or 'page' not in page or 'coordinates' not in page:
                    raise ValueError("Each page in the template must contain 'page' and 'coordinates' keys.")
                if not isinstance(page['coordinates'], list):
                    raise ValueError("The 'coordinates' key must be associated with a list of areas.")
                for area in page['coordinates']:
                    if not isinstance(area, dict) or not all(k in area for k in ['x', 'y', 'width', 'height']):
                        raise ValueError("Each area must contain 'x', 'y', 'width', and 'height' keys.")

            data = []

            for idx, pdf_path in enumerate(self.pdf_paths):
                doc = Document(pdf_path)

                for page_template in template:
                    page_num = page_template['page']
                    coordinates = page_template['coordinates']

                    page = doc[page_num]

                    for area in coordinates:
                        rect = Rect(
                            area['x'],
                            area['y'],
                            area['x'] + area['width'],
                            area['y'] + area['height']
                        )
                        text = page.get_text("blocks", clip=rect)
                        extracted_text = "\n".join(block[4].strip() for block in text)
                        data.append(extracted_text)

                # Add a blank line between files
                data.append("")

                self.progress.emit(int((idx + 1) / len(self.pdf_paths) * 100))

            # Save data to Excel
            df = pd.DataFrame({"Extracted Text": data})
            df.to_excel(self.output_path, index=False, header=False)

        except Exception as e:
            print(f"Error: {e}")

        self.finished.emit()

class PDFMarkupTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Markup Tool")
        self.resize(800, 600)
        self.show()

        # Main layout
        self.main_layout = QHBoxLayout()
        self.side_layout = QVBoxLayout()

        self.pdf_document = None
        self.current_page_index = 0
        self.selected_areas = []  # Stores areas in original scale
        self.rect_items = []  # Stores QGraphicsRectItems
        self.scale_factor = 1.0

        # Load button
        self.load_button = QPushButton("Load PDF")
        self.load_button.clicked.connect(self.load_pdf)
        self.side_layout.addWidget(self.load_button)

        # Process files button
        self.process_button = QPushButton("Process Files")
        self.process_button.clicked.connect(self.process_files)
        self.side_layout.addWidget(self.process_button)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.side_layout.addWidget(self.progress_bar)

        # Graphics View for PDF display
        self.view = QGraphicsView()
        self.scene = QGraphicsScene()
        self.view.setScene(self.scene)
        self.view.setMouseTracking(True)
        self.main_layout.addWidget(self.view, stretch=3)

        # List of rectangles
        self.rect_list = QListWidget()
        self.rect_list.setFixedWidth(200)
        self.side_layout.addWidget(self.rect_list)

        # Navigation buttons
        self.prev_button = QPushButton("Previous Page")
        self.prev_button.clicked.connect(self.prev_page)
        self.side_layout.addWidget(self.prev_button)

        self.next_button = QPushButton("Next Page")
        self.next_button.clicked.connect(self.next_page)
        self.side_layout.addWidget(self.next_button)

        # Save template button
        self.save_button = QPushButton("Save Template")
        self.save_button.clicked.connect(self.save_template)
        self.side_layout.addWidget(self.save_button)

        # Add side layout to main layout
        self.main_layout.addLayout(self.side_layout)

        # Container widget
        self.container = QWidget()
        self.container.setLayout(self.main_layout)
        self.setCentralWidget(self.container)

        # Variables for selection
        self.start_pos = None
        self.rect_item = None
        self.is_shift_pressed = False

        # Event filter for capturing mouse events
        self.view.viewport().installEventFilter(self)

    def load_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf_document = Document(file_path)
            self.current_page_index = 0
            self.show_page()

    def process_files(self):
        template_path, _ = QFileDialog.getOpenFileName(self, "Open Template", "", "JSON Files (*.json)")
        if not template_path:
            return

        pdf_paths, _ = QFileDialog.getOpenFileNames(self, "Select PDF Files", "", "PDF Files (*.pdf)")
        if not pdf_paths:
            return

        output_path, _ = QFileDialog.getSaveFileName(self, "Save Output", "", "Excel Files (*.xlsx)")
        if not output_path:
            return

        self.thread = PDFProcessingThread(template_path, pdf_paths, output_path)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.finished.connect(self.processing_finished)
        self.thread.start()

    def processing_finished(self):
        QMessageBox.information(self, "Processing Complete", "The files have been processed and saved successfully.")
        self.progress_bar.setValue(0)

    def show_page(self):
        if not self.pdf_document:
            return

        # Clear current scene
        self.scene.clear()
        self.rect_items.clear()
        self.rect_list.clear()

        # Render current page to QPixmap and add to the scene
        page = self.pdf_document[self.current_page_index]
        matrix = Matrix(self.scale_factor, self.scale_factor)
        pix = page.get_pixmap(matrix=matrix)
        image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(image)
        self.scene.addPixmap(pixmap)

        # Reapply selected areas
        for index, rect in enumerate(self.selected_areas):
            scaled_rect = QRectF(
                rect.x() * self.scale_factor,
                rect.y() * self.scale_factor,
                rect.width() * self.scale_factor,
                rect.height() * self.scale_factor
            )
            fixed_rect = QGraphicsRectItem(scaled_rect)
            fixed_rect.setPen(QPen(QColor("red")))
            self.scene.addItem(fixed_rect)
            self.rect_items.append(fixed_rect)

            # Add label with number inside the rectangle
            label = self.scene.addText(f"{index + 1}")
            label.setDefaultTextColor(QColor("blue"))
            label.setFont(QFont("Arial", 12))
            label.setPos(scaled_rect.x(), scaled_rect.y())

            # Add to rect_list
            item_widget = QWidget()
            item_layout = QHBoxLayout()
            item_layout.setContentsMargins(0, 0, 0, 0)
            item_layout.setSpacing(5)

            label_text = QLabel(f"Frame №{index + 1}")
            remove_button = QPushButton(f"Remove №{index + 1}")
            remove_button.clicked.connect(lambda _, i=index: self.remove_rect(i))

            item_layout.addWidget(label_text)
            item_layout.addWidget(remove_button)
            item_widget.setLayout(item_layout)

            list_item = QListWidgetItem()
            self.rect_list.addItem(list_item)
            self.rect_list.setItemWidget(list_item, item_widget)

    def prev_page(self):
        if self.current_page_index > 0:
            self.current_page_index -= 1
            self.show_page()

    def next_page(self):
        if self.current_page_index < len(self.pdf_document) - 1:
            self.current_page_index += 1
            self.show_page()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Shift:
            self.is_shift_pressed = True
            self.setCursor(QCursor(Qt.CrossCursor))

    def keyReleaseEvent(self, event):
        if event.key() == Qt.Key_Shift:
            self.is_shift_pressed = False
            self.setCursor(QCursor(Qt.ArrowCursor))

    def wheelEvent(self, event):
        modifiers = QApplication.keyboardModifiers()
        if modifiers == Qt.ControlModifier:
            if event.angleDelta().y() > 0:
                self.scale_factor *= 1.1
            else:
                self.scale_factor *= 0.9

            self.show_page()

    def eventFilter(self, source, event):
        if source == self.view.viewport():
            if event.type() == event.MouseButtonPress and self.is_shift_pressed:
                if event.button() == Qt.LeftButton:
                    self.start_pos = self.view.mapToScene(event.pos())

            elif event.type() == event.MouseMove and self.is_shift_pressed:
                if self.start_pos:
                    if self.rect_item:
                        self.scene.removeItem(self.rect_item)
                    end_pos = self.view.mapToScene(event.pos())
                    rect = QRectF(self.start_pos, end_pos).normalized()
                    self.rect_item = QGraphicsRectItem(rect)
                    pen = QPen(QColor("red"))
                    pen.setStyle(Qt.DotLine)
                    self.rect_item.setPen(pen)
                    self.scene.addItem(self.rect_item)

            elif event.type() == event.MouseButtonRelease and self.is_shift_pressed:
                if event.button() == Qt.LeftButton and self.start_pos:
                    end_pos = self.view.mapToScene(event.pos())
                    rect = QRectF(self.start_pos, end_pos).normalized()
                    original_rect = QRectF(
                        rect.x() / self.scale_factor,
                        rect.y() / self.scale_factor,
                        rect.width() / self.scale_factor,
                        rect.height() / self.scale_factor
                    )
                    self.selected_areas.append(original_rect)

                    fixed_rect = QGraphicsRectItem(rect)
                    fixed_rect.setPen(QPen(QColor("red")))
                    self.scene.addItem(fixed_rect)
                    self.rect_items.append(fixed_rect)

                    # Add label with number inside the rectangle
                    index = len(self.selected_areas) - 1
                    label = self.scene.addText(f"{index + 1}")
                    label.setDefaultTextColor(QColor("blue"))
                    label.setFont(QFont("Arial", 12))
                    label.setPos(rect.x(), rect.y())

                    # Add to rect_list
                    item_widget = QWidget()
                    item_layout = QHBoxLayout()
                    item_layout.setContentsMargins(0, 0, 0, 0)
                    item_layout.setSpacing(5)

                    label_text = QLabel(f"Frame №{index + 1}")
                    remove_button = QPushButton(f"Remove №{index + 1}")
                    remove_button.clicked.connect(lambda _, i=index: self.remove_rect(i))

                    item_layout.addWidget(label_text)
                    item_layout.addWidget(remove_button)
                    item_widget.setLayout(item_layout)

                    list_item = QListWidgetItem()
                    self.rect_list.addItem(list_item)
                    self.rect_list.setItemWidget(list_item, item_widget)

                    self.start_pos = None
                    self.rect_item = None

        return super().eventFilter(source, event)

    def remove_rect(self, index):
        if 0 <= index < len(self.rect_items):
            # Remove from scene
            self.scene.removeItem(self.rect_items[index])
            del self.rect_items[index]

            # Remove from data
            del self.selected_areas[index]

            # Refresh list
            self.show_page()

    def save_template(self):
        template_data = [
            {
                "page": self.current_page_index,
                "coordinates": [
                    {
                        "x": rect.x(),
                        "y": rect.y(),
                        "width": rect.width(),
                        "height": rect.height()
                    }
                    for rect in self.selected_areas
                ]
            }
        ]

        file_path, _ = QFileDialog.getSaveFileName(self, "Save Template", "", "JSON Files (*.json)")
        if file_path:
            with open(file_path, "w") as f:
                json.dump(template_data, f, indent=4)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFMarkupTool()
    window.show()
    sys.exit(app.exec_())
