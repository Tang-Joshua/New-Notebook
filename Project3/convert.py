import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QVBoxLayout, QWidget, 
    QPushButton, QHBoxLayout, QDockWidget, QLineEdit, QLabel
)
from PyQt6.QtGui import QTextCursor, QTextFormat, QDrag
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtCore import Qt, QMimeData
from docx import Document
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

class DraggableTextBox(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setText("Type here...")
        self.setFixedSize(150, 50)
        self.setStyleSheet("border: 1px solid black; padding: 5px;")
        self.setDragEnabled(True)

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton:
            drag = QDrag(self)
            mime_data = QMimeData()
            mime_data.setText(self.text())  # Store the text to drag
            drag.setMimeData(mime_data)
            drag.exec(Qt.DropAction.CopyAction)

class DocumentEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Document Editor (PyQt6)")
        self.setGeometry(100, 100, 800, 600)

        # Main Widget & Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # Left side: Text Editor
        editor_layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setAcceptDrops(True)  # Enable drop on the text edit
        self.text_edit.dropEvent = self.handle_drop  # Override drop event
        editor_layout.addWidget(self.text_edit)

        # Buttons for adding components
        button_layout = QHBoxLayout()

        self.add_text_btn = QPushButton("Add Text")
        self.add_text_btn.clicked.connect(self.add_text)
        button_layout.addWidget(self.add_text_btn)

        self.add_box_btn = QPushButton("Add Box")
        self.add_box_btn.clicked.connect(self.add_box)
        button_layout.addWidget(self.add_box_btn)

        # Export Buttons
        self.export_word_btn = QPushButton("Export to Word")
        self.export_word_btn.clicked.connect(self.export_to_word)
        button_layout.addWidget(self.export_word_btn)

        self.export_pdf_btn = QPushButton("Export to PDF")
        self.export_pdf_btn.clicked.connect(self.export_to_pdf)
        button_layout.addWidget(self.export_pdf_btn)

        editor_layout.addLayout(button_layout)
        main_layout.addLayout(editor_layout)

        # Right side: Sidebar with draggable text box
        sidebar_widget = QWidget()
        sidebar_layout = QVBoxLayout(sidebar_widget)
        sidebar_layout.addWidget(QLabel("Drag this box into the editor:"))
        self.draggable_box = DraggableTextBox()
        sidebar_layout.addWidget(self.draggable_box)
        sidebar_layout.addStretch()  # Push the box to the top
        main_layout.addWidget(sidebar_widget)

    def handle_drop(self, event):
        """Handle the drop event when the text box is dragged into the editor."""
        mime_data = event.mimeData()
        if mime_data.hasText():
            text = mime_data.text()
            cursor = self.text_edit.textCursor()
            # Insert the dragged text as a box (HTML div)
            cursor.insertHtml(f'<div style="border: 1px solid black; padding: 5px; width: 150px;">{text}</div>')
            event.acceptProposedAction()

    def add_text(self):
        """Insert sample text at cursor position."""
        cursor = self.text_edit.textCursor()
        cursor.insertText("This is added text.\n")

    def add_box(self):
        """Add a box (rectangle) around selected text."""
        cursor = self.text_edit.textCursor()
        if cursor.hasSelection():
            fmt = QTextFormat()
            fmt.setProperty(QTextFormat.Property.FrameBorder, 1)
            fmt.setProperty(QTextFormat.Property.FrameBorderStyle, Qt.PenStyle.SolidLine)
            fmt.setProperty(QTextFormat.Property.FrameBorderBrush, Qt.GlobalColor.black)
            cursor.insertHtml(f'<div style="border: 1px solid black; padding: 5px;">{cursor.selectedText()}</div>')
        else:
            cursor.insertHtml('<div style="border: 1px solid black; padding: 5px; width: 100px; height: 50px;">Box</div>')

    def export_to_word(self):
        """Export content to a Word (.docx) file."""
        doc = Document()
        doc.add_paragraph(self.text_edit.toPlainText())
        doc.save("output.docx")
        print("Exported to Word (output.docx)")

    def export_to_pdf(self):
        """Export content to a PDF file."""
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName("output.pdf")
        self.text_edit.document().print(printer)
        print("Exported to PDF (output.pdf)")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentEditor()
    window.show()
    sys.exit(app.exec())