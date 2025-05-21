from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QTableWidget, QTableWidgetItem
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QScrollArea, QLabel, QSplitter, QTableWidget, 
    QAbstractItemView, QHeaderView, QTableWidgetItem, QStyledItemDelegate, QStyleOptionViewItem, QTableView
)
from PyQt6.QtCore import Qt, QPropertyAnimation, pyqtProperty, QEasingCurve, QRect, QTimer
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush, QStandardItemModel, QStandardItem, QWheelEvent, QCursor
from PyQt6.QtWidgets import  QStyle,QStyledItemDelegate, QApplication, QTableWidgetItem, QTableWidget, QMenu
from Main_File.Formating_toolbar import ExcelToolbar 

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.table = QTableWidget(5, 5)
        self.toolbar = ExcelToolbar()
        self.toolbar.bold_btn.clicked.connect(self.apply_formatting)
        self.toolbar.italic_btn.clicked.connect(self.apply_formatting)
        self.toolbar.underline_btn.clicked.connect(self.apply_formatting)
        self.toolbar.h_align_group.buttonClicked.connect(self.apply_formatting)
        self.toolbar.v_align_group.buttonClicked.connect(self.apply_formatting)

        layout = QVBoxLayout()
        layout.addWidget(self.toolbar)
        layout.addWidget(self.table)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def apply_formatting(self):
        fmt = self.toolbar.get_formatting_state()
        for item in self.table.selectedItems():
            font = item.font()
            font.setBold(fmt["bold"])
            font.setItalic(fmt["italic"])
            font.setUnderline(fmt["underline"])
            item.setFont(font)

            alignment = Qt.AlignmentFlag()
            if fmt["h_align"]:
                alignment |= fmt["h_align"]
            if fmt["v_align"]:
                alignment |= fmt["v_align"]

            if alignment:
                item.setTextAlignment(alignment)

if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
