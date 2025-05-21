from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableView, QVBoxLayout,
    QWidget, QPushButton, QStyledItemDelegate, QTextEdit
)
from PyQt6.QtGui import QTextOption
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QRect
import sys


class MergeTableModel(QAbstractTableModel):
    def __init__(self, rows=10, columns=5):
        super().__init__()
        self.rows = rows
        self.columns = columns
        self.data_matrix = [['' for _ in range(columns)] for _ in range(rows)]
        self.merged_cells = []  # Each item: (top_row, left_col, row_span, col_span)

    def rowCount(self, parent=QModelIndex()):
        return self.rows

    def columnCount(self, parent=QModelIndex()):
        return self.columns

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            row, col = index.row(), index.column()
            return self.data_matrix[row][col]
        return None

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role == Qt.ItemDataRole.EditRole:
            row, col = index.row(), index.column()
            self.data_matrix[row][col] = value
            self.dataChanged.emit(index, index)
            return True
        return False

    def flags(self, index):
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable

    def merge_cells(self, top_row, left_col, row_span, col_span):
        self.merged_cells.append((top_row, left_col, row_span, col_span))
        self.layoutChanged.emit()


class MergeTableView(QTableView):
    def __init__(self):
        super().__init__()
        self.setWordWrap(True)
        self.setItemDelegate(WrapTextDelegate())
        self.setSpan(0, 0, 1, 1)  # Needed for initialization

    def setModel(self, model: MergeTableModel):
        super().setModel(model)
        self.model().layoutChanged.connect(self.apply_merges)

    def apply_merges(self):
        self.clearSpans()
        for top_row, left_col, row_span, col_span in self.model().merged_cells:
            self.setSpan(top_row, left_col, row_span, col_span)


class WrapTextDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QTextEdit(parent)
        editor.setWordWrapMode(QTextOption.WrapMode.WordWrap)
        return editor

    def setEditorData(self, editor, index):
        editor.setText(index.model().data(index, Qt.ItemDataRole.DisplayRole))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText(), Qt.ItemDataRole.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel-style Merge Table")
        self.resize(600, 400)

        self.table = MergeTableView()
        self.model = MergeTableModel()
        self.table.setModel(self.model)

        merge_button = QPushButton("Merge Selected Cells")
        merge_button.clicked.connect(self.merge_selected_cells)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(merge_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def merge_selected_cells(self):
        selection = self.table.selectionModel().selection()
        if selection:
            indexes = selection.indexes()
            if not indexes:
                return
            rows = [i.row() for i in indexes]
            cols = [i.column() for i in indexes]
            top_row, bottom_row = min(rows), max(rows)
            left_col, right_col = min(cols), max(cols)
            row_span = bottom_row - top_row + 1
            col_span = right_col - left_col + 1

            # Only keep text from top-left
            top_left_index = self.model.index(top_row, left_col)
            top_left_value = self.model.data(top_left_index)
            for r in range(top_row, top_row + row_span):
                for c in range(left_col, left_col + col_span):
                    if r == top_row and c == left_col:
                        continue
                    self.model.setData(self.model.index(r, c), '')

            self.model.setData(top_left_index, top_left_value)
            self.model.merge_cells(top_row, left_col, row_span, col_span)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
