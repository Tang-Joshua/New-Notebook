import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget
from PyQt6.QtCore import QAbstractTableModel, QModelIndex, Qt

from ExcelStyleTableView import ExcelStyleTableView  # Import your view

class MergeTableModel(QAbstractTableModel):
    def __init__(self, rows=10, columns=5):
        super().__init__()
        self.rows = rows
        self.columns = columns
        self.data_matrix = [['Cell {}-{}'.format(r, c) for c in range(columns)] for r in range(rows)]
        self.merged_cells = []  # List of tuples: (top_row, left_col, row_span, col_span)

    def rowCount(self, parent=QModelIndex()):
        return self.rows

    def columnCount(self, parent=QModelIndex()):
        return self.columns

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            return self.data_matrix[index.row()][index.column()]
        return None

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role == Qt.ItemDataRole.EditRole:
            self.data_matrix[index.row()][index.column()] = value
            self.dataChanged.emit(index, index)
            return True
        return False

    def flags(self, index):
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable

    def merge_cells(self, top_row, left_col, row_span, col_span):
        self.merged_cells.append((top_row, left_col, row_span, col_span))
        self.layoutChanged.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Merge Cells Demo")
        self.resize(600, 400)

        self.model = MergeTableModel()
        self.table = ExcelStyleTableView()
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
        if not selection:
            return
        indexes = selection.indexes()
        if not indexes:
            return

        rows = [index.row() for index in indexes]
        cols = [index.column() for index in indexes]
        top_row, bottom_row = min(rows), max(rows)
        left_col, right_col = min(cols), max(cols)
        row_span = bottom_row - top_row + 1
        col_span = right_col - left_col + 1

        # Keep text only in top-left cell
        top_left_index = self.model.index(top_row, left_col)
        top_left_value = self.model.data(top_left_index)
        for r in range(top_row, top_row + row_span):
            for c in range(left_col, left_col + col_span):
                if r == top_row and c == left_col:
                    continue
                self.model.setData(self.model.index(r, c), '')

        self.model.setData(top_left_index, top_left_value)
        self.model.merge_cells(top_row, left_col, row_span, col_span)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
