import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget,
    QPushButton, QTableView, QHeaderView, QSplitter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem

class MergeTableLogic:
    def __init__(self, table_view: QTableView, model: QStandardItemModel):
        self.table_view = table_view
        self.model = model

    def merge_selected_cells(self):
        selected = self.table_view.selectedIndexes()
        if not selected:
            return

        rows = sorted(set(idx.row() for idx in selected))
        cols = sorted(set(idx.column() for idx in selected))

        top_row = rows[0]
        left_col = cols[0]
        row_span = len(rows)
        col_span = len(cols)

        self.table_view.setSpan(top_row, left_col, row_span, col_span)

        for r in rows:
            for c in cols:
                if r != top_row or c != left_col:
                    self.model.setItem(r, c, QStandardItem(""))  # Clear merged cells
