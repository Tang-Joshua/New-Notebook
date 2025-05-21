from PyQt6.QtWidgets import QTableView, QStyledItemDelegate, QTextEdit
from PyQt6.QtGui import QTextOption
from PyQt6.QtCore import Qt

class ExcelStyleTableView2(QTableView):
    def __init__(self):
        super().__init__()
        self.setWordWrap(True)
        self.setItemDelegate(WrapTextDelegate())

    def setModel(self, model):
        super().setModel(model)
        self.model().layoutChanged.connect(self.apply_merges)
        self.apply_merges()

    def apply_merges(self):
        self.clearSpans()
        for top_row, left_col, row_span, col_span in self.model().merged_cells:
            if row_span == 1 and col_span == 1:
                continue  # Skip single-cell spans
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
