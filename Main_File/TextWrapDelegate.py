from PyQt6.QtWidgets import QStyledItemDelegate
from PyQt6.QtGui import QFontMetrics, QTextOption
from PyQt6.QtCore import Qt, QSize

class TextWrapDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)

    def sizeHint(self, option, index):
        text = index.data(Qt.ItemDataRole.DisplayRole)
        if not text:
            return super().sizeHint(option, index)

        font = option.font
        fm = QFontMetrics(font)
        width = option.rect.width()

        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapMode.WordWrap)

        # Estimate height for word-wrapped text
        lines = fm.boundingRect(0, 0, width, 1000, Qt.TextFlag.TextWordWrap, text)
        height = lines.height()
        return QSize(width, height + 4)  # Add a small buffer
