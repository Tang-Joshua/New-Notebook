import sys
from PyQt6.QtWidgets import (
    QApplication, QTableView, QStyledItemDelegate, QStyleOptionViewItem,
    QHeaderView, QAbstractItemView
)
from PyQt6.QtGui import (
    QStandardItemModel, QStandardItem, QBrush, QColor, QPen
)
from PyQt6.QtCore import Qt, QRect


class ExcelStyleDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._drawn = set()  # Track which selections were already drawn

    def paint(self, painter, option, index):
        option2 = QStyleOptionViewItem(option)
        if option.state & option.state & QStyleOptionViewItem.StateFlag.State_Selected:
            option2.state &= ~QStyleOptionViewItem.StateFlag.State_Selected

        painter.fillRect(option.rect, QBrush(QColor(255, 255, 255)))
        super().paint(painter, option2, index)

        # Only draw the selection once per selected region
        view = self.parent()
        selection_model = view.selectionModel()
        if not selection_model:
            return

        selection = selection_model.selection()
        for i, sel_range in enumerate(selection):
            key = (sel_range.topLeft().row(), sel_range.topLeft().column())
            if key in self._drawn:
                continue  # Already drawn this region

            rect = view.visualRect(sel_range.topLeft()).united(
                view.visualRect(sel_range.bottomRight())
            ).adjusted(1, 1, -1, -1)

            pen = QPen(QColor(0, 128, 0), 2)
            painter.setPen(pen)
            painter.drawRect(rect)

            # Autofill handle
            handle_size = 6
            handle_rect = QRect(
                rect.right() - handle_size,
                rect.bottom() - handle_size,
                handle_size,
                handle_size
            )
            painter.setBrush(QBrush(QColor(0, 128, 0)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRect(handle_rect)

            self._drawn.add(key)  # Mark this selection as drawn

    def initStyleOption(self, option, index):
        # Clear drawn regions before painting a new set
        self._drawn.clear()
        super().initStyleOption(option, index)


def main():
    app = QApplication(sys.argv)

    view = QTableView()
    model = QStandardItemModel(10, 10)
    for row in range(10):
        for col in range(10):
            model.setItem(row, col, QStandardItem(f"{row},{col}"))
    view.setModel(model)

    view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
    view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
    view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    view.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    delegate = ExcelStyleDelegate(view)
    view.setItemDelegate(delegate)

    view.setWindowTitle("Excel-Style Selection")
    view.resize(600, 400)
    view.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
