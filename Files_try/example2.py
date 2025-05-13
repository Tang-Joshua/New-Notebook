import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableView, QMenu, QToolButton, QAbstractItemView, QStandardItemModel
    , QStandardItem
)
from PyQt6.QtCore import Qt, QPoint
from PyQt6.QtGui import QAction


class AutoFillTableView(QTableView):
    def __init__(self):
        super().__init__()
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)

        # Dropdown button
        self.autofill_button = QToolButton(self)
        self.autofill_button.setText("⮟")
        self.autofill_button.setVisible(False)
        self.autofill_button.setFixedSize(20, 20)
        self.autofill_button.clicked.connect(self.show_autofill_menu)

        # Dropdown menu
        self.autofill_menu = QMenu(self)
        for option in ["Copy Cells", "Fill Series", "Fill Formatting Only", "Fill Without Formatting", "Flash Fill"]:
            self.autofill_menu.addAction(option)
        self.autofill_menu.triggered.connect(self.autofill_option_selected)

    def autofill_finished(self):
        """Call this after autofill to show the dropdown."""
        indexes = self.selectionModel().selectedIndexes()
        if not indexes:
            return

        # Get bottom-right corner of selection
        bottom_right = max(indexes, key=lambda i: (i.row(), i.column()))
        rect = self.visualRect(bottom_right)

        # Move button to bottom-right
        dropdown_pos = QPoint(rect.right() + 5, rect.bottom() + 5)
        self.autofill_button.move(dropdown_pos)
        self.autofill_button.setVisible(True)

    def show_autofill_menu(self):
        self.autofill_menu.exec(self.autofill_button.mapToGlobal(QPoint(0, self.autofill_button.height())))

    def autofill_option_selected(self, action: QAction):
        print(f"Selected: {action.text()}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.table = AutoFillTableView()
        model = QStandardItemModel(10, 5)
        for row in range(10):
            for col in range(5):
                model.setItem(row, col, QStandardItem(str((row + 1) * (col + 1))))
        self.table.setModel(model)

        self.setCentralWidget(self.table)

        # Simulate autofill
        self.table.selectRow(0)
        self.table.selectRow(1)
        self.table.autofill_finished()


app = QApplication(sys.argv)
window = MainWindow()
window.resize(600, 400)
window.show()
sys.exit(app.exec())
