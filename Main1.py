import sys
import os
import win32com.client as win32
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QScrollArea, QLabel, QSplitter, QTableWidget, 
    QAbstractItemView, QHeaderView, QTableWidgetItem, QStyledItemDelegate, QStyleOptionViewItem, QTableView
)
from PyQt6.QtCore import Qt, QPropertyAnimation, pyqtProperty, QEasingCurve, QRect
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush, QStandardItemModel
from PyQt6.QtWidgets import QScroller, QScroller, QScrollerProperties, QStyle,QStyledItemDelegate, QApplication, QTableWidgetItem, QTableWidget

import string
import re

class AnimatedButton(QPushButton):
    def __init__(self, text):
        super().__init__(text)
        self._color = QColor("#6a11cb")
        # self.setFixedSize(170, 70)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet("font-size: 16px; font-weight: bold; border-radius: 12px; border: none;")
        self.update_background(self._color)

        self.animation = QPropertyAnimation(self, b"color")
        self.animation.setDuration(300)

    def update_background(self, color):
        style = f"""
            QPushButton {{
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 {color.name()}, stop:1 #00ff84
                );
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 10px 20px;
                border: none;
                border-radius: 12px;
            }}
        """
        self.setStyleSheet(style)

    def enterEvent(self, event):
        self.animate_to(QColor("#00b85c"))  # Green gradient start
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.animate_to(QColor("#6a11cb"))  # Original purple
        super().leaveEvent(event)

    def animate_to(self, target_color):
        self.animation.stop()
        self.animation.setStartValue(self._color)
        self.animation.setEndValue(target_color)
        self.animation.valueChanged.connect(self.update_background)
        self.animation.start()

    def get_color(self):
        return self._color

    def set_color(self, color):
        self._color = color
        self.update_background(color)

    color = pyqtProperty(QColor, get_color, set_color)

class ExcelStyleTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)

        # Don't connect here — model isn't set yet

    def setModel(self, model):
        super().setModel(model)

        # Now the selectionModel is valid
        self.selectionModel().selectionChanged.connect(self._refresh_selection)

    def _refresh_selection(self, selected, deselected):
        self.viewport().update()

    def mouseMoveEvent(self, event):
        super().mouseMoveEvent(event)
        self.viewport().update()

    def paintEvent(self, event):
        super().paintEvent(event)

        selection = self.selectionModel().selection()
        if not selection.isEmpty():
            selected_range = selection.first()
            top_left = self.model().index(selected_range.top(), selected_range.left())
            bottom_right = self.model().index(selected_range.bottom(), selected_range.right())

            rect_top_left = self.visualRect(top_left)
            rect_bottom_right = self.visualRect(bottom_right)

            selection_rect = rect_top_left.united(rect_bottom_right).adjusted(0, 0, -1, -1)

            painter = QPainter(self.viewport())
            pen = QPen(QColor(0, 128, 0), 2)
            painter.setPen(pen)
            painter.drawRect(selection_rect)

            # Autofill handle
            handle_size = 6
            handle_rect = QRect(
                selection_rect.right() - handle_size + 1,
                selection_rect.bottom() - handle_size + 1,
                handle_size,
                handle_size
            )
            painter.fillRect(handle_rect, QColor(0, 128, 0))

class WhiteBackgroundDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        # Disable selection background
        if option.state & QStyle.StateFlag.State_Selected:
            option.state &= ~QStyle.StateFlag.State_Selected

        # Set background to white explicitly
        painter.fillRect(option.rect, QBrush(QColor(255, 255, 255)))

        # Draw the rest of the item
        super().paint(painter, option, index)


class CustomTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.delegate = ExcelStyleDelegate(self)
        self.setItemDelegate(self.delegate)
        self.setSelectionMode(QTableView.ExtendedSelection)  # Enable multi-selection
        self.selection_start_pos = None

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            index = self.indexAt(event.pos())
            if index.isValid():
                if event.modifiers() & Qt.ControlModifier:
                    # Toggle selection for Ctrl+click
                    row, col = index.row(), index.column()
                    if (row, col) in self.delegate.selected_cells:
                        self.delegate.selected_cells.remove((row, col))
                    else:
                        self.delegate.selected_cells.add((row, col))
                else:
                    # Start new selection
                    self.selection_start_pos = event.pos()
                    self.delegate.selection_start = (index.row(), index.column())
                    self.delegate.selected_cells.clear()
                    self.delegate.selected_cells.add((index.row(), index.column()))
                self.viewport().update()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.LeftButton and self.selection_start_pos:
            index = self.indexAt(event.pos())
            if index.isValid() and self.delegate.selection_start:
                start_row, start_col = self.delegate.selection_start
                end_row, end_col = index.row(), index.column()
                self.delegate.set_selection(start_row, start_col, end_row, end_col)
                self.viewport().update()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.selection_start_pos = None
        super().mouseReleaseEvent(event)

class MainPage(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Main Page")
        self.setGeometry(100,100,800,500)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.setCentralWidget(splitter)

        selector_widget = QWidget(self)
        selector_widget.setStyleSheet("background-color: #f8f9fa;")
        layout = QVBoxLayout()
        selector_widget.setLayout(layout)

        button = AnimatedButton("To truncate long text with an ellipsis")
        button.setMinimumWidth(80)
        button2 = AnimatedButton("Click Me 2")

        button.setSizePolicy(button.sizePolicy().horizontalPolicy(), button.sizePolicy().verticalPolicy())
        button2.setSizePolicy(button2.sizePolicy().horizontalPolicy(), button2.sizePolicy().verticalPolicy())


        layout.addWidget(button)
        layout.addWidget(button2)
        layout.addStretch()

        content = QWidget()
        content_layout = QVBoxLayout(content)

        tools_for_table = QWidget()
        tft_layout = QVBoxLayout(tools_for_table)
        tools_for_table.setStyleSheet("background-color: #f8f9fa;")


        button3 = AnimatedButton("To truncate long text with an ellipsis")


        tft_layout.addWidget(button3)

         # Create table
        self.table_widget = QTableWidget(60, 10)  # 20 rows, 10 columns

        # Disable editing
        self.table_widget.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)

        # Single item selection
        self.table_widget.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)

        # Stretch headers
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_widget.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


        # Set horizontal headers (A-Z)
        column_labels = list(string.ascii_uppercase[:self.table_widget.columnCount()])
        self.table_widget.setHorizontalHeaderLabels(column_labels)

        # Set vertical headers (1-n)
        row_labels = [str(i+1) for i in range(self.table_widget.rowCount())]
        self.table_widget.setVerticalHeaderLabels(row_labels)

        
        def evaluate_formula(formula):
            # Regular expression to match cell references (e.g., A1, C3, etc.)
            pattern = re.compile(r"([A-Z]+)(\d+)")
            matches = pattern.findall(formula)

            if not matches:
                return None  # Return None if no valid cell references are found

            total = 0
            for match in matches:
                column = match[0]
                row = int(match[1]) - 1  # Convert to 0-based index

                # Find the column index (A -> 0, B -> 1, etc.)
                col_index = string.ascii_uppercase.index(column)

                # Get the value in the cell (assume the cell contains a number)
                try:
                    item = self.table_widget.item(row, col_index)
                    if item and item.text().isdigit():  # Check if the cell contains a valid number
                        total += int(item.text())
                except IndexError:
                    pass  # Handle out-of-bounds cells if necessary

            return total

        def on_cell_edit(cell):
            row = cell.row()
            column = cell.column()

            item = self.table_widget.item(row, column)
            if item is None:
                return

            cell_text = item.text()

            if cell_text.startswith('='):
                # Strip the "=" sign to get the formula part
                formula = cell_text[1:]

                result = evaluate_formula(formula)
                if result is not None:
                    # Set the result back into the current cell
                    item.setText(str(result))

        # Connect the edit signal to the slot
        self.table_widget.itemChanged.connect(on_cell_edit)


        # self.table_widget.setItemDelegate(ExcelStyleDelegate(self.table_widget))
        self.table_widget = ExcelStyleTableView(self)
        model = QStandardItemModel(60, 10)
        self.table_widget.setModel(model)
        self.table_widget.setItemDelegate(WhiteBackgroundDelegate(self.table_widget))

        # Then connect after model is set
        self.table_widget.selectionModel().selectionChanged.connect(
            self.table_widget._refresh_selection
        )


        def show_cell_location(index):
                    row = index.row()
                    column = index.column()

                    # Convert column index to a letter (A-Z)
                    column_letter = column_labels[column] if column < len(column_labels) else f"Column {column + 1}"

                    # Row is numbered from 1 (not 0-based)
                    row_number = row + 1

                    # Create the message box showing the cell location
                    cell_location = f"{column_letter}{row_number}"
                    QMessageBox.information(self.table_widget, "Cell Location", f"You clicked on cell {cell_location}")        

        # ///////////////////////////////////////////////////////////////////////////////////////////

        content_layout.addWidget(tools_for_table)
        content_layout.addWidget(self.table_widget)

        splitter.addWidget(selector_widget)
        splitter.addWidget(content)
        

        splitter.setSizes([200, 600])
        # selector_widget.setLayout(layout)






if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainPage()
    window.show()
    sys.exit(app.exec())