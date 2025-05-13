import sys
import os
import win32com.client as win32
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QScrollArea, QLabel, QSplitter, QTableWidget, 
    QAbstractItemView, QHeaderView, QTableWidgetItem, QStyledItemDelegate, QStyleOptionViewItem, QTableView
)
from PyQt6.QtCore import Qt, QPropertyAnimation, pyqtProperty, QEasingCurve, QRect, QTimer
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush, QStandardItemModel, QStandardItem, QWheelEvent, QCursor
from PyQt6.QtWidgets import  QStyle,QStyledItemDelegate, QApplication, QTableWidgetItem, QTableWidget

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
        self.my_model = None  # This replaces 'self.model'
        self.middle_mouse_pressed = False
        self.middle_click_position = None
        self.middle_scroll_timer = QTimer(self)
        self.middle_scroll_timer.timeout.connect(self.auto_scroll_update)
        self.setMouseTracking(True)  # Important to track mouse without pressing buttons

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setVerticalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)

        self._velocity = 0.0
        self._velocity_history = []
        self._last_wheel_time = 0

        self._animation_timer = QTimer(self)
        self._animation_timer.timeout.connect(self._update_scroll)
        self._animation_timer.start(16)  # ~60 FPS

        # ✅ Add these lines for middle mouse auto-scroll:
        self.middle_mouse_pressed = False
        self.middle_click_position = None
        self.middle_scroll_timer = QTimer(self)
        self.middle_scroll_timer.timeout.connect(self.auto_scroll_update)
        self.setMouseTracking(True)

        self.setStyleSheet("""
            QScrollBar:vertical {
                width: 10px;
                background: transparent;
            }
            QScrollBar::handle:vertical {
                background: #a0a0a0;
                min-height: 30px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }

            QScrollBar:horizontal {
                height: 10px;
                background: transparent;
            }
            QScrollBar::handle:horizontal {
                background: #a0a0a0;
                min-width: 30px;
                border-radius: 5px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
        """)


    def wheelEvent(self, event: QWheelEvent):
        now = event.timestamp()
        delta = event.angleDelta().y()

        self._velocity_history.append((now, delta))
        self._velocity_history = [(t, d) for t, d in self._velocity_history if now - t <= 100]

        # Weighted average of recent scroll deltas
        weighted_sum = 0
        weight_total = 0
        for t, d in self._velocity_history:
            weight = 1.0 - (now - t) / 100.0
            weight = max(0, weight)
            weighted_sum += d * weight
            weight_total += weight

        avg_delta = weighted_sum / weight_total if weight_total else 0

        # Convert delta to velocity (tuned scale)
        new_velocity = avg_delta * 0.1

        # Blend with previous velocity (natural kick)
        self._velocity = self._velocity * 0.2 + new_velocity * 0.8

        event.accept()

    def _update_scroll(self):
        scroll_bar = self.verticalScrollBar()
        if abs(self._velocity) < 0.2:
            self._velocity = 0.0
            return

        # Apply movement
        new_pos = scroll_bar.value() - self._velocity

        # Clamp within bounds and bounce lightly
        if new_pos < scroll_bar.minimum():
            new_pos = scroll_bar.minimum()
            self._velocity = 0.0  # Stop completely at top
        elif new_pos > scroll_bar.maximum():
            new_pos = scroll_bar.maximum()
            self._velocity = 0.0  # Stop completely at bottom

            # Check if we are at the last row
            if self.my_model is not None:
                self.add_row_if_needed()

        scroll_bar.setValue(int(new_pos))

        # Apply tuned friction: smoother, but not slippery
        self._velocity *= 0.88  # The key value for "not too icy"

    def add_row_if_needed(self):
        row_count = self.my_model.rowCount()

        # ✅ Prevent adding more than 1000 rows
        if row_count >= 1000:
            return

        self.my_model.insertRow(row_count)
        self.my_model.setVerticalHeaderItem(row_count, QStandardItem(str(row_count + 1)))

        # Initialize each new cell in the new row (optional)
        for col in range(self.my_model.columnCount()):
            self.my_model.setItem(row_count, col, QStandardItem(""))


    def add_row_if_needed(self):
        row_count = self.my_model.rowCount()
        if row_count >= 1000:  # Limit to 1000 rows
            return
        self.my_model.insertRow(row_count)
        self.my_model.setVerticalHeaderItem(row_count, QStandardItem(str(row_count + 1)))

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Right:
            current_index = self.currentIndex()
            if current_index.isValid():
                row = current_index.row()
                col = current_index.column()
                if col == self.my_model.columnCount() - 1:
                    self.add_column_if_needed()
                    # Set focus to the newly added column
                    self.setCurrentIndex(self.my_model.index(row, col + 1))
        elif event.key() == Qt.Key.Key_Down:
            current_index = self.currentIndex()
            if current_index.isValid():
                row = current_index.row()
                col = current_index.column()
                if row == self.my_model.rowCount() - 1:
                    self.add_row_if_needed()
                    # Set focus to the newly added column
                    self.setCurrentIndex(self.my_model.index(row + 1, col))

        super().keyPressEvent(event)

    def add_column_if_needed(self):
        col_count = self.my_model.columnCount()
        # ✅ Prevent adding more than 1000 columns
        if col_count >= 1000:
            QMessageBox.warning(self, "Limit Reached", "Maximum number of columns (1000) reached.")
            return
        self.my_model.insertColumn(col_count)

        # Generate column label (A, B, ..., Z, AA, etc.)
        col_label = ""
        index = col_count
        while index >= 0:
            col_label = chr(index % 26 + ord('A')) + col_label
            index = index // 26 - 1
            if index < 0:
                break

        self.my_model.setHorizontalHeaderItem(col_count, QStandardItem(col_label))

        # Optionally initialize cells
        for row in range(self.my_model.rowCount()):
            self.my_model.setItem(row, col_count, QStandardItem(""))


    def setModelWithHeaders(self, rows, cols):
        self.my_model = QStandardItemModel(rows, cols)

        # Set column headers A, B, C...
        headers = []
        for i in range(cols):
            result = ""
            col = i
            while True:
                result = chr(col % 26 + ord('A')) + result
                col = col // 26 - 1
                if col < 0:
                    break
            headers.append(result)
        self.my_model.setHorizontalHeaderLabels(headers)

        # Set row headers 1, 2, ...
        for row in range(rows):
            self.my_model.setVerticalHeaderItem(row, QStandardItem(str(row + 1)))

        self.setModel(self.my_model)

        # Connect formula handler
        self.my_model.dataChanged.connect(self.handle_formula)

    def handle_formula(self, topLeft, bottomRight, roles):
        for row in range(topLeft.row(), bottomRight.row() + 1):
            for col in range(topLeft.column(), bottomRight.column() + 1):
                item = self.my_model.item(row, col)
                if not item:
                    continue
                text = item.text().strip()
                if text.upper().startswith("=SUM("):
                    inside = text[5:]
                    if inside.endswith(")"):
                        inside = inside[:-1]

                    cells = [ref.strip() for ref in inside.split(',')]
                    total = 0
                    for cell_ref in cells:
                        match = re.match(r"([A-Z]+)(\d+)", cell_ref)
                        if match:
                            col_letters, row_str = match.groups()
                            col_index = self.column_letters_to_index(col_letters)
                            row_index = int(row_str) - 1
                            target_item = self.my_model.item(row_index, col_index)
                            if target_item and target_item.text().isdigit():
                                total += int(target_item.text())

                    item.setText(str(total))



    def column_letters_to_index(self, letters):
        index = 0
        for char in letters:
            index = index * 26 + (ord(char.upper()) - ord('A') + 1)
        return index - 1
    def setModel(self, model):
        super().setModel(model)

        # Now the selectionModel is valid
        self.selectionModel().selectionChanged.connect(self._refresh_selection)

    def _refresh_selection(self, selected, deselected):
        self.viewport().update()

    def mouseMoveEvent(self, event):
        super().mouseMoveEvent(event)
        self.viewport().update()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self.middle_mouse_pressed = True
            self.middle_click_position = event.globalPosition().toPoint()  # or event.position() in older PyQt
            self.middle_scroll_timer.start(16)  # ~60 FPS
            # Optional: change cursor or display visual cue
            self.setCursor(Qt.CursorShape.SizeAllCursor)
        else:
            super().mousePressEvent(event)
    
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self.middle_mouse_pressed = False
            self.middle_scroll_timer.stop()
            self.unsetCursor()
        else:
            super().mouseReleaseEvent(event)

    def auto_scroll_update(self):
        if not self.middle_mouse_pressed or not self.middle_click_position:
            return

        cursor_pos = self.mapFromGlobal(QCursor.pos())
        delta = cursor_pos - self.middle_click_position

        # Control how sensitive the movement is
        scroll_speed_factor = 0.2

        # Scroll vertically
        v_scroll = self.verticalScrollBar()
        new_v = v_scroll.value() + int(delta.y() * scroll_speed_factor)
        v_scroll.setValue(new_v)

        # If we're near bottom, add a new row
        if v_scroll.value() >= v_scroll.maximum() - 5:
            model = self.model()
            if isinstance(model, QStandardItemModel):
                current_cols = model.columnCount()
                model.insertRow(model.rowCount())
                model.setVerticalHeaderItem(model.rowCount() - 1, QStandardItem(str(model.rowCount())))
                for col in range(current_cols):
                    model.setItem(model.rowCount() - 1, col, QStandardItem(""))

        # Scroll horizontally
        h_scroll = self.horizontalScrollBar()
        new_h = h_scroll.value() + int(delta.x() * scroll_speed_factor)
        h_scroll.setValue(new_h)

        # If we're near right, add a new column
        if h_scroll.value() >= h_scroll.maximum() - 5:
            model = self.model()
            if isinstance(model, QStandardItemModel):
                current_rows = model.rowCount()
                model.insertColumn(model.columnCount())

                # Generate next header letter (A, B, ..., AA, AB, etc.)
                new_col_index = model.columnCount() - 1
                result = ""
                col = new_col_index
                while True:
                    result = chr(col % 26 + ord('A')) + result
                    col = col // 26 - 1
                    if col < 0:
                        break
                model.setHorizontalHeaderItem(new_col_index, QStandardItem(result))

                for row in range(current_rows):
                    model.setItem(row, new_col_index, QStandardItem(""))





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
        
        self.table_widget = ExcelStyleTableView(self)
        self.table_widget.setModelWithHeaders(30, 30)

        self.table_widget.setItemDelegate(WhiteBackgroundDelegate(self.table_widget))

        # Connect the selectionModel after model is set
        self.table_widget.selectionModel().selectionChanged.connect(
            self.table_widget._refresh_selection
        )
        

        content_layout.addWidget(tools_for_table)
        content_layout.addWidget(self.table_widget)

        splitter.addWidget(selector_widget)
        splitter.addWidget(content)

        splitter.setSizes([200, 600])




if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainPage()
    window.show()
    sys.exit(app.exec())