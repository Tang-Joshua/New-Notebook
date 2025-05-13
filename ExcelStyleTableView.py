
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QScrollArea, QLabel, QSplitter, QTableWidget, 
    QAbstractItemView, QHeaderView, QTableWidgetItem, QStyledItemDelegate, QStyleOptionViewItem, QTableView
)
from PyQt6.QtCore import Qt, QPropertyAnimation, pyqtProperty, QEasingCurve, QRect, QTimer
from PyQt6.QtGui import QPainter, QColor, QPen, QBrush, QStandardItemModel, QStandardItem, QWheelEvent, QCursor
from PyQt6.QtWidgets import  QStyle,QStyledItemDelegate, QApplication, QTableWidgetItem, QTableWidget, QMenu

import re



class ExcelStyleTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.my_model = None  # This replaces 'self.model'
        self.middle_mouse_pressed = False
        self.middle_click_position = None
        self.middle_scroll_timer = QTimer(self)
        self.middle_scroll_timer.timeout.connect(self.auto_scroll_update)
        self.setMouseTracking(True)  # Important to track mouse without pressing buttons

        self.autofill_dragging = False
        self.autofill_direction = None
        self.autofill_start_index = None
        self.autofill_end_index = None

        self.keep_preview_rect = None            # For visual rectangle after release
        self.keep_preview_indexes = []           # For tracking pasted cells

        self.dash_offset = 0.0
        self.dash_timer = QTimer(self)
        self.dash_timer.timeout.connect(self.update_dash_animation)
        self.dash_timer.start(100)  # Lower = faster animation

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Smooth scrolling
        self.setVerticalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)

        # Animated scroll state
        self._velocity = 0.0
        self._velocity_history = []
        self._last_wheel_time = 0
        self._animation_timer = QTimer(self)
        self._animation_timer.timeout.connect(self._update_scroll)
        self._animation_timer.start(16)  # ~60 FPS

        # Middle mouse auto-scroll
        self.middle_mouse_pressed = False
        self.middle_click_position = None
        self.middle_scroll_timer = QTimer(self)
        self.middle_scroll_timer.timeout.connect(self.auto_scroll_update)
        self.setMouseTracking(True)

        # Autofill state
        self.autofill_dragging = False
        self.autofill_direction = None
        self.autofill_start_index = None
        self.autofill_end_index = None
        self.selected_values = []  # Store selected cells for autofill
        self.handle_size = 6

        self.keep_preview_rect = None            # For visual rectangle after release
        self.keep_preview_indexes = []           # For tracking pasted cells

        self.dash_offset = 0.0
        self.dash_timer = QTimer(self)
        self.dash_timer.timeout.connect(self.update_dash_animation)
        self.dash_timer.start(100)  # Lower = faster animation

        # Autofill state
        self.autofill_dragging = False
        self.autofill_direction = None
        self.autofill_start_index = None
        self.autofill_end_index = None
        self.selected_values = []  # Store selected cells for autofill
        self.handle_size = 6

        # Track the last number used in Fill Series
        self.last_series_number = 2  # Initialize to 2 so first Fill Series starts at 3




        # Custom scrollbar styling
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

    def update_dash_animation(self):
        self.dash_offset += 1.0
        if self.dash_offset >= 100.0:  # Prevent overflow
            self.dash_offset = 0.0
        self.viewport().update()

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

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self.middle_mouse_pressed = True
            self.middle_click_position = event.globalPosition().toPoint()
            self.middle_scroll_timer.start(16)  # ~60 FPS
            self.setCursor(Qt.CursorShape.SizeAllCursor)
            return  # Skip further processing for middle button

        elif event.button() == Qt.MouseButton.LeftButton and hasattr(self, "handle_rect"):
            if self.handle_rect.contains(event.pos()):
                self.autofill_dragging = True
                self.autofill_start_index = self.currentIndex()
                self.autofill_start_pos = event.pos()  # Save the initial position for direction detection
                return  # Skip normal selection while dragging autofill handle

        index = self.indexAt(event.pos())
        if index.isValid():
            # Clicking anywhere clears previous visual
            self.keep_preview_indexes.clear()
            self.viewport().update()

        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        # Check if mouse is over the autofill handle corner
        if hasattr(self, 'handle_rect') and self.handle_rect.contains(event.pos()):
            self.setCursor(Qt.CursorShape.CrossCursor)  # Set to plus cursor when over autofill handle
        elif self.autofill_dragging:
            self.setCursor(Qt.CursorShape.CrossCursor)  # Keep the plus cursor while dragging
        else:
            self.unsetCursor()  # Reset to default cursor when not interacting with autofill
            
        if self.autofill_dragging:
            # Calculate the movement direction
            delta_x = event.pos().x() - self.autofill_start_pos.x()
            delta_y = event.pos().y() - self.autofill_start_pos.y()

            # Cursor detection for autofill handle
            if hasattr(self, 'handle_rect') and self.handle_rect.contains(event.pos()):
                # Set cursor to plus sign when over the autofill handle
                self.setCursor(Qt.CursorShape.CrossCursor)
            elif self.autofill_dragging:
                # Keep the cross cursor while dragging
                self.setCursor(Qt.CursorShape.CrossCursor)
            else:
                self.unsetCursor()  # Reset cursor to default when not dragging

            if abs(delta_x) > abs(delta_y):
                self.autofill_direction = 'horizontal'
            else:
                self.autofill_direction = 'vertical'

            index = self.indexAt(event.pos())
            if not index.isValid():
                return

            # Update the end index based on the direction
            if self.autofill_direction == 'vertical':
                self.autofill_end_index = self.model().index(index.row(), self.autofill_start_index.column())
            elif self.autofill_direction == 'horizontal':
                self.autofill_end_index = self.model().index(self.autofill_start_index.row(), index.column())

            self.viewport().update()
            return

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self.middle_mouse_pressed = False
            self.middle_scroll_timer.stop()
            self.unsetCursor()
            return

        elif self.autofill_dragging:
            self.autofill_dragging = False

            if not self.autofill_end_index or not self.autofill_start_index:
                self.autofill_start_index = None
                self.autofill_end_index = None
                self.autofill_direction = None
                self.viewport().update()
                self.unsetCursor()
                return

            model = self.model()
            selection = self.selectionModel().selectedIndexes()
            if not selection:
                return

            # Normalize selection into a grid
            selection.sort(key=lambda index: (index.row(), index.column()))
            sel_top = min(index.row() for index in selection)
            sel_left = min(index.column() for index in selection)
            sel_bottom = max(index.row() for index in selection)
            sel_right = max(index.column() for index in selection)

            sel_rows = sel_bottom - sel_top + 1
            sel_cols = sel_right - sel_left + 1

            # Get selected values BEFORE writing
            selected_values = [
                [model.data(model.index(sel_top + r, sel_left + c)) for c in range(sel_cols)]
                for r in range(sel_rows)
            ]

            # Check if selected values are all numeric
            is_numeric = all(
                all(
                    value and value.strip() and value.strip().lstrip('-').replace('.', '', 1).isdigit()
                    for value in row
                )
                for row in selected_values
            )

            # Determine base number for series mode
            try:
                if self.autofill_direction == 'vertical':
                    last_value_str = selected_values[-1][0]
                else:
                    last_value_str = selected_values[0][-1]

                self.last_series_number = int(last_value_str)
            except (ValueError, IndexError, TypeError):
                self.last_series_number = None


            start_row = self.autofill_start_index.row()
            start_col = self.autofill_start_index.column()
            end_row = self.autofill_end_index.row()
            end_col = self.autofill_end_index.column()

            delta_row = end_row - start_row
            delta_col = end_col - start_col

            # Determine fill behavior based on content and user choice
            fill_mode = "copy"  # Default to copying
            if is_numeric:
                # Show context menu for user to choose
                menu = QMenu(self)
                copy_action = menu.addAction("Copy Cells")
                fill_series_action = menu.addAction("Fill Series")
                
                # Position the menu near the autofill handle
                pos = self.mapToGlobal(self.visualRect(self.autofill_end_index).bottomRight())
                action = menu.exec(pos)

                if action == fill_series_action:
                    fill_mode = "series"
                # If copy_action or no action selected, fill_mode remains "copy"

            # Apply the fill based on the direction and mode
            if self.autofill_direction == 'vertical':
                step = 1 if delta_row > 0 else -1
                fill_start_row = sel_bottom + 1 if step == 1 else sel_top - 1
                fill_end_row = end_row
                fill_range = list(range(fill_start_row, fill_end_row + step, step))
                num_steps = len(fill_range)

                for row_index, r in enumerate(fill_range):
                    for col_offset in range(sel_cols):
                        value_index = (row_index % sel_rows)
                        base_value = selected_values[value_index][col_offset] or ""

                        if fill_mode == "series" and self.last_series_number is not None:
                            series_value = self.last_series_number + (row_index + 1) * step
                            new_value = str(series_value)
                        else:
                            new_value = base_value

                        model.setData(model.index(r, sel_left + col_offset), new_value)


            elif self.autofill_direction == 'horizontal':
                step = 1 if delta_col > 0 else -1
                fill_start_col = sel_right + 1 if step == 1 else sel_left - 1
                fill_end_col = end_col
                fill_range = list(range(fill_start_col, fill_end_col + step, step))
                num_steps = len(fill_range)

                for col_index, c in enumerate(fill_range):
                    for row_offset in range(sel_rows):
                        value_index = (col_index % sel_cols)
                        base_value = selected_values[row_offset][value_index] or ""

                        if fill_mode == "series" and self.last_series_number is not None:
                            series_value = self.last_series_number + (col_index + 1) * step
                            new_value = str(series_value)
                        else:
                            new_value = base_value

                        model.setData(model.index(sel_top + row_offset, c), new_value)


            # Determine the area to keep visual (for the dashed outline after autofill)
            if self.autofill_end_index and self.autofill_direction:
                selection.sort(key=lambda idx: (idx.row(), idx.column()))
                sel_top = min(idx.row() for idx in selection)
                sel_bottom = max(idx.row() for idx in selection)
                sel_left = min(idx.column() for idx in selection)
                sel_right = max(idx.column() for idx in selection)

                if self.autofill_direction == 'vertical':
                    fill_top = sel_bottom + 1 if self.autofill_end_index.row() > sel_bottom else self.autofill_end_index.row()
                    fill_bottom = self.autofill_end_index.row() if self.autofill_end_index.row() > sel_bottom else sel_top - 1
                    fill_left = sel_left
                    fill_right = sel_right
                else:
                    fill_top = sel_top
                    fill_bottom = sel_bottom
                    fill_left = sel_right + 1 if self.autofill_end_index.column() > sel_right else self.autofill_end_index.column()
                    fill_right = self.autofill_end_index.column() if self.autofill_end_index.column() > sel_right else sel_left - 1

                # Save indexes for repaint
                self.keep_preview_indexes = [
                    model.index(r, c)
                    for r in range(min(fill_top, fill_bottom), max(fill_top, fill_bottom) + 1)
                    for c in range(min(fill_left, fill_right), max(fill_left, fill_right) + 1)
                ]

            # Reset autofill state
            self.autofill_start_index = None
            self.autofill_end_index = None
            self.autofill_direction = None
            self.autofill_dragging = False
            self.viewport().update()

            # Reset cursor to default on release
            self.unsetCursor()
            return

        # Default behavior
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

        selection = self.selectionModel().selectedIndexes()
        if not selection:
            return

        model = self.model()
        selection.sort(key=lambda idx: (idx.row(), idx.column()))
        sel_top = min(idx.row() for idx in selection)
        sel_bottom = max(idx.row() for idx in selection)
        sel_left = min(idx.column() for idx in selection)
        sel_right = max(idx.column() for idx in selection)

        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Draw main selection border
        top_left_index = model.index(sel_top, sel_left)
        bottom_right_index = model.index(sel_bottom, sel_right)

        rect_top_left = self.visualRect(top_left_index)
        rect_bottom_right = self.visualRect(bottom_right_index)

        selection_rect = rect_top_left.united(rect_bottom_right).adjusted(0, 0, -1, -1)

        pen = QPen(QColor(0, 128, 0), 2)
        painter.setPen(pen)
        painter.drawRect(selection_rect)

        # Draw autofill handle
        self.handle_rect = QRect(
            selection_rect.right() - self.handle_size + 1,
            selection_rect.bottom() - self.handle_size + 1,
            self.handle_size,
            self.handle_size
        )
        painter.fillRect(self.handle_rect, QColor(0, 0, 255))

        # Store drag_rect for checking cursor position in mouseMoveEvent
        self.drag_rect = None  # Initialize to None each time

        # Draw drag selection preview if dragging
        if self.autofill_dragging and self.autofill_start_index and self.autofill_end_index:
            start_row = self.autofill_start_index.row()
            start_col = self.autofill_start_index.column()
            end_row = self.autofill_end_index.row()
            end_col = self.autofill_end_index.column()

            if self.autofill_direction == 'vertical':
                if end_row > sel_bottom:
                    drag_top = sel_bottom + 1
                    drag_bottom = end_row
                elif end_row < sel_top:
                    drag_top = end_row
                    drag_bottom = sel_top - 1
                else:
                    return  # No visual if drag within selected rows

                drag_left = sel_left
                drag_right = sel_right

            elif self.autofill_direction == 'horizontal':
                if end_col > sel_right:
                    drag_left = sel_right + 1
                    drag_right = end_col
                elif end_col < sel_left:
                    drag_left = end_col
                    drag_right = sel_left - 1
                else:
                    return  # No visual if drag within selected columns

                drag_top = sel_top
                drag_bottom = sel_bottom

            drag_top_left = self.visualRect(model.index(drag_top, drag_left))
            drag_bottom_right = self.visualRect(model.index(drag_bottom, drag_right))

            self.drag_rect = drag_top_left.united(drag_bottom_right).adjusted(0, 0, -1, -1)

            # Draw dashed rectangle for drag preview
            pen.setStyle(Qt.PenStyle.DashLine)
            pen.setColor(QColor(0, 128, 0))
            painter.setPen(pen)
            painter.drawRect(self.drag_rect)

        # Draw last autofill rect (after mouse release)
            # Persistent preview after autofill ends
        if self.keep_preview_indexes:
            preview_rects = [self.visualRect(index) for index in self.keep_preview_indexes]
            if preview_rects:
                united_rect = preview_rects[0]
                for rect in preview_rects[1:]:
                    united_rect = united_rect.united(rect)
                united_rect.adjust(0, 0, -1, -1)

                pen = QPen(QColor(0, 128, 0), 1)
                pen.setStyle(Qt.PenStyle.CustomDashLine)
                pen.setDashPattern([4, 2])
                pen.setDashOffset(self.dash_offset)

                painter.setPen(pen)
                painter.drawRect(united_rect)
   


