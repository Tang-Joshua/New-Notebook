from PyQt6.QtWidgets import (QApplication, QTableWidget, QTableWidgetItem, 
                            QStyledItemDelegate, QStyleOptionViewItem)
from PyQt6.QtCore import (Qt, QPropertyAnimation, QEasingCurve, QPoint, 
                         QTimer, pyqtProperty)
from PyQt6.QtGui import QWheelEvent
import math


class SmoothScrollTableWidget(QTableWidget):
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

        scroll_bar.setValue(int(new_pos))

        # Apply tuned friction: smoother, but not slippery
        self._velocity *= 0.88  # The key value for "not too icy"


# Usage example
app = QApplication([])

table = SmoothScrollTableWidget(100, 5)
table.setColumnWidth(0, 150)

# Fill with sample data
for row in range(100):
    for col in range(5):
        table.setItem(row, col, QTableWidgetItem(f"Row {row+1}, Col {col+1}"))
    table.setRowHeight(row, 30)  # Consistent row height

table.resize(800, 500)
table.show()
app.exec()