import sys
from PyQt6.QtCore import Qt, QEasingCurve
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableWidget, QVBoxLayout, QWidget, QAbstractItemView, QHeaderView
import string

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set up the table widget with 60 rows and 10 columns
        self.table_widget = QTableWidget(60, 10, self)

        # Set the scroll mode to ScrollPerPixel for smooth scrolling
        self.table_widget.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table_widget.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)

        # Set up basic table scroll without QScroller
        self.table_widget.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.table_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        # Set headers
        column_labels = list(string.ascii_uppercase[:self.table_widget.columnCount()])
        self.table_widget.setHorizontalHeaderLabels(column_labels)
        row_labels = [str(i + 1) for i in range(self.table_widget.rowCount())]
        self.table_widget.setVerticalHeaderLabels(row_labels)

        # Stretch headers using QHeaderView.Stretch
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_widget.verticalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)

        # Disable editing in the table
        self.table_widget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        # Single item selection
        self.table_widget.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.table_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)

        # Set up the layout
        layout = QVBoxLayout()
        layout.addWidget(self.table_widget)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.setWindowTitle("Smooth Scrolling Table")
        self.setGeometry(100, 100, 800, 500)

        # Debug: Confirm that the table is created
        print("Table setup complete with {} rows and {} columns".format(self.table_widget.rowCount(), self.table_widget.columnCount()))

# Run the application
app = QApplication(sys.argv)
window = MainWindow()
window.show()

# Debug: Confirm if app is running
print("Application is running...")

sys.exit(app.exec())
