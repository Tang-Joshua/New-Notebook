import sys
import os
import win32com.client as win32
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton,
    QFileDialog, QMessageBox, QHBoxLayout, QLabel, QGroupBox
)
from PyQt6.QtCore import Qt

textme = "Wow"

class ExcelCellMarkerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Cell Marker")
        self.setGeometry(100, 100, 800, 500)

        # Central Widget
        selector_widget = QWidget(self)
        self.setCentralWidget(selector_widget)

        # Layout
        main_layout = QVBoxLayout(selector_widget)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # GroupBox for Buttons
        button_group = QGroupBox("Main Controls")
        group_layout = QVBoxLayout()
        button_group.setLayout(group_layout)

        # Label
        Label_it = QLabel("Click me if you love it.")
        Label_it.setAlignment(Qt.AlignmentFlag.AlignCenter)
        Label_it.setStyleSheet("font-weight: bold; font-size: 16px;")

        # Buttons
        Buttondd = QPushButton("Push me")
        Button_two = QPushButton("Toggle Visibility")
        Button_three = QPushButton("Open New Window")
        Button_four = QPushButton("Pass the Message")

        # Add widgets to layout
        group_layout.addWidget(Label_it)
        group_layout.addWidget(Buttondd)
        group_layout.addWidget(Button_two)
        group_layout.addWidget(Button_three)
        group_layout.addWidget(Button_four)

        main_layout.addWidget(button_group)

        # Button Events
        self.number1 = 0

        Buttondd.clicked.connect(self.show_message)
        Button_two.clicked.connect(self.toggle_visibility(Buttondd))
        Button_three.clicked.connect(self.show_box)
        Button_four.clicked.connect(self.update_text)

    def show_message(self):
        QMessageBox.information(
            self,
            "Title",
            "This is the message!"
        )

    def toggle_visibility(self, target_button):
        def _toggle():
            if self.number1 == 0:
                target_button.setVisible(False)
                self.number1 = 1
            else:
                target_button.setVisible(True)
                self.number1 = 0
        return _toggle

    def update_text(self):
        global textme
        textme = "Cute"
        print(textme)

    def show_box(self):
        self.new_window = NewWindow()
        self.new_window.show()


class NewWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("New Window")
        self.setGeometry(150, 150, 600, 400)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)

        label = QLabel(textme)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-size: 18px; font-weight: bold;")

        button = QPushButton("Close Me")

        layout.addWidget(label)
        layout.addWidget(button)

        button.clicked.connect(self.close)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Apply global style
    app.setStyleSheet("""
        QMainWindow {
            background-color: #f8f9fa;
        }
        QPushButton {
            background-color: #4CAF50;
            color: white;
            padding: 10px;
            font-size: 14px;
            border: none;
            border-radius: 8px;
        }
        QPushButton:hover {
            background-color: #45a049;
        }
        QLabel {
            color: #333333;
        }
        QGroupBox {
            border: 2px solid #cccccc;
            border-radius: 10px;
            margin-top: 10px;
            padding: 10px;
            font-size: 16px;
        }
        QGroupBox:title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 3px 0 3px;
        }
    """)

    window = ExcelCellMarkerApp()
    window.show()
    sys.exit(app.exec())
