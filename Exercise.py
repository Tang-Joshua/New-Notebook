import sys
import os
import win32com.client as win32
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QScrollArea, QLabel 
)
from PyQt6.QtCore import Qt, QTimer

from PyQt6.QtGui import QDrag

textme = "Wow"

class ExcelCellMarkerApp(QMainWindow):

    

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Cell Marker")
        self.setGeometry(100, 100, 800, 500)

        

        selector_widget = QWidget(self)
        self.setCentralWidget(selector_widget)

        Label_it = QLabel("Click me if you love it.", selector_widget)

        Buttondd = QPushButton("Push me", selector_widget)

        Button_two = QPushButton("Visible", selector_widget)

        Button_three = QPushButton("Third", selector_widget)

        Button_four = QPushButton("Pass the message", selector_widget)

        Buttondd.setGeometry(50,50,100,40)
        Button_two.setGeometry(50,80,100,40)
        Button_three.setGeometry(50,140,100,40)
        Button_four.setGeometry(50,180,100,40)
        Label_it.setGeometry(50,20,140,30)


        def show_message():
            QMessageBox.information(
                None,
                "Title",
                "This is the bullshit!"
            )


        Buttondd.clicked.connect(show_message)

        # Button_two.clicked.connect(lambda: Buttondd.setVisible(False))
        self.number1 = 0

        def hide_button():
            if self.number1 == 0:
                Buttondd.setVisible(False)
                self.number1 = 1
            else:
                Buttondd.setVisible(True)
                self.number1 = 0

        


        Button_two.clicked.connect(lambda: hide_button())

        Button_three.clicked.connect(self.show_box)

        Button_four.clicked.connect(self.update_text)

        # Buttondd.setVisible(False)

        # btn_remove.clicked.connect(lambda: self.remove_selector(selector_index))
        

        # self.setCentralWidget(button)

        # selector_layout.addWidget(button)
    def update_text(self):
            globals().update(textme="Cute")
            print(textme)

    def show_box(self):
        """Creates and shows a new window when button is clicked"""
        self.new_window = NewWindow()  # Create instance
        self.new_window.show()         # Show the window

        
class NewWindow(QMainWindow):
    """The new window that appears when button is clicked"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Cell Marker")
        self.setGeometry(150, 150, 600, 400)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout(central_widget)

        text = "This is a new window!"
        
        label = QLabel(textme)
        button = QPushButton("Close Me")
        
        layout.addWidget(label)
        layout.addWidget(button)
        
        button.clicked.connect(self.close)
    

app = QApplication(sys.argv)
window = ExcelCellMarkerApp()
window.show()
sys.exit(app.exec())