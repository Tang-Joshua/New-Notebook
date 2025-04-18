import sys
import xlwings as xw
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QTextEdit, QMessageBox)

class ExcelJokeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel JOKE Function Creator")
        self.setGeometry(100, 100, 600, 400)
        
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("<h1>Excel JOKE Function GUI</h1>")
        title.setStyleSheet("margin-bottom: 20px;")
        layout.addWidget(title)
        
        # Input fields
        input_layout = QVBoxLayout()
        
        # Range input
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("Excel Range (e.g., A1:B2):"))
        self.range_input = QLineEdit()
        range_layout.addWidget(self.range_input)
        input_layout.addLayout(range_layout)
        
        # Individual cells
        cell_layout = QHBoxLayout()
        cell_layout.addWidget(QLabel("Cells (comma separated):"))
        self.cells_input = QLineEdit()
        cell_layout.addWidget(self.cells_input)
        input_layout.addLayout(cell_layout)
        
        # Output cell
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Output Cell:"))
        self.output_input = QLineEdit("A1")
        output_layout.addWidget(self.output_input)
        input_layout.addLayout(output_layout)
        
        layout.addLayout(input_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        self.run_btn = QPushButton("Run JOKE in Excel")
        self.run_btn.clicked.connect(self.run_joke)
        button_layout.addWidget(self.run_btn)
        
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_inputs)
        button_layout.addWidget(self.clear_btn)
        
        layout.addLayout(button_layout)
        
        # Log/Output
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)
        
        main_widget.setLayout(layout)
        
        # Initialize Excel connection
        self.app = None
        self.wb = None
        self.connect_to_excel()
    
    def connect_to_excel(self):
        try:
            self.app = xw.apps.active
            if not self.app:
                self.app = xw.App(visible=True)
            self.wb = self.app.books.active
            self.log.append("Connected to Excel successfully!")
        except Exception as e:
            self.log.append(f"Error connecting to Excel: {str(e)}")
    
    def run_joke(self):
        try:
            if not self.wb:
                self.connect_to_excel()
            
            # Get input values
            cells = [cell.strip() for cell in self.cells_input.text().split(",") if cell.strip()]
            range_ref = self.range_input.text().strip()
            output_cell = self.output_input.text().strip()
            
            if not cells and not range_ref:
                QMessageBox.warning(self, "Input Error", "Please enter either cells or a range")
                return
            
            # Create the JOKE formula
            cell_refs = ",".join(cells)
            if range_ref:
                if cell_refs:
                    cell_refs += ","
                cell_refs += range_ref
            
            formula = f"=JOKE({cell_refs})"
            
            # Write to Excel
            sheet = self.wb.sheets.active
            sheet[output_cell].formula = formula
            
            self.log.append(f"Added JOKE function to cell {output_cell}: {formula}")
            
        except Exception as e:
            self.log.append(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
    
    def clear_inputs(self):
        self.range_input.clear()
        self.cells_input.clear()
        self.output_input.setText("A1")
        self.log.clear()
    
    def closeEvent(self, event):
        # Clean up Excel connection
        if hasattr(self, 'wb'):
            self.wb.close()
        if hasattr(self, 'app'):
            self.app.quit()
        event.accept()

# Register the JOKE function with Excel
@xw.func
def JOKE(*args):
    """A custom Excel function that returns a joke response"""
    total = 0
    for arg in args:
        if isinstance(arg, (list, tuple)):  # Handle range inputs
            for item in arg:
                if isinstance(item, (int, float)):
                    total += item
        elif isinstance(arg, (int, float)):
            total += arg
    
    return f"😂 The joke's on you! Total: {total}"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelJokeApp()
    window.show()
    sys.exit(app.exec())




# This file is for excel (joke(A1, B2 etc))