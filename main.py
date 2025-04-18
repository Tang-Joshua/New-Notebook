import sys
import os
import win32com.client as win32
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QScrollArea, QLabel
)
from PyQt6.QtCore import Qt, QTimer

from PyQt6.QtGui import QDrag

class DraggableTextBox(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_start_position = event.pos()
    
    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.MouseButton.LeftButton):
            return
        if (event.pos() - self.drag_start_position).manhattanLength() < QApplication.startDragDistance():
            return
        
        drag = QDrag(self)
        mimedata = QMimeData()
        mimedata.setText(self.text())
        drag.setMimeData(mimedata)
        
        drag.exec(Qt.DropAction.CopyAction | Qt.DropAction.MoveAction)

class ExcelCellMarkerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Cell Marker")
        self.setGeometry(100, 100, 800, 500)
        
        # Variables for Excel interaction
        self.excel_app = None
        self.workbook = None
        self.file_path = ""
        self.cell_selections = []
        self.current_selector_index = None  # Track which selector is being set
        
        # Main widgets
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        
        # File selection button
        self.btn_select_file = QPushButton("Select Excel File")
        self.btn_select_file.clicked.connect(self.select_excel_file)
        self.layout.addWidget(self.btn_select_file)
        
        # Current file label
        self.lbl_current_file = QLabel("No file selected")
        self.layout.addWidget(self.lbl_current_file)
        
        # Instructions label
        self.lbl_instructions = QLabel("After clicking 'Select Cell', click on the cell in Excel")
        self.lbl_instructions.setStyleSheet("color: blue; font-weight: bold;")
        self.lbl_instructions.hide()
        self.layout.addWidget(self.lbl_instructions)
        
        # Scroll area for cell selection buttons
        self.scroll_area = QScrollArea()
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        self.scroll_area.setWidgetResizable(True)
        self.layout.addWidget(self.scroll_area)
        
        # Button to add more cell selectors
        self.btn_add_selector = QPushButton("Add Cell Selector")
        self.btn_add_selector.clicked.connect(self.add_cell_selector)
        self.btn_add_selector.setEnabled(False)
        self.layout.addWidget(self.btn_add_selector)
        
        # Save button
        self.btn_save = QPushButton("Save Changes to Excel")
        self.btn_save.clicked.connect(self.save_changes)
        self.btn_save.setEnabled(False)
        self.layout.addWidget(self.btn_save)
        
        # Timer to check for cell selection
        self.selection_timer = QTimer(self)
        self.selection_timer.timeout.connect(self.check_excel_selection)
        
        # Add one initial selector
        self.add_cell_selector()
    
    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # Close any existing Excel instance
                if self.excel_app:
                    try:
                        self.workbook.Close(False)
                        self.excel_app.Quit()
                    except:
                        pass
                
                # Open Excel using win32com
                self.excel_app = win32.gencache.EnsureDispatch('Excel.Application')
                self.excel_app.Visible = True
                self.workbook = self.excel_app.Workbooks.Open(os.path.abspath(file_path))
                
                self.file_path = file_path
                self.lbl_current_file.setText(f"Current file: {file_path}")
                self.btn_add_selector.setEnabled(True)
                self.btn_save.setEnabled(True)
                
                QMessageBox.information(
                    self, "Success", f"Opened Excel file: {file_path}\n"
                    "Please select cells in Excel after clicking 'Select Cell' buttons."
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Failed to open Excel file: {str(e)}"
                )
    
    def add_cell_selector(self):
        selector_index = len(self.cell_selections) + 1
        selector_widget = QWidget()
        selector_layout = QHBoxLayout(selector_widget)
        
        # Button to select cell
        btn_select = QPushButton(f"Select Cell {selector_index}")
        btn_select.clicked.connect(lambda: self.start_cell_selection(selector_index))
        
        # Label to show selected cell reference
        lbl_cell_ref = QLabel("Not selected")
        lbl_cell_ref.setStyleSheet("font-weight: bold; min-width: 80px;")
        
        # Label to show cell value
        lbl_cell_value = QLabel("No value")
        lbl_cell_ref = DraggableTextBox()
        lbl_cell_value.setStyleSheet("font-style: italic; min-width: 200px;")
        
        # Button to remove this selector
        btn_remove = QPushButton("Remove")
        btn_remove.clicked.connect(lambda: self.remove_selector(selector_index))
        
        selector_layout.addWidget(btn_select)
        selector_layout.addWidget(lbl_cell_ref)
        selector_layout.addWidget(QLabel("Value:"))
        selector_layout.addWidget(lbl_cell_value)
        selector_layout.addWidget(btn_remove)
        
        self.scroll_layout.addWidget(selector_widget)
        
        # Store the reference
        self.cell_selections.append({
            "widget": selector_widget,
            "select_button": btn_select,
            "cell_ref_label": lbl_cell_ref,
            "cell_value_label": lbl_cell_value,
            "cell_ref": None,
            "cell_value": None
        })
    
    def start_cell_selection(self, selector_index):
        if not self.workbook:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first")
            return
        
        self.current_selector_index = selector_index - 1
        self.lbl_instructions.show()
        
        # Bring Excel to foreground
        try:
            self.excel_app.Visible = True
            # self.excel_app.WindowState = -4137  # xlNormal
            self.excel_app.Activate()
        except:
            pass
        
        # Start checking for selection
        self.selection_timer.start(500)  # Check every 500ms
    
    def check_excel_selection(self):
        if self.current_selector_index is None:
            return
            
        try:
            # Get the current selection in Excel
            selection = self.excel_app.Selection
            if selection:
                cell_ref = selection.Address.replace('$', '')
                cell_value = str(selection.Value) if selection.Value is not None else "Empty"
                
                # Update the GUI
                self.cell_selections[self.current_selector_index]["cell_ref"] = cell_ref
                self.cell_selections[self.current_selector_index]["cell_value"] = cell_value
                self.cell_selections[self.current_selector_index]["cell_ref_label"].setText(cell_ref)
                self.cell_selections[self.current_selector_index]["cell_value_label"].setText(cell_value)
                
                # Stop the timer
                self.selection_timer.stop()
                self.current_selector_index = None
                self.lbl_instructions.hide()
                
                # Bring our app back to foreground
                self.activateWindow()
                self.raise_()
        except Exception as e:
            print(f"Error checking selection: {e}")
    
    def remove_selector(self, selector_index):
        # Find the widget to remove
        for i, selector in enumerate(self.cell_selections):
            if i == selector_index - 1:
                selector["widget"].setParent(None)
                self.cell_selections.remove(selector)
                break
        
        # Renumber remaining selectors
        for i, selector in enumerate(self.cell_selections):
            selector["select_button"].setText(f"Select Cell {i+1}")
    
    def save_changes(self):
        if not self.workbook or not self.file_path:
            QMessageBox.warning(self, "Warning", "No Excel file loaded")
            return
            
        # Get all valid cell references
        cells_to_mark = [
            selector["cell_ref"] for selector in self.cell_selections 
            if selector["cell_ref"]
        ]
        
        if not cells_to_mark:
            QMessageBox.warning(self, "Warning", "No cells selected for marking")
            return
            
        try:
            # Apply borders to selected cells
            for cell_ref in cells_to_mark:
                cell = self.excel_app.Range(cell_ref)
                cell.Borders.LineStyle = 1  # xlContinuous
                cell.Borders.Weight = 2      # xlThin
            
            # Save the workbook
            self.workbook.Save()
            QMessageBox.information(
                self, "Success", f"Borders added to cells and file saved: {self.file_path}"
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Error", f"Failed to save changes: {str(e)}"
            )
    
    def closeEvent(self, event):
        # Clean up Excel when closing
        if self.excel_app:
            try:
                self.workbook.Close(False)
                self.excel_app.Quit()
            except:
                pass
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelCellMarkerApp()
    window.show()
    sys.exit(app.exec())