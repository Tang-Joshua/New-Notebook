import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, 
    QFileDialog, QMessageBox, QHBoxLayout, QScrollArea, QLabel,
    QInputDialog, QComboBox
)
from PyQt6.QtCore import Qt, QTimer
import xlwings as xw
import pythoncom


class ExcelFormulaApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Formula Generator")
        self.setGeometry(100, 100, 900, 600)
        
        # Excel application and workbook references
        self.excel_app = None
        self.workbook = None
        self.file_path = ""
        self.cell_selections = []
        
        # Setup UI
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        
        # File selection
        self.btn_select_file = QPushButton("Select Excel File")
        self.btn_select_file.clicked.connect(self.select_excel_file)
        self.layout.addWidget(self.btn_select_file)
        
        self.lbl_current_file = QLabel("No file selected")
        self.layout.addWidget(self.lbl_current_file)
        
        # Formula type selection
        self.formula_combo = QComboBox()
        self.formula_combo.addItems(["SUM", "AVERAGE", "COUNT", "MAX", "MIN"])
        self.formula_combo.setEnabled(False)
        self.layout.addWidget(QLabel("Select Formula Type:"))
        self.layout.addWidget(self.formula_combo)
        
        # Target cell selection
        self.btn_select_target = QPushButton("Select Target Cell for Formula")
        self.btn_select_target.clicked.connect(self.select_target_cell)
        self.btn_select_target.setEnabled(False)
        self.layout.addWidget(self.btn_select_target)
        
        self.lbl_target_cell = QLabel("No target cell selected")
        self.layout.addWidget(self.lbl_target_cell)
        
        # Instructions
        self.lbl_instructions = QLabel(
            "After selecting file, click 'Add Cell to Formula' buttons to select cells for your formula"
        )
        self.layout.addWidget(self.lbl_instructions)
        
        # Scroll area for cell selectors
        self.scroll_area = QScrollArea()
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        self.scroll_area.setWidgetResizable(True)
        self.layout.addWidget(self.scroll_area)
        
        # Control buttons
        self.btn_add_selector = QPushButton("Add Cell to Formula")
        self.btn_add_selector.clicked.connect(self.add_cell_selector)
        self.btn_add_selector.setEnabled(False)
        self.layout.addWidget(self.btn_add_selector)
        
        self.btn_generate = QPushButton("Generate Formula in Excel")
        self.btn_generate.clicked.connect(self.generate_formula)
        self.btn_generate.setEnabled(False)
        self.layout.addWidget(self.btn_generate)
        
        # Status label
        self.lbl_status = QLabel("Ready")
        self.layout.addWidget(self.lbl_status)
        
        # Add first selector
        self.add_cell_selector()
        
        # Variables for selection
        self.selection_timer = QTimer()
        self.selection_timer.timeout.connect(self.check_excel_selection)
        self.currently_selecting_for = None
        self.in_selection_mode = False
        self.target_cell = None
    
    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls *.xlsm)"
        )
        
        if file_path:
            try:
                # Close any existing connection
                if self.workbook:
                    try:
                        self.workbook.close()
                    except:
                        pass
                
                # Initialize Excel connection
                pythoncom.CoInitialize()
                self.excel_app = xw.App(visible=True)
                self.workbook = self.excel_app.books.open(file_path)
                self.file_path = file_path
                
                self.lbl_current_file.setText(f"Current file: {os.path.basename(file_path)}")
                self.btn_add_selector.setEnabled(True)
                self.btn_generate.setEnabled(True)
                self.btn_select_target.setEnabled(True)
                self.formula_combo.setEnabled(True)
                self.lbl_status.setText("File loaded. Select cells and generate formulas.")
                
                # Bring Excel to front
                self.excel_app.visible = True
                self.excel_app.activate()
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load Excel file: {str(e)}")
                self.lbl_status.setText(f"Error: {str(e)}")
    
    def select_target_cell(self):
        if not self.workbook:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first")
            return
            
        self.lbl_status.setText("Select target cell in Excel for the formula...")
        self.excel_app.activate()
        
        # Use timer to check for selection
        self.currently_selecting_for = "target"
        self.in_selection_mode = True
        self.selection_timer.start(500)
    
    def add_cell_selector(self):
        selector_index = len(self.cell_selections) + 1
        selector_widget = QWidget()
        selector_layout = QHBoxLayout(selector_widget)
        
        # Select cell button
        btn_select = QPushButton(f"Select Cell {selector_index}")
        btn_select.clicked.connect(lambda: self.start_cell_selection(selector_index))
        
        # Cell reference display
        lbl_cell_ref = QLabel("Not selected")
        lbl_cell_ref.setStyleSheet("font-weight: bold; min-width: 100px;")
        
        # Cell value display
        lbl_cell_value = QLabel("No value")
        lbl_cell_value.setStyleSheet("min-width: 200px;")
        
        # Cell address display
        lbl_cell_address = QLabel("")
        lbl_cell_address.setStyleSheet("font-style: italic; min-width: 150px;")
        
        # Remove button
        btn_remove = QPushButton("Remove")
        btn_remove.clicked.connect(lambda: self.remove_selector(selector_index))
        
        selector_layout.addWidget(btn_select)
        selector_layout.addWidget(QLabel("Reference:"))
        selector_layout.addWidget(lbl_cell_ref)
        selector_layout.addWidget(QLabel("Value:"))
        selector_layout.addWidget(lbl_cell_value)
        selector_layout.addWidget(QLabel("Address:"))
        selector_layout.addWidget(lbl_cell_address)
        selector_layout.addWidget(btn_remove)
        
        self.scroll_layout.addWidget(selector_widget)
        
        self.cell_selections.append({
            "widget": selector_widget,
            "select_button": btn_select,
            "cell_ref_label": lbl_cell_ref,
            "cell_value_label": lbl_cell_value,
            "cell_address_label": lbl_cell_address,
            "cell_ref": None,
            "cell_value": None,
            "cell_address": None
        })
    
    def start_cell_selection(self, selector_index):
        if not self.workbook:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first")
            return
            
        self.currently_selecting_for = selector_index - 1
        self.in_selection_mode = True
        
        # Update status and bring Excel to front
        self.lbl_status.setText(f"Select a cell in Excel for Selector {selector_index}...")
        self.excel_app.activate()
        
        # Start checking for selection
        self.selection_timer.start(500)
    
    def check_excel_selection(self):
        if not self.in_selection_mode:
            self.selection_timer.stop()
            return
            
        try:
            # Get the current selection in Excel
            selected_range = self.excel_app.selection
            
            # If it's a single cell
            if selected_range.count == 1:
                cell = selected_range
                cell_ref = cell.address.replace('$', '')
                cell_value = str(cell.value) if cell.value is not None else "Empty"
                sheet_name = cell.sheet.name
                full_address = f"{sheet_name}!{cell_ref}"
                
                if self.currently_selecting_for == "target":
                    # Handle target cell selection
                    self.target_cell = full_address
                    self.lbl_target_cell.setText(full_address)
                    self.lbl_status.setText("Target cell selected. Now select cells for the formula.")
                else:
                    # Handle formula cell selection
                    selector = self.cell_selections[self.currently_selecting_for]
                    selector["cell_ref"] = cell_ref
                    selector["cell_value"] = cell_value
                    selector["cell_address"] = full_address
                    
                    selector["cell_ref_label"].setText(cell_ref)
                    selector["cell_value_label"].setText(cell_value)
                    selector["cell_address_label"].setText(full_address)
                
                # Stop selection mode
                self.in_selection_mode = False
                self.currently_selecting_for = None
                self.selection_timer.stop()
                
                # Bring our app back to front
                self.activateWindow()
                self.raise_()
                
        except Exception as e:
            self.lbl_status.setText(f"Error checking selection: {str(e)}")
            self.selection_timer.stop()
    
    def remove_selector(self, selector_index):
        for i, selector in enumerate(self.cell_selections):
            if i == selector_index - 1:
                selector["widget"].setParent(None)
                self.cell_selections.remove(selector)
                break
        
        # Renumber remaining selectors
        for i, selector in enumerate(self.cell_selections):
            selector["select_button"].setText(f"Select Cell {i+1}")
    
    def generate_formula(self):
        if not self.workbook or not self.file_path:
            QMessageBox.warning(self, "Warning", "No Excel file loaded")
            return
            
        if not self.target_cell:
            QMessageBox.warning(self, "Warning", "Please select a target cell first")
            return
            
        selected_cells = [
            selector["cell_address"] for selector in self.cell_selections 
            if selector["cell_address"]
        ]
        
        if not selected_cells:
            QMessageBox.warning(self, "Warning", "No cells selected for formula")
            return
            
        try:
            # Get formula type
            formula_type = self.formula_combo.currentText()
            
            # Create the formula string
            cell_refs = ",".join([addr.split('!')[1] for addr in selected_cells])
            formula = f"={formula_type}({cell_refs})"
            
            # Parse target cell
            if '!' in self.target_cell:
                sheet_name, cell_ref = self.target_cell.split('!')
                sheet = self.workbook.sheets[sheet_name]
            else:
                sheet = self.workbook.sheets.active
                cell_ref = self.target_cell
            
            # Apply the formula
            sheet.range(cell_ref).formula = formula
            
            # Save the workbook
            self.workbook.save()
            
            # Show the generated formula
            self.lbl_status.setText(f"Formula generated: {formula}")
            QMessageBox.information(
                self, 
                "Success", 
                f"Formula added to {self.target_cell}:\n{formula}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate formula: {str(e)}")
    
    def closeEvent(self, event):
        # Clean up Excel connection when closing
        if self.workbook:
            try:
                self.workbook.close() 
            except:
                pass
        if self.excel_app:
            try:
                self.excel_app.quit()
            except:
                pass
        pythoncom.CoUninitialize()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelFormulaApp()
    window.show()
    sys.exit(app.exec())