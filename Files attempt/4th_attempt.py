import sys
import pythoncom
import win32com.client as win32
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget,
                            QLabel, QLineEdit, QPushButton, QFileDialog, QTextEdit,
                            QHBoxLayout, QGroupBox)
from PyQt6.QtCore import Qt, QObject, pyqtSignal, QThread, QTimer
import time

class ExcelEventWorker(QObject):
    cell_selected = pyqtSignal(str, str, object, int)  # sheet_name, cell_ref, value, selection_num
    status_update = pyqtSignal(str)
    finished = pyqtSignal()
    highlight_request = pyqtSignal(str, str, int)  # sheet_name, cell_ref, selection_num

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self._running = True
        self.excel = None
        self.workbook = None
        self.last_range1 = None
        self.last_range2 = None
        self.current_selection_num = 1  # Tracks which selection we're capturing (1 or 2)
        self.highlight_colors = {1: (255, 0, 0), 2: (0, 0, 255)}  # Red for 1st, Blue for 2nd

    def run(self):
        pythoncom.CoInitialize()
        try:
            self.status_update.emit("Starting Excel...")
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')
            self.excel.Visible = True
            self.status_update.emit(f"Opening workbook: {self.excel_path}")
            self.workbook = self.excel.Workbooks.Open(self.excel_path)
            self.status_update.emit("Ready - select cells in Excel")

            # Clear any existing highlights
            self.clear_highlights()

            while self._running:
                try:
                    selection = self.excel.Selection
                    current_range = str(selection.Address)
                    sheet_name = str(selection.Worksheet.Name)
                    
                    # Check if this is a new selection for the current selection number
                    if ((self.current_selection_num == 1 and current_range != self.last_range1) or
                        (self.current_selection_num == 2 and current_range != self.last_range2)):
                        
                        # Update the appropriate last range
                        if self.current_selection_num == 1:
                            self.last_range1 = current_range
                        else:
                            self.last_range2 = current_range
                        
                        range_ref = current_range.replace('$', '')
                        
                        # Get the cell value
                        try:
                            if selection.MergeCells:
                                merged_area = selection.MergeArea
                                cell_value = merged_area.Cells(1, 1).Value2
                                range_ref = merged_area.Address.replace('$', '')
                            else:
                                cell_value = selection.Value2
                            
                            if isinstance(cell_value, (list, tuple)):
                                cell_value = cell_value[0][0] if cell_value[0][0] is not None else ""
                            elif cell_value is None:
                                cell_value = ""
                            
                            # Emit the selection
                            self.cell_selected.emit(
                                sheet_name,
                                range_ref,
                                str(cell_value) if cell_value is not None else "",
                                self.current_selection_num
                            )
                            
                            # Request highlight
                            self.highlight_request.emit(sheet_name, range_ref, self.current_selection_num)
                            
                        except Exception as e:
                            self.status_update.emit(f"Value read error: {str(e)}")

                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.2)
                    
                except pythoncom.com_error as e:
                    if e.excepinfo[5] == -2146827284:  # Common Excel COM error
                        time.sleep(0.5)
                        continue
                    self.status_update.emit(f"Selection error: {str(e)}")
                    time.sleep(1)
                except Exception as e:
                    self.status_update.emit(f"Unexpected error: {str(e)}")
                    time.sleep(1)

        except Exception as e:
            self.status_update.emit(f"Initialization error: {str(e)}")
        finally:
            self.cleanup()
            self.finished.emit()

    def clear_highlights(self):
        """Remove all existing highlights from the worksheet"""
        if not self.workbook:
            return
            
        try:
            for sheet in self.workbook.Sheets:
                used_range = sheet.UsedRange
                used_range.Borders.LineStyle = win32.constants.xlLineStyleNone
                used_range.Interior.ColorIndex = win32.constants.xlColorIndexNone
        except Exception as e:
            self.status_update.emit(f"Error clearing highlights: {str(e)}")

    def highlight_cell(self, sheet_name, cell_ref, selection_num):
        """Highlight the specified cell with the appropriate color"""
        try:
            sheet = self.workbook.Sheets(sheet_name)
            cell = sheet.Range(cell_ref)
            
            # Clear previous highlight for this selection number
            if selection_num == 1 and self.last_range1:
                prev_cell = sheet.Range(self.last_range1)
                prev_cell.Borders.LineStyle = win32.constants.xlLineStyleNone
            elif selection_num == 2 and self.last_range2:
                prev_cell = sheet.Range(self.last_range2)
                prev_cell.Borders.LineStyle = win32.constants.xlLineStyleNone
            
            # Apply new highlight
            color = self.highlight_colors[selection_num]
            border = cell.Borders
            border.LineStyle = win32.constants.xlContinuous
            border.Weight = win32.constants.xlThick
            border.Color = RGB(*color)
            
        except Exception as e:
            self.status_update.emit(f"Highlight error: {str(e)}")

    def cleanup(self):
        try:
            if hasattr(self, 'workbook') and self.workbook:
                self.workbook.Close(False)
            if hasattr(self, 'excel') and self.excel:
                self.excel.Quit()
        except:
            pass
        finally:
            pythoncom.CoUninitialize()

    def stop(self):
        self._running = False

def RGB(r, g, b):
    """Helper function to convert RGB to Excel color format"""
    return r + (g << 8) + (b << 16)

class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Dual Cell Selector")
        self.setGeometry(100, 100, 600, 500)
        self.worker = None
        self.thread = None
        self.current_selection = 1  # 1 or 2
        
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        # File selection
        self.file_path = QLineEdit()
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.clicked.connect(self.browse_file)

        # Control buttons
        self.start_btn = QPushButton("Start Listening")
        self.start_btn.clicked.connect(self.start_listening)
        self.start_btn.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; padding: 8px; }"
            "QPushButton:disabled { background-color: #cccccc; }"
        )

        self.stop_btn = QPushButton("Stop Listening")
        self.stop_btn.clicked.connect(self.stop_listening)
        self.stop_btn.setStyleSheet(
            "QPushButton { background-color: #f44336; color: white; padding: 8px; }"
            "QPushButton:disabled { background-color: #cccccc; }"
        )
        self.stop_btn.setEnabled(False)

        # Selection control buttons
        self.selection1_btn = QPushButton("Select First Cell")
        self.selection1_btn.setStyleSheet("background-color: #ffcccc;")
        self.selection1_btn.clicked.connect(lambda: self.set_selection_mode(1))
        
        self.selection2_btn = QPushButton("Select Second Cell")
        self.selection2_btn.setStyleSheet("background-color: #ccccff;")
        self.selection2_btn.clicked.connect(lambda: self.set_selection_mode(2))
        
        selection_btn_layout = QHBoxLayout()
        selection_btn_layout.addWidget(self.selection1_btn)
        selection_btn_layout.addWidget(self.selection2_btn)

        # Status display
        self.status_label = QLabel("Ready to connect to Excel")
        self.status_label.setStyleSheet("font-size: 12px; color: #666;")

        # Selection displays
        self.create_selection_group("First Selection", 1)
        self.create_selection_group("Second Selection", 2)

        # Layout
        layout.addWidget(QLabel("Excel File:"))
        layout.addWidget(self.file_path)
        layout.addWidget(self.browse_btn)
        layout.addWidget(self.start_btn)
        layout.addWidget(self.stop_btn)
        layout.addLayout(selection_btn_layout)
        layout.addWidget(self.status_label)
        layout.addWidget(self.selection_group1)
        layout.addWidget(self.selection_group2)

        central_widget.setLayout(layout)

    def create_selection_group(self, title, selection_num):
        """Create a group box for a selection display"""
        group = QGroupBox(title)
        layout = QVBoxLayout()
        
        sheet_label = QLabel("Sheet: Not selected")
        cell_label = QLabel("Cell: Not selected")
        value_display = QTextEdit()
        value_display.setReadOnly(True)
        value_display.setStyleSheet("""
            font-family: Consolas, monospace;
            font-size: 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            background-color: #f9f9f9;
            min-height: 80px;
        """)
        
        if selection_num == 1:
            self.sheet_label1 = sheet_label
            self.cell_label1 = cell_label
            self.value_display1 = value_display
            group.setStyleSheet("QGroupBox { border: 2px solid #ff0000; }")
        else:
            self.sheet_label2 = sheet_label
            self.cell_label2 = cell_label
            self.value_display2 = value_display
            group.setStyleSheet("QGroupBox { border: 2px solid #0000ff; }")
        
        layout.addWidget(sheet_label)
        layout.addWidget(cell_label)
        layout.addWidget(QLabel("Value:"))
        layout.addWidget(value_display)
        group.setLayout(layout)
        
        if selection_num == 1:
            self.selection_group1 = group
        else:
            self.selection_group2 = group

    def browse_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if file:
            self.file_path.setText(file)
            self.status_label.setText("File selected - ready to connect")

    def set_selection_mode(self, selection_num):
        """Set which selection we're currently capturing (1 or 2)"""
        self.current_selection = selection_num
        if selection_num == 1:
            self.status_label.setText("Select FIRST cell in Excel (red border)")
            self.selection1_btn.setStyleSheet("background-color: #ff9999; font-weight: bold;")
            self.selection2_btn.setStyleSheet("background-color: #ccccff; font-weight: normal;")
        else:
            self.status_label.setText("Select SECOND cell in Excel (blue border)")
            self.selection1_btn.setStyleSheet("background-color: #ffcccc; font-weight: normal;")
            self.selection2_btn.setStyleSheet("background-color: #9999ff; font-weight: bold;")

    def start_listening(self):
        if not self.file_path.text():
            self.status_label.setText("Error: Please select an Excel file first")
            return

        self.stop_listening()  # Ensure clean start

        # Setup worker thread
        self.worker = ExcelEventWorker(self.file_path.text())
        self.thread = QThread()
        self.worker.moveToThread(self.thread)

        # Connect signals
        self.worker.cell_selected.connect(self.update_cell_info)
        self.worker.status_update.connect(self.status_label.setText)
        self.worker.finished.connect(self.thread.quit)
        self.worker.highlight_request.connect(self.request_highlight)
        
        # Pass current selection mode to worker
        self.set_selection_mode(1)
        QTimer.singleShot(100, lambda: setattr(self.worker, 'current_selection_num', self.current_selection))

        self.thread.started.connect(self.worker.run)
        self.thread.finished.connect(lambda: self.stop_btn.setEnabled(False))
        self.thread.finished.connect(lambda: self.start_btn.setEnabled(True))

        # Start the thread
        self.thread.start()

        # Update UI
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

    def stop_listening(self):
        if self.worker:
            self.worker.stop()
        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait(500)

    def request_highlight(self, sheet_name, cell_ref, selection_num):
        """Forward highlight request to worker thread"""
        if self.worker:
            self.worker.current_selection_num = selection_num
            QTimer.singleShot(100, lambda: self.worker.highlight_cell(sheet_name, cell_ref, selection_num))

    def update_cell_info(self, sheet_name, cell_ref, value, selection_num):
        """Update the appropriate display based on which selection this is"""
        if selection_num == 1:
            self.sheet_label1.setText(f"Sheet: {sheet_name}")
            self.cell_label1.setText(f"Cell: {cell_ref}")
            self.value_display1.setPlainText(str(value))
        else:
            self.sheet_label2.setText(f"Sheet: {sheet_name}")
            self.cell_label2.setText(f"Cell: {cell_ref}")
            self.value_display2.setPlainText(str(value))
        
        # Automatically switch to next selection if applicable
        if selection_num == 1 and self.current_selection == 1:
            self.set_selection_mode(2)
            if self.worker:
                self.worker.current_selection_num = 2

    def closeEvent(self, event):
        self.stop_listening()
        super().closeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ExcelViewer()
    window.show()
    sys.exit(app.exec())