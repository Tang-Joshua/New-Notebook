import sys
import pythoncom
import win32com.client as win32
from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget,
                            QLabel, QLineEdit, QPushButton, QFileDialog)
from PyQt6.QtCore import Qt, QObject, pyqtSignal, QThread
import time

class ExcelEventWorker(QObject):
    cell_selected = pyqtSignal(str, str, object)  # sheet_name, cell_ref, value
    status_update = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self._running = True
        self.excel = None
        self.workbook = None
        self.last_cell = None

    def run(self):
        pythoncom.CoInitialize()
        try:
            self.status_update.emit("Starting Excel...")
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')
            self.excel.Visible = True
            self.status_update.emit(f"Opening workbook: {self.excel_path}")
            self.workbook = self.excel.Workbooks.Open(self.excel_path)
            self.status_update.emit("Ready - select cells in Excel")

            while self._running:
                try:
                    selection = self.excel.Selection
                    current_cell = str(selection.Address)
                    
                    if current_cell != self.last_cell:
                        self.last_cell = current_cell
                        sheet_name = str(selection.Worksheet.Name)
                        cell_ref = current_cell.replace('$', '')
                        
                        try:
                            if selection.MergeCells:
                                merged_area = selection.MergeArea
                                cell_value = merged_area.Cells(1, 1).Value
                                cell_ref = merged_area.Address.replace('$', '')
                            else:
                                if selection.Count == 1:
                                    cell_value = selection.Value
                                    if isinstance(cell_value, (list, tuple)):
                                        cell_value = cell_value[0][0] if cell_value[0][0] is not None else ""
                                else:
                                    cell_value = "Multiple cells selected"
                        except:
                            cell_value = "N/A"

                        self.cell_selected.emit(
                            sheet_name,
                            cell_ref,
                            str(cell_value) if cell_value is not None else ""
                        )
                    
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.2)
                    
                except pythoncom.com_error as e:
                    if e.excepinfo[5] == -2146827284:
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

class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Cell Listener")
        self.setGeometry(100, 100, 500, 350)
        self.worker = None
        self.thread = None
        
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

        # Status display
        self.status_label = QLabel("Ready to connect to Excel")
        self.status_label.setStyleSheet("font-size: 12px; color: #666;")

        # Cell info display
        self.sheet_label = QLabel("Sheet: Not connected")
        self.cell_label = QLabel("Cell: None selected")
        self.value_label = QLabel("Value: -")
        
        # Formatting
        for label in [self.sheet_label, self.cell_label, self.value_label]:
            label.setStyleSheet("""
                font-size: 14px; 
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f9f9f9;
            """)

        # Layout
        layout.addWidget(QLabel("Excel File:"))
        layout.addWidget(self.file_path)
        layout.addWidget(self.browse_btn)
        layout.addWidget(self.start_btn)
        layout.addWidget(self.stop_btn)
        layout.addWidget(self.status_label)
        layout.addWidget(QLabel("Current Selection:"))
        layout.addWidget(self.sheet_label)
        layout.addWidget(self.cell_label)
        layout.addWidget(self.value_label)

        central_widget.setLayout(layout)

    def browse_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if file:
            self.file_path.setText(file)
            self.status_label.setText("File selected - ready to connect")

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
        
        self.thread.started.connect(self.worker.run)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
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

    def update_cell_info(self, sheet_name, cell_ref, value):
        self.sheet_label.setText(f"Sheet: {sheet_name}")
        self.cell_label.setText(f"Cell: {cell_ref}")
        self.value_label.setText(f"Value: {value if value is not None else ''}")

    def closeEvent(self, event):
        self.stop_listening()
        super().closeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ExcelViewer()
    window.show()
    sys.exit(app.exec())