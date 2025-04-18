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

            # Get the active sheet
            sheet = self.excel.ActiveSheet

            # Main monitoring loop
            while self._running:
                try:
                    current_cell = str(self.excel.Selection.Address)
                    if current_cell != self.last_cell:
                        self.last_cell = current_cell
                        cell_ref = current_cell.replace('$', '')
                        selection = self.excel.Selection
                        
                        # Get the complete value (handles different data types)
                        if selection.Value is not None:
                            # For single cells
                            if selection.Count == 1:
                                cell_value = selection.Value
                                # Handle different value formats
                                if isinstance(cell_value, (list, tuple)):
                                    cell_value = cell_value[0][0] if cell_value[0][0] is not None else ""
                                elif cell_value is None:
                                    cell_value = ""
                            else:
                                # For multiple cells (range selection)
                                cell_value = "Multiple cells selected"
                        else:
                            cell_value = ""
                        
                        self.cell_selected.emit(
                            sheet.Name,
                            cell_ref,
                            str(cell_value)  # Convert to string to handle all types
                        )
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.1)
                except Exception as e:
                    self.status_update.emit(f"Monitoring error: {str(e)}")
                    time.sleep(1)

        except Exception as e:
            self.status_update.emit(f"Error: {str(e)}")
        finally:
            self.cleanup()
            self.finished.emit()

    def cleanup(self):
        try:
            if hasattr(self, 'workbook') and self.workbook:
                self.workbook.Close(False)
            if hasattr(self, 'excel') and self.excel:
                self.excel.Quit()
        except Exception as e:
            self.status_update.emit(f"Cleanup error: {str(e)}")
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
        self.value_label.setText(f"Value: {value if value is not None else 'Empty'}")

    def closeEvent(self, event):
        self.stop_listening()
        super().closeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ExcelViewer()
    window.show()
    sys.exit(app.exec())