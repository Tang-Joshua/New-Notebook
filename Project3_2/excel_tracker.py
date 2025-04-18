import xlwings as xw
import json
import os
from datetime import datetime

# Configuration
TRACKED_CELLS_FILE = "tracked_cells.json"

class ExcelCellTracker:
    def __init__(self):
        self.tracked_cells = self._load_tracked_cells()
        self.app = xw.App(visible=True)
    
    def _load_tracked_cells(self):
        """Load previously tracked cells from file"""
        if os.path.exists(TRACKED_CELLS_FILE):
            with open(TRACKED_CELLS_FILE, 'r') as f:
                return json.load(f)
        return {}
    
    def _save_tracked_cells(self):
        """Save current tracked cells to file"""
        with open(TRACKED_CELLS_FILE, 'w') as f:
            json.dump(self.tracked_cells, f, indent=2)
    
    def track_selected_cell(self):
        """Track the currently selected cell in active Excel workbook"""
        try:
            # Get the active Excel application
            excel = self.app
            
            # Get the active sheet and selection
            sheet = excel.books.active.sheets.active
            selection = sheet.range(xw.selection.address)
            
            # Create unique identifier for this cell
            cell_id = f"{sheet.name}!{selection.address}"
            
            # Store cell information
            self.tracked_cells[cell_id] = {
                "value": selection.value,
                "address": selection.address,
                "sheet": sheet.name,
                "workbook": excel.books.active.name,
                "timestamp": datetime.now().isoformat()
            }
            
            # Save to file
            self._save_tracked_cells()
            
            print(f"Tracked cell: {cell_id} with value: {selection.value}")
            return True
        except Exception as e:
            print(f"Error tracking cell: {e}")
            return False
    
    def run(self):
        """Run the tracker (you would call this from Excel)"""
        print("Excel Cell Tracker is running...")
        return self.track_selected_cell()

# For direct execution
if __name__ == "__main__":
    tracker = ExcelCellTracker()
    tracker.run()