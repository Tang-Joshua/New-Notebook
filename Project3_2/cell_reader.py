import json

TRACKED_CELLS_FILE = "tracked_cells.json"

class TrackedCellReader:
    def __init__(self):
        self.tracked_cells = self._load_tracked_cells()
    
    def _load_tracked_cells(self):
        """Load tracked cells from file"""
        try:
            with open(TRACKED_CELLS_FILE, 'r') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}
    
    def get_cell_value(self, cell_id):
        """Get value for a specific cell ID"""
        return self.tracked_cells.get(cell_id, {}).get('value')
    
    def get_all_cells(self):
        """Get all tracked cells"""
        return self.tracked_cells
    
    def get_recent_cell(self):
        """Get the most recently tracked cell"""
        if not self.tracked_cells:
            return None
        
        # Find cell with latest timestamp
        latest = max(self.tracked_cells.values(), key=lambda x: x['timestamp'])
        return latest

# Example usage
if __name__ == "__main__":
    reader = TrackedCellReader()
    
    # Get all tracked cells
    all_cells = reader.get_all_cells()
    print("All tracked cells:", all_cells)
    
    # Get most recent cell
    recent = reader.get_recent_cell()
    if recent:
        print(f"Most recent cell: {recent['sheet']}!{recent['address']} = {recent['value']}")