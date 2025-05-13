from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QPushButton, QButtonGroup, QFrame
)
from PyQt6.QtCore import Qt, pyqtSignal


class ExcelToolbarKit(QWidget):
    # Signals to communicate with the main spreadsheet
    alignmentChanged = pyqtSignal(Qt.AlignmentFlag, name='alignmentChanged')
    wrapTextChanged = pyqtSignal(bool, name='wrapTextChanged')
    mergeChanged = pyqtSignal(bool, name='mergeChanged')

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_cell = None  # Track the current cell
        self.setup_ui()
        self.connect_signals()

    def setup_ui(self):
        self.setStyleSheet("""
            QWidget {
                background: #f3f3f3;
            }
            QPushButton {
                padding: 4px 8px;
                border: 1px solid #c0c0c0;
                border-radius: 3px;
                background: white;
                min-width: 24px;
                color: #333;
            }
            QPushButton:hover {
                background: #e6e6e6;
            }
            QPushButton:pressed {
                background: #d4d4d4;
            }
            QPushButton:checked {
                background: #d0ebff;
                border-color: #7fbbda;
            }
            QPushButton[flat="true"] {
                border: none;
                background: transparent;
            }
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(4)

        def add_separator():
            sep = QFrame()
            sep.setFrameShape(QFrame.Shape.VLine)
            sep.setFrameShadow(QFrame.Shadow.Sunken)
            sep.setStyleSheet("color: #c0c0c0;")
            layout.addWidget(sep)

        # Horizontal alignment group
        self.h_align_group = QButtonGroup(self)
        self.h_left_btn = QPushButton("⯇")
        self.h_left_btn.setToolTip("Align Left")
        self.h_center_btn = QPushButton("⇔")
        self.h_center_btn.setToolTip("Align Center")
        self.h_right_btn = QPushButton("⯈")
        self.h_right_btn.setToolTip("Align Right")
        
        for btn in [self.h_left_btn, self.h_center_btn, self.h_right_btn]:
            btn.setCheckable(True)
            btn.setFixedSize(28, 24)
            self.h_align_group.addButton(btn)
            layout.addWidget(btn)
        self.h_align_group.setExclusive(True)

        add_separator()

        # Vertical alignment group
        self.v_align_group = QButtonGroup(self)
        self.v_top_btn = QPushButton("⯅")
        self.v_top_btn.setToolTip("Align Top")
        self.v_middle_btn = QPushButton("=")
        self.v_middle_btn.setToolTip("Align Middle")
        self.v_bottom_btn = QPushButton("⯆")
        self.v_bottom_btn.setToolTip("Align Bottom")
        
        for btn in [self.v_top_btn, self.v_middle_btn, self.v_bottom_btn]:
            btn.setCheckable(True)
            btn.setFixedSize(28, 24)
            self.v_align_group.addButton(btn)
            layout.addWidget(btn)
        self.v_align_group.setExclusive(True)

        add_separator()

        # Wrap Text button
        self.wrap_text_btn = QPushButton("Wrap Text")
        self.wrap_text_btn.setCheckable(True)
        self.wrap_text_btn.setToolTip("Wrap Text")
        self.wrap_text_btn.setFixedHeight(24)
        layout.addWidget(self.wrap_text_btn)

        add_separator()

        # Merge & Center button
        self.merge_center_btn = QPushButton("Merge & Center")
        self.merge_center_btn.setCheckable(True)
        self.merge_center_btn.setToolTip("Merge & Center")
        self.merge_center_btn.setFixedHeight(24)
        layout.addWidget(self.merge_center_btn)

        layout.addStretch()

    def connect_signals(self):
        """Connect button signals to actions"""
        self.h_left_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignLeft))
        self.h_center_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignHCenter))
        self.h_right_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignRight))
        
        self.v_top_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignTop))
        self.v_middle_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignVCenter))
        self.v_bottom_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignBottom))
        
        self.wrap_text_btn.clicked.connect(self.wrapTextChanged.emit)
        self.merge_center_btn.clicked.connect(self.mergeChanged.emit)

    def emit_alignment(self, alignment):
        """Emit alignment signal and update button states"""
        self.alignmentChanged.emit(alignment)
        self.update_button_states()

    def update_for_cell(self, cell):
        """Update toolbar state based on cell properties"""
        self.current_cell = cell
        if cell is None:
            return
            
        # Update horizontal alignment buttons
        alignment = cell.alignment()
        self.h_left_btn.setChecked(alignment & Qt.AlignmentFlag.AlignLeft)
        self.h_center_btn.setChecked(alignment & Qt.AlignmentFlag.AlignHCenter)
        self.h_right_btn.setChecked(alignment & Qt.AlignmentFlag.AlignRight)
        
        # Update vertical alignment buttons
        self.v_top_btn.setChecked(alignment & Qt.AlignmentFlag.AlignTop)
        self.v_middle_btn.setChecked(alignment & Qt.AlignmentFlag.AlignVCenter)
        self.v_bottom_btn.setChecked(alignment & Qt.AlignmentFlag.AlignBottom)
        
        # Update other formatting buttons
        self.wrap_text_btn.setChecked(cell.wrapText())
        self.merge_center_btn.setChecked(cell.columnSpan() > 1 or cell.rowSpan() > 1)

    def update_button_states(self):
        """Update button states based on current cell"""
        if self.current_cell:
            self.update_for_cell(self.current_cell)