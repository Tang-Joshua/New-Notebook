from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QPushButton, QButtonGroup, QFrame
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QIcon


class ExcelToolbarKit(QWidget):
    alignmentChanged = pyqtSignal(Qt.AlignmentFlag, name='alignmentChanged')
    wrapTextChanged = pyqtSignal(bool, name='wrapTextChanged')
    mergeChanged = pyqtSignal(bool, name='mergeChanged')

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_cell = None
        self.setup_ui()
        self.connect_signals()

    def setup_ui(self):
        self.setStyleSheet("""
            QPushButton {
                padding: 4px;
                border: none;
                border-radius: 4px;
                background-color: transparent;
                min-width: 28px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:checked {
                background-color: #cce4ff;
                border: 1px solid #66aaff;
            }
            QFrame {
                color: #d0d0d0;
            }
        """)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(6)

        def add_separator():
            sep = QFrame()
            sep.setFrameShape(QFrame.Shape.VLine)
            sep.setFrameShadow(QFrame.Shadow.Sunken)
            sep.setFixedWidth(2)
            layout.addWidget(sep)

        def icon(path):
            return QIcon(path)

        # Horizontal alignment buttons
        self.h_align_group = QButtonGroup(self)
        self.h_left_btn = QPushButton()
        self.h_left_btn.setIcon(icon("Main_File/icons/align_left.png"))
        self.h_left_btn.setToolTip("Align Left")

        self.h_center_btn = QPushButton()
        self.h_center_btn.setIcon(icon("Main_File/icons/align_center.png"))
        self.h_center_btn.setToolTip("Align Center")

        self.h_right_btn = QPushButton()
        self.h_right_btn.setIcon(icon("Main_File/icons/align_right.png"))
        self.h_right_btn.setToolTip("Align Right")

        for btn in [self.h_left_btn, self.h_center_btn, self.h_right_btn]:
            btn.setCheckable(True)
            btn.setFixedSize(32, 28)
            self.h_align_group.addButton(btn)
            layout.addWidget(btn)
        self.h_align_group.setExclusive(True)

        add_separator()

        # Vertical alignment buttons
        self.v_align_group = QButtonGroup(self)
        self.v_top_btn = QPushButton()
        self.v_top_btn.setIcon(icon("icons/align_top.png"))
        self.v_top_btn.setToolTip("Align Top")

        self.v_middle_btn = QPushButton()
        self.v_middle_btn.setIcon(icon("icons/align_middle.png"))
        self.v_middle_btn.setToolTip("Align Middle")

        self.v_bottom_btn = QPushButton()
        self.v_bottom_btn.setIcon(icon("icons/align_bottom.png"))
        self.v_bottom_btn.setToolTip("Align Bottom")

        for btn in [self.v_top_btn, self.v_middle_btn, self.v_bottom_btn]:
            btn.setCheckable(True)
            btn.setFixedSize(32, 28)
            self.v_align_group.addButton(btn)
            layout.addWidget(btn)
        self.v_align_group.setExclusive(True)

        add_separator()

        # Wrap Text
        self.wrap_text_btn = QPushButton("Wrap Text")
        self.wrap_text_btn.setIcon(icon("icons/wrap_text.png"))
        self.wrap_text_btn.setCheckable(True)
        self.wrap_text_btn.setToolTip("Wrap Text in Cell")
        layout.addWidget(self.wrap_text_btn)

        add_separator()

        # Merge & Center
        self.merge_center_btn = QPushButton("Merge & Center")
        self.merge_center_btn.setIcon(icon("icons/merge_cells.png"))
        self.merge_center_btn.setCheckable(True)
        self.merge_center_btn.setToolTip("Merge & Center Selected Cells")
        layout.addWidget(self.merge_center_btn)

        layout.addStretch()

    def connect_signals(self):
        self.h_left_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignLeft))
        self.h_center_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignHCenter))
        self.h_right_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignRight))

        self.v_top_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignTop))
        self.v_middle_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignVCenter))
        self.v_bottom_btn.clicked.connect(lambda: self.emit_alignment(Qt.AlignmentFlag.AlignBottom))

        self.wrap_text_btn.clicked.connect(self.wrapTextChanged.emit)
        self.merge_center_btn.clicked.connect(self.mergeChanged.emit)

    def emit_alignment(self, alignment):
        self.alignmentChanged.emit(alignment)
        self.update_button_states()

    def update_for_cell(self, cell):
        self.current_cell = cell
        if cell is None:
            return

        alignment = cell.alignment()
        self.h_left_btn.setChecked(alignment & Qt.AlignmentFlag.AlignLeft)
        self.h_center_btn.setChecked(alignment & Qt.AlignmentFlag.AlignHCenter)
        self.h_right_btn.setChecked(alignment & Qt.AlignmentFlag.AlignRight)

        self.v_top_btn.setChecked(alignment & Qt.AlignmentFlag.AlignTop)
        self.v_middle_btn.setChecked(alignment & Qt.AlignmentFlag.AlignVCenter)
        self.v_bottom_btn.setChecked(alignment & Qt.AlignmentFlag.AlignBottom)

        self.wrap_text_btn.setChecked(cell.wrapText())
        self.merge_center_btn.setChecked(cell.columnSpan() > 1 or cell.rowSpan() > 1)

    def update_button_states(self):
        if self.current_cell:
            self.update_for_cell(self.current_cell)
