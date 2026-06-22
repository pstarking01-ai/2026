from PyQt6.QtWidgets import (
    QGroupBox, QGridLayout, QLabel, QLineEdit, QComboBox,
    QPushButton, QVBoxLayout, QHBoxLayout, QTableWidget, QHeaderView,
)
from PyQt6.QtCore import Qt
from shared.config import Config


def create_box(title, color, labels, selected_media, select_media_callback,
               dropdown_data=None, btn_add_media_ref=None):
    """Create a labeled input group box with dynamic fields.

    Returns (QGroupBox, field_dict, btn_add_media) where btn_add_media is the
    media button widget if one was created, else None.
    """
    if dropdown_data is None:
        dropdown_data = Config.DROPDOWN_DATA

    box = QGroupBox(title)
    box.setStyleSheet(f"""
        QGroupBox {{ 
            background-color: {color}; 
            border: 1px solid #c0c0c0; 
            border-radius: 8px; 
            margin-top: 20px; 
            font-weight: bold; 
        }}
        QGroupBox::title {{ 
            subcontrol-origin: margin; 
            subcontrol-position: top left;
            left: 10px; 
            padding: 0 10px; 
            top: 0px; 
            font-size: 14px;
            color: #333333;
        }}
    """)

    layout = QGridLayout()
    layout.setVerticalSpacing(15)
    layout.setHorizontalSpacing(10)
    layout.setContentsMargins(15, 30, 15, 15)

    field_dict = {}
    btn_media = None
    for i, text in enumerate(labels):
        lbl = QLabel(text)
        lbl.setMinimumWidth(90)
        lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        if text == "사진,영상링크":
            edit = QPushButton(f"사진,영상링크 ({len(selected_media)})")
            edit.setFixedHeight(32)
            edit.setStyleSheet(
                "background-color: #f39c12; color: white; font-weight: bold; border-radius: 4px;"
            )
            edit.clicked.connect(select_media_callback)
            btn_media = edit
        elif any(key in text for key in dropdown_data):
            match_key = next(key for key in dropdown_data if key in text)
            edit = QComboBox()
            edit.addItems(dropdown_data[match_key])
            edit.setFixedHeight(32)
            edit.setStyleSheet(
                "background-color: white; border: 1px solid #ced4da; border-radius: 4px;"
            )
        else:
            edit = QLineEdit()
            edit.setFixedHeight(32)
            edit.setStyleSheet(
                "background-color: white; border: 1px solid #ced4da; border-radius: 4px; padding: 2px 5px;"
            )

        field_dict[text] = edit
        layout.addWidget(lbl, i, 0)
        layout.addWidget(edit, i, 1)

    box.setLayout(layout)
    return box, field_dict, btn_media


def create_detail_table_ui(title, headers, col_count, add_callback, del_callback,
                           item_changed_callback, totals_label_text, totals_color,
                           add_btn_text="+ 행 추가", del_btn_text="- 행 삭제"):
    """Create a detail table group box with add/delete buttons and a totals label.

    Returns (QGroupBox, QTableWidget, QLabel).
    """
    group = QGroupBox(title)
    group.setStyleSheet(
        "QGroupBox { font-weight: bold; border: 1px solid #ced4da; margin-top: 30px; }"
    )
    layout = QVBoxLayout()
    layout.setContentsMargins(15, 35, 15, 15)
    layout.setSpacing(10)

    table = QTableWidget(0, col_count)
    table.setHorizontalHeaderLabels(headers)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    table.horizontalHeader().setStretchLastSection(True)
    table.setFixedHeight(200)
    table.itemChanged.connect(item_changed_callback)

    btn_layout = QHBoxLayout()
    btn_add = QPushButton(add_btn_text)
    btn_del = QPushButton(del_btn_text)
    btn_add.clicked.connect(add_callback)
    btn_del.clicked.connect(del_callback)
    btn_layout.addWidget(btn_add)
    btn_layout.addWidget(btn_del)
    btn_layout.addStretch()

    lbl_totals = QLabel(totals_label_text)
    lbl_totals.setStyleSheet(f"font-weight: bold; color: {totals_color}; font-size: 14px;")

    layout.addWidget(table)
    layout.addLayout(btn_layout)
    layout.addWidget(lbl_totals)
    group.setLayout(layout)
    group.setVisible(False)
    return group, table, lbl_totals
