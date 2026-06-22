"""Base class for the estate management widget.

Subclasses provide property-specific configuration (prefix, category fields,
detail table layout, etc.) via class attributes or __init__ parameters.
All shared UI, data-management, formatting, and Excel logic lives here.
"""
import os
import datetime
import shutil

import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox,
    QGroupBox, QGridLayout, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QScrollArea, QFileDialog, QCheckBox, QTabWidget,
)
from PyQt6.QtCore import Qt

from shared.config import Config
from shared.excel_utils import apply_excel_styles
from shared.format_utils import format_comma, format_phone
from shared.ui_helpers import create_box, create_detail_table_ui


class BaseEstateWidget(QWidget):
    """Abstract base for property-type estate widgets.

    Subclasses MUST set the following class attributes (or pass them to super().__init__):
        property_name   : str   – Korean name shown in UI (e.g. "상가")
        prefix          : str   – ID prefix (e.g. "COM")
        category_config : dict  – {property_name: {"G2": [...], "G3": [...], "MKT": [...]}}
        detail_config   : dict  – see _default_detail_config() for schema
        g4_labels       : list  – labels for the private-info section
    """

    property_name = ""
    prefix = ""
    category_config = {}
    detail_config = {}
    g4_labels = []

    def __init__(self):
        super().__init__()
        self.db_path = Config.DB_PATH
        self.photo_base_dir = Config.PHOTO_BASE_DIR
        self.selected_media = []
        self.all_fields = {"G2": {}, "G3": {}, "MKT": {}, "G4": {}}
        self.btn_add_media = None

        self.load_apt_master_db()
        self.initUI()

    # ------------------------------------------------------------------
    # UI Initialization
    # ------------------------------------------------------------------
    def initUI(self):
        self.setWindowTitle('성균관 부동산 Master_system v13.6')
        self.setGeometry(100, 100, 1450, 900)
        self.setStyleSheet("background-color: #f8f9fa; font: 12pt 'Malgun Gothic';")

        main_layout = QVBoxLayout()

        self.tabs = QTabWidget()
        self.property_tab = QWidget()
        self.analysis_tab = QWidget()
        self.marketing_tab = QWidget()
        self.consult_tab = QWidget()
        self.activity_tab = QWidget()
        self.contract_tab = QWidget()

        self.tabs.addTab(self.property_tab, "🏠 매물 관리")
        self.tabs.addTab(self.analysis_tab, "📈 데이터 분석")
        self.tabs.addTab(self.marketing_tab, "📊 마케팅")
        self.tabs.addTab(self.consult_tab, " 🧑‍💼 고객상담")
        self.tabs.addTab(self.activity_tab, "🏃 임장활동")
        self.tabs.addTab(self.contract_tab, " 📄 계약관리")

        # --- Tab 1: Property management ---
        prop_layout = QVBoxLayout()
        prop_layout.setSpacing(15)

        top_group = QGroupBox("매물 분류,위치 (Smart Selector)")
        top_group.setStyleSheet("""
            QGroupBox { border: 2px solid #2c3e50; border-radius: 10px; margin-top: 15px; font-weight: bold; background: white; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
        """)
        v_sel_main = QVBoxLayout()
        v_sel_main.setContentsMargins(15, 25, 15, 15)
        v_sel_main.setSpacing(10)

        # Row 1: property ID
        h_sel_1 = QHBoxLayout()
        self.le_m_id = QLineEdit()
        self.le_m_id.setPlaceholderText("매물ID (수정시 입력)")
        self.le_m_id.setFixedWidth(130)
        self.btn_load = QPushButton("조회/수정")
        self.btn_load.setFixedWidth(80)
        self.btn_load.clicked.connect(self.load_data)
        self.btn_load.setStyleSheet("background-color: #3498db; color: white; font-weight: bold;")
        h_sel_1.addWidget(QLabel("매물ID:"))
        h_sel_1.addWidget(self.le_m_id)
        h_sel_1.addWidget(self.btn_load)
        h_sel_1.addStretch()
        v_sel_main.addLayout(h_sel_1)

        # Row 2: address
        h_sel_2 = QHBoxLayout()
        self.le_do = QLineEdit(); self.le_do.setPlaceholderText("도/특별시/광역시"); self.le_do.setFixedWidth(150)
        self.le_si_gun = QLineEdit(); self.le_si_gun.setPlaceholderText("시/군/구"); self.le_si_gun.setFixedWidth(140)
        self.le_emd_ri = QLineEdit(); self.le_emd_ri.setPlaceholderText("읍/면/동/리"); self.le_emd_ri.setFixedWidth(180)
        self.le_jibon = QLineEdit(); self.le_jibon.setPlaceholderText("지번"); self.le_jibon.setFixedWidth(100)
        self.chk_loc_public = QCheckBox("상세정보 공개"); self.chk_loc_public.setChecked(False)

        self.addr_fields = [self.le_do, self.le_si_gun, self.le_emd_ri, self.le_jibon]
        h_sel_2.addWidget(QLabel("주소(위치):"))
        for w in self.addr_fields:
            h_sel_2.addWidget(w)
        h_sel_2.addWidget(self.chk_loc_public)
        v_sel_main.addLayout(h_sel_2)

        # Row 3: type selector + apartment dropdowns
        h_sel_3 = QHBoxLayout()
        self.cb_type = QComboBox()
        self.cb_type.addItems(self.category_config.keys())
        self.cb_type.currentTextChanged.connect(self.on_type_changed)

        self.lbl_reg = QLabel("지역:"); self.cb_region = QComboBox(); self.cb_region.setFixedWidth(130); self.cb_region.setEditable(True)
        self.lbl_cpx = QLabel("단지:"); self.cb_complex = QComboBox(); self.cb_complex.setFixedWidth(160); self.cb_complex.setEditable(True)
        self.lbl_dong = QLabel("동:"); self.cb_dong = QComboBox(); self.cb_dong.setFixedWidth(70); self.cb_dong.setEditable(True)
        self.lbl_ho = QLabel("호:"); self.cb_ho = QComboBox(); self.cb_ho.setFixedWidth(70); self.cb_ho.setEditable(True)

        apt_selection_widgets = [self.cb_region, self.cb_complex, self.cb_dong, self.cb_ho]
        for w in apt_selection_widgets:
            w.setStyleSheet("background-color: white; color: black; border: 1px solid #ced4da; border-radius: 4px;")

        self.cb_region.currentTextChanged.connect(self.update_complex_list)
        self.cb_complex.currentTextChanged.connect(self.update_dong_list)
        self.cb_dong.currentTextChanged.connect(self.update_ho_list)
        self.cb_ho.currentTextChanged.connect(self.auto_fill_apt_info)

        h_sel_3.addWidget(QLabel("매물종류:"))
        h_sel_3.addWidget(self.cb_type)
        for w in [self.lbl_reg, self.cb_region, self.lbl_cpx, self.cb_complex,
                  self.lbl_dong, self.cb_dong, self.lbl_ho, self.cb_ho]:
            h_sel_3.addWidget(w)
        h_sel_3.addStretch()
        v_sel_main.addLayout(h_sel_3)

        self.cb_region.addItems(self.emd_ri_list)
        if self.cb_region.completer():
            self.cb_region.completer().setFilterMode(Qt.MatchFlag.MatchContains)
        if self.cb_complex.completer():
            self.cb_complex.completer().setFilterMode(Qt.MatchFlag.MatchContains)
        self.update_complex_list()

        top_group.setLayout(v_sel_main)
        prop_layout.addWidget(top_group)

        self.entry_panel = QHBoxLayout()
        self.entry_panel.setSpacing(10)
        prop_layout.addLayout(self.entry_panel)

        # Create detail table
        self._init_detail_table()
        prop_layout.addWidget(self.detail_group)

        # Bottom buttons
        self.btn_save = QPushButton("매물등록")
        self.chk_data_analysis = QCheckBox("데이터분석 및 저장")
        self.chk_data_analysis.setChecked(True)
        self.cb_sns_type = QComboBox()
        self.cb_sns_type.addItems(["SNS문자 (개별)", "SNS문자 (다중)"])
        self.btn_sns = QPushButton("📱 SNS 문자 생성")

        self.btn_save.clicked.connect(self.process_save)
        self.btn_sns.clicked.connect(self.handle_sns_action)

        self.btn_save.setFixedHeight(50)
        self.btn_save.setStyleSheet(
            "background-color: #2c3e50; color: white; font-weight: bold; font-size: 16px; border-radius: 5px;"
        )
        self.chk_data_analysis.setFixedHeight(50)
        self.chk_data_analysis.setStyleSheet("""
            QCheckBox { background-color: #34495e; color: white; font-weight: bold;
                        font-size: 16px; border-radius: 5px; padding: 10px 15px; spacing: 10px; }
            QCheckBox::indicator { width: 20px; height: 20px; background-color: white; border-radius: 3px; }
            QCheckBox::indicator:checked { background-color: #2ecc71; }
        """)
        self.cb_sns_type.setFixedHeight(50)
        self.cb_sns_type.setStyleSheet("""
            QComboBox { background-color: #3498db; color: white; font-weight: bold;
                        font-size: 16px; border-radius: 5px; padding-left: 10px; }
            QComboBox::drop-down { border: 0px; }
        """)
        self.btn_sns.setFixedHeight(50)
        self.btn_sns.setStyleSheet(
            "background-color: #27ae60; color: white; font-weight: bold; font-size: 16px; border-radius: 5px;"
        )

        self.output = QTextEdit()
        self.output.setFixedHeight(180)
        self.output.setStyleSheet("border: 1px solid #ced4da; border-radius: 5px; background: white; font-size: 13px;")

        btn_h_layout = QHBoxLayout()
        btn_h_layout.addWidget(self.btn_save)
        btn_h_layout.addWidget(self.chk_data_analysis)
        btn_h_layout.addWidget(self.cb_sns_type)
        btn_h_layout.addWidget(self.btn_sns)

        prop_layout.addLayout(btn_h_layout)
        prop_layout.addWidget(self.output)

        prop_scroll = QScrollArea()
        prop_container = QWidget()
        prop_container.setLayout(prop_layout)
        prop_scroll.setWidget(prop_container)
        prop_scroll.setWidgetResizable(True)
        prop_scroll.setStyleSheet("border: none;")

        tab1_vbox = QVBoxLayout()
        tab1_vbox.addWidget(prop_scroll)
        self.property_tab.setLayout(tab1_vbox)

        # Other tabs
        self.init_analysis_tab()
        self.init_marketing_tab()
        self.init_consult_tab()
        self.init_activity_tab()
        self.init_contract_tab()

        main_layout.addWidget(self.tabs)
        self.on_type_changed()
        self.setLayout(main_layout)

    # ------------------------------------------------------------------
    # Detail table (subclass-configurable)
    # ------------------------------------------------------------------
    def _init_detail_table(self):
        """Create the property-specific detail table from self.detail_config."""
        cfg = self.detail_config
        self.detail_group, self.detail_table, self.lbl_detail_totals = create_detail_table_ui(
            title=cfg["title"],
            headers=cfg["headers"],
            col_count=len(cfg["headers"]),
            add_callback=self._add_detail_row,
            del_callback=self._del_detail_row,
            item_changed_callback=self._on_detail_table_item_changed,
            totals_label_text=cfg.get("totals_label", "합계: 0"),
            totals_color=cfg.get("totals_color", "blue"),
            add_btn_text=cfg.get("add_btn_text", "+ 행 추가"),
            del_btn_text=cfg.get("del_btn_text", "- 행 삭제"),
        )

    def _add_detail_row(self):
        self.detail_table.insertRow(self.detail_table.rowCount())

    def _del_detail_row(self):
        curr = self.detail_table.currentRow()
        if curr >= 0:
            self.detail_table.removeRow(curr)

    def _on_detail_table_item_changed(self, item):
        """Format numeric columns and recalculate totals."""
        col = item.column()
        numeric_cols = self.detail_config.get("numeric_cols", [])
        if col in numeric_cols:
            text = item.text().replace(",", "").strip()
            if text.isdigit():
                formatted = format(int(text), ",")
                if item.text() != formatted:
                    self.detail_table.blockSignals(True)
                    item.setText(formatted)
                    self.detail_table.blockSignals(False)
        self.calculate_detail_totals()

    def calculate_detail_totals(self):
        """Override in subclass for property-specific totals calculation."""
        pass

    # ------------------------------------------------------------------
    # Tab initializers (shared across all property types)
    # ------------------------------------------------------------------
    def init_marketing_tab(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)

        dash_group = QGroupBox("📌 매물 보유 현황")
        dash_group.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 2px solid #34495e; border-radius: 10px; "
            "margin-top: 10px; background: white; }"
        )
        dash_layout = QHBoxLayout()
        dash_layout.setContentsMargins(10, 25, 10, 10)

        self.stat_labels = {}
        for cat in self.category_config.keys():
            v_box = QVBoxLayout(); v_box.setSpacing(0)
            title = QLabel(cat)
            title.setAlignment(Qt.AlignmentFlag.AlignCenter)
            title.setFixedHeight(30)
            val = QLabel("0건")
            val.setStyleSheet("font-size: 12px; font-weight: bold; color: #2980b9;")
            val.setAlignment(Qt.AlignmentFlag.AlignCenter)
            val.setFixedHeight(30)
            self.stat_labels[cat] = val
            v_box.addWidget(title); v_box.addWidget(val)
            dash_layout.addLayout(v_box)

        btn_refresh = QPushButton("🔄 통계 새로고침")
        btn_refresh.clicked.connect(self.refresh_dashboard)
        btn_refresh.setFixedWidth(120)
        btn_open_db = QPushButton("📊 엑셀 DB 열기")
        btn_open_db.clicked.connect(self.open_excel_db)
        btn_open_db.setFixedWidth(120)

        dash_layout.addWidget(btn_refresh)
        dash_layout.addWidget(btn_open_db)
        dash_group.setLayout(dash_layout)
        layout.addWidget(dash_group)

        btn_generate_urgent = QPushButton("🚀 관심고객용 [급매물] SNS 문구 생성")
        btn_generate_urgent.setStyleSheet(
            "background-color: #e74c3c; color: white; font-weight: bold; height: 35px;"
        )
        btn_generate_urgent.clicked.connect(self.generate_urgent_sns_text)

        self.btn_export_blog = QPushButton("🌐 블로그용 공개 매물장 추출 (엑셀)")
        self.btn_export_blog.setStyleSheet(
            "background-color: #27ae60; color: white; font-weight: bold; height: 35px;"
        )
        self.btn_export_blog.clicked.connect(self.export_for_blog)

        btn_box = QHBoxLayout()
        btn_box.addWidget(btn_generate_urgent)
        btn_box.addWidget(self.btn_export_blog)
        layout.addLayout(btn_box)
        layout.addStretch()

        mark_scroll = QScrollArea()
        mark_container = QWidget()
        mark_container.setLayout(layout)
        mark_scroll.setWidget(mark_container)
        mark_scroll.setWidgetResizable(True)
        mark_scroll.setStyleSheet("border: none;")

        tab2_vbox = QVBoxLayout()
        tab2_vbox.addWidget(mark_scroll)
        self.marketing_tab.setLayout(tab2_vbox)

    def init_analysis_tab(self):
        layout = QVBoxLayout()
        label = QLabel("📈 데이터 분석 기능 준비 중입니다.")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
        self.analysis_tab.setLayout(layout)

    def init_consult_tab(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)

        consult_group = QGroupBox("📝 고객 상담일지 작성")
        consult_group.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 1px solid #ced4da; background: #fdfefe; margin-top: 10px; }"
        )
        consult_vbox = QVBoxLayout()
        consult_vbox.setContentsMargins(15, 25, 15, 15)
        consult_vbox.setSpacing(10)

        c_layout = QGridLayout(); c_layout.setVerticalSpacing(10)
        self.le_client_name = QLineEdit(); self.le_client_name.setPlaceholderText("고객명")
        self.le_client_contact = QLineEdit(); self.le_client_contact.setPlaceholderText("연락처")
        self.cb_target_type = QComboBox()
        self.cb_target_type.addItems(list(self.category_config.keys()))
        self.cb_target_type.setFixedWidth(120)
        self.le_target_id = QLineEdit(); self.le_target_id.setPlaceholderText("매물 ID (예: 260528-01)")
        self.cb_consult_type = QComboBox(); self.cb_consult_type.addItems(["WI", "전화"])
        self.te_consult_memo = QTextEdit()
        self.te_consult_memo.setPlaceholderText("상담 상세 내용을 입력하세요. (요청사항, 제한사항, 예산 등)")
        self.te_consult_memo.setFixedHeight(240)

        c_layout.addWidget(QLabel("고객정보:"), 0, 0)
        c_layout.addWidget(self.le_client_name, 0, 1, 1, 2)
        c_layout.addWidget(self.le_client_contact, 0, 3, 1, 2)
        c_layout.addWidget(QLabel("관심매물:"), 1, 0)
        c_layout.addWidget(self.cb_target_type, 1, 1)
        c_layout.addWidget(self.le_target_id, 1, 2)
        c_layout.addWidget(QLabel("방문:"), 1, 3)
        c_layout.addWidget(self.cb_consult_type, 1, 4)
        c_layout.addWidget(QLabel("상담내용:"), 2, 0)
        c_layout.addWidget(self.te_consult_memo, 2, 1, 1, 4)

        consult_vbox.addLayout(c_layout)
        consult_group.setLayout(consult_vbox)
        layout.addWidget(consult_group)

        btn_save_consult = QPushButton("💾 상담내용 DB 저장")
        btn_save_consult.setStyleSheet(
            "background-color: #8e44ad; color: white; font-weight: bold; height: 35px;"
        )
        btn_save_consult.clicked.connect(self.save_consultation_data)

        btn_box = QHBoxLayout()
        btn_box.addWidget(btn_save_consult)
        layout.addLayout(btn_box)
        layout.addStretch()

        cons_scroll = QScrollArea()
        cons_container = QWidget()
        cons_container.setLayout(layout)
        cons_scroll.setWidget(cons_container)
        cons_scroll.setWidgetResizable(True)
        cons_scroll.setStyleSheet("border: none;")

        tab_vbox = QVBoxLayout()
        tab_vbox.addWidget(cons_scroll)
        self.consult_tab.setLayout(tab_vbox)

    def init_activity_tab(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(10, 10, 10, 10)

        # Stage 1
        prep_group = QGroupBox("📋 1. 준비단계 (매물 선정 및 확인)")
        prep_group.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 1px solid #ced4da; background: #fdfefe; margin-top: 10px; }"
        )
        prep_layout = QGridLayout()
        prep_layout.setContentsMargins(15, 25, 15, 15)
        prep_layout.setVerticalSpacing(10)
        self.te_recommend_list = QTextEdit()
        self.te_recommend_list.setPlaceholderText("중개사 추천매물 List up (매물 ID 또는 간략 정보)")
        self.te_recommend_list.setFixedHeight(80)
        self.le_view_pw = QLineEdit(); self.le_view_pw.setPlaceholderText("비밀번호 (공실)")
        self.chk_view_confirmed = QCheckBox("집보기 예약 확인")
        self.le_owner_contact_act = QLineEdit(); self.le_owner_contact_act.setPlaceholderText("소유자 연락처 확인")

        prep_layout.addWidget(QLabel("추천매물 리스트:"), 0, 0)
        prep_layout.addWidget(self.te_recommend_list, 0, 1, 1, 2)
        prep_layout.addWidget(QLabel("공실 비번:"), 1, 0)
        prep_layout.addWidget(self.le_view_pw, 1, 1)
        prep_layout.addWidget(self.chk_view_confirmed, 1, 2)
        prep_layout.addWidget(QLabel("소유자 연락처:"), 2, 0)
        prep_layout.addWidget(self.le_owner_contact_act, 2, 1, 1, 2)
        prep_group.setLayout(prep_layout)
        layout.addWidget(prep_group)

        # Stage 2
        activity_group = QGroupBox("🚶 2. 임장활동 (현장 방문 계획)")
        activity_group.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 1px solid #ced4da; background: #f4f6f7; margin-top: 10px; }"
        )
        activity_grid = QGridLayout()
        activity_grid.setContentsMargins(15, 25, 15, 15)
        activity_grid.setVerticalSpacing(10)
        self.le_route_plan = QLineEdit(); self.le_route_plan.setPlaceholderText("동선 계획")
        self.le_time_plan = QLineEdit(); self.le_time_plan.setPlaceholderText("시간 계획")
        activity_grid.addWidget(QLabel("동선 계획:"), 0, 0)
        activity_grid.addWidget(self.le_route_plan, 0, 1)
        activity_grid.addWidget(QLabel("시간 계획:"), 1, 0)
        activity_grid.addWidget(self.le_time_plan, 1, 1)
        activity_group.setLayout(activity_grid)
        layout.addWidget(activity_group)

        # Stage 3
        review_group = QGroupBox("📝 3. 임장후기 정리 (피드백 관리)")
        review_group.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 1px solid #ced4da; background: #fffde7; margin-top: 10px; }"
        )
        review_layout = QGridLayout()
        review_layout.setContentsMargins(15, 25, 15, 15)
        review_layout.setVerticalSpacing(10)
        self.te_activity_list_mgmt = QTextEdit()
        self.te_activity_list_mgmt.setPlaceholderText("임장 List 관리")
        self.te_activity_list_mgmt.setFixedHeight(80)
        self.le_customer_reaction = QLineEdit(); self.le_customer_reaction.setPlaceholderText("고객 반응 (선호도 등)")
        self.le_hope_price = QLineEdit(); self.le_hope_price.setPlaceholderText("고객 희망 가격")
        self.te_prop_changes = QTextEdit()
        self.te_prop_changes.setPlaceholderText("임장 시 매물 변경사항 체크정리 (수선 필요, 옵션 변동 등)")
        self.te_prop_changes.setFixedHeight(80)

        review_layout.addWidget(QLabel("임장 List 관리:"), 0, 0)
        review_layout.addWidget(self.te_activity_list_mgmt, 0, 1, 1, 3)
        review_layout.addWidget(QLabel("고객 반응:"), 1, 0)
        review_layout.addWidget(self.le_customer_reaction, 1, 1)
        review_layout.addWidget(QLabel("희망 가격:"), 1, 2)
        review_layout.addWidget(self.le_hope_price, 1, 3)
        review_layout.addWidget(QLabel("변경사항 정리:"), 2, 0)
        review_layout.addWidget(self.te_prop_changes, 2, 1, 1, 3)
        review_group.setLayout(review_layout)
        layout.addWidget(review_group)

        layout.addStretch()

        act_scroll = QScrollArea()
        act_container = QWidget()
        act_container.setLayout(layout)
        act_scroll.setWidget(act_container)
        act_scroll.setWidgetResizable(True)
        act_scroll.setStyleSheet("border: none;")

        tab_vbox = QVBoxLayout()
        tab_vbox.addWidget(act_scroll)
        self.activity_tab.setLayout(tab_vbox)

    def init_contract_tab(self):
        layout = QVBoxLayout()
        label = QLabel("📄 계약 관리 기능 준비 중입니다.")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
        self.contract_tab.setLayout(layout)

    # ------------------------------------------------------------------
    # Apartment master DB (shared across all property types)
    # ------------------------------------------------------------------
    def load_apt_master_db(self):
        self.apt_master_db = {}
        self.emd_ri_to_full_reg_map = {}
        self.complex_dong_ho_map = {}
        self.apt_unit_details = {}
        self.emd_ri_list = []

        excel_path = Config.APT_MASTER_EXCEL_PATH
        if not os.path.exists(excel_path):
            return

        try:
            excel_data = pd.read_excel(excel_path, sheet_name=None, engine='openpyxl')

            if 'apt_complex_info' in excel_data:
                df_info = excel_data['apt_complex_info'].fillna('')
                for _, item in df_info.iterrows():
                    reg = str(item.get('지역', '미분류'))
                    emd_ri = str(item.get('읍면동리', ''))
                    c_name = str(item.get('단지명', ''))
                    if emd_ri:
                        if emd_ri not in self.emd_ri_to_full_reg_map:
                            self.emd_ri_to_full_reg_map[emd_ri] = set()
                        self.emd_ri_to_full_reg_map[emd_ri].add(reg)
                    if reg not in self.apt_master_db:
                        self.apt_master_db[reg] = set()
                    self.apt_master_db[reg].add(c_name)
                self.emd_ri_list = sorted(list(self.emd_ri_to_full_reg_map.keys()))

            if 'apt_complex_floor' in excel_data:
                df_floor = excel_data['apt_complex_floor'].fillna('')
                for _, item in df_floor.iterrows():
                    c_name = str(item.get('단지명', ''))
                    dong = str(item.get('동', ''))
                    ho = str(item.get('호', ''))
                    if c_name not in self.complex_dong_ho_map:
                        self.complex_dong_ho_map[c_name] = {}
                    if dong not in self.complex_dong_ho_map[c_name]:
                        self.complex_dong_ho_map[c_name][dong] = []
                    if ho:
                        self.complex_dong_ho_map[c_name][dong].append(ho)

            if 'apt_complex_cost' in excel_data:
                df_cost = excel_data['apt_complex_cost'].fillna('')
                for _, item in df_cost.iterrows():
                    c_name = str(item.get('단지명', ''))
                    dong = str(item.get('동', ''))
                    ho = str(item.get('호', ''))
                    self.apt_unit_details[(c_name, dong, ho)] = {
                        "type": str(item.get('타입', '-')),
                        "supp": str(item.get('공급면적', '-')),
                        "priv": str(item.get('전용면적', '-')),
                    }
        except Exception as e:
            print(f"아파트 데이터 로드 중 오류: {e}")

    def update_complex_list(self):
        selected_emd_ri = self.cb_region.currentText()
        self.cb_complex.clear()
        if selected_emd_ri in self.emd_ri_to_full_reg_map:
            complexes = set()
            for full_reg in self.emd_ri_to_full_reg_map[selected_emd_ri]:
                if full_reg in self.apt_master_db:
                    complexes.update(self.apt_master_db[full_reg])
            self.cb_complex.addItems(sorted(list(complexes)))
            if complexes:
                self.cb_complex.showPopup()
        self.update_dong_list()

    def update_dong_list(self):
        c_name = self.cb_complex.currentText()
        self.cb_dong.clear()
        if c_name in self.complex_dong_ho_map:
            dongs = self.complex_dong_ho_map[c_name].keys()
            self.cb_dong.addItems(sorted(list(dongs)))
        self.update_ho_list()

    def update_ho_list(self):
        c_name = self.cb_complex.currentText()
        dong = self.cb_dong.currentText()
        self.cb_ho.clear()
        if c_name in self.complex_dong_ho_map and dong in self.complex_dong_ho_map[c_name]:
            hos = self.complex_dong_ho_map[c_name][dong]
            self.cb_ho.addItems(sorted(list(set(hos))))
        self.auto_fill_apt_info()

    def auto_fill_apt_info(self):
        selected_emd_ri = self.cb_region.currentText()
        c_name = self.cb_complex.currentText()
        dong = self.cb_dong.currentText()
        ho = self.cb_ho.currentText().strip()

        if selected_emd_ri:
            self.le_emd_ri.setText(selected_emd_ri)

        info = self.apt_unit_details.get((c_name, dong, ho))
        if info:
            g2 = self.all_fields.get("G2", {})
            if "공급면적(㎡)" in g2:
                g2["공급면적(㎡)"].setText(info.get('supp', '-'))
            if "전용면적(㎡)" in g2:
                g2["전용면적(㎡)"].setText(info.get('priv', '-'))

            if all([selected_emd_ri, c_name, dong, ho]) and self.cb_type.currentText() == "아파트":
                update_data = {
                    "공급면적": info.get('supp', '-'),
                    "전용면적": info.get('priv', '-'),
                    "타입": info.get('type', '-'),
                }
                try:
                    from reception.estate_db import update_apartment_in_excel
                    update_apartment_in_excel(self.db_path, c_name, dong, ho, update_data)
                except ImportError:
                    pass

    # ------------------------------------------------------------------
    # Dashboard / DB
    # ------------------------------------------------------------------
    def refresh_dashboard(self):
        if not os.path.exists(self.db_path):
            return
        try:
            with pd.ExcelFile(self.db_path) as reader:
                for cat in self.category_config.keys():
                    if cat in reader.sheet_names:
                        df = pd.read_excel(reader, sheet_name=cat)
                        self.stat_labels[cat].setText(f"{len(df)}건")
        except Exception as e:
            self.output.append(f"📊 통계 새로고침 오류: {e}")

    def open_excel_db(self):
        if not os.path.exists(self.db_path):
            QMessageBox.warning(self, "오류", "데이터베이스 파일이 존재하지 않습니다.")
            return
        try:
            os.startfile(self.db_path)
        except Exception as e:
            QMessageBox.critical(self, "오류", f"엑셀 파일을 열 수 없습니다.\n원인: {str(e)}")
        except:
            pass

    def save_consultation_data(self):
        client = self.le_client_name.text().strip()
        if not client:
            QMessageBox.warning(self, "알림", "고객명을 입력해주세요.")
            return

        record = {
            "상담일자": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "고객명": client,
            "연락처": self.le_client_contact.text(),
            "관심매물ID": f"{self.cb_target_type.currentText()}-{self.le_target_id.text()}",
            "상담구분": self.cb_consult_type.currentText(),
            "상담내용": self.te_consult_memo.toPlainText(),
        }

        try:
            df_new = pd.DataFrame([record])
            if os.path.exists(self.db_path):
                with pd.ExcelWriter(self.db_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    if '상담일지' in writer.book.sheetnames:
                        start_row = writer.book['상담일지'].max_row
                        df_new.to_excel(writer, sheet_name='상담일지', index=False, header=False, startrow=start_row)
                    else:
                        df_new.to_excel(writer, sheet_name='상담일지', index=False)
            else:
                df_new.to_excel(self.db_path, sheet_name='상담일지', index=False)

            QMessageBox.information(self, "저장 완료", f"{client} 고객님의 상담일지가 저장되었습니다.")
            self.le_client_name.clear()
            self.le_client_contact.clear()
            self.te_consult_memo.clear()
        except Exception as e:
            QMessageBox.critical(self, "오류", f"상담일지 저장 실패: {str(e)}")

    # ------------------------------------------------------------------
    # SNS / Marketing
    # ------------------------------------------------------------------
    def generate_urgent_sns_text(self):
        target_id_text = self.le_target_id.text().strip()
        if not target_id_text:
            QMessageBox.warning(self, "알림", "관심매물 ID를 입력한 후 문구를 생성해주세요.")
            return
        m_id = f"{self.cb_target_type.currentText()}-{target_id_text}"

        text = f"""[성균관 부동산 추천매물]
━━━━━━━━━━━━━━
⭐ 특별 관리 매물 안내 ⭐

고객님께서 관심 가지실 만한 
특별한 [급매/우수] 매물이 접수되었습니다!

🆔 매물번호: {m_id}
📍 상세내용: 지금 바로 확인하세요.

💰 특징: 가격 조정 완료 및 로열층 확보!

상세 내역은 문의주시면 
친절하게 안내해 드리겠습니다.\n{Config.SNS_FOOTER}"""
        self.output.clear()
        self.output.append(text)
        self.tabs.setCurrentIndex(0)
        QMessageBox.information(self, "생성 완료", "추천 문구가 생성되었습니다. 하단 출력창을 확인하세요.")

    def export_for_blog(self):
        if not os.path.exists(self.db_path):
            QMessageBox.warning(self, "오류", "데이터베이스 파일이 없습니다.")
            return

        export_path = "매물장_블로그공개용.xlsx"
        try:
            with pd.ExcelFile(self.db_path, engine='openpyxl') as reader:
                with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                    for sheet_name in reader.sheet_names:
                        if sheet_name == "상담일지":
                            continue
                        df = pd.read_excel(reader, sheet_name=sheet_name)
                        if df.empty:
                            continue
                        if "미디어경로(사진/영상)" in df.columns:
                            media_idx = list(df.columns).index("미디어경로(사진/영상)")
                            df_public = df.iloc[:, :media_idx + 1]
                        else:
                            private_triggers = ["지번", "호", "소유자", "연락처"]
                            idx = len(df.columns)
                            for trigger in private_triggers:
                                if trigger in df.columns:
                                    idx = min(idx, list(df.columns).index(trigger))
                            df_public = df.iloc[:, :idx]
                        df_public.to_excel(writer, sheet_name=sheet_name, index=False)

            QMessageBox.information(
                self, "추출 완료",
                f"공개용 파일이 생성되었습니다.\n경로: {os.path.abspath(export_path)}\n\n"
                "이 파일을 구글 시트에 업로드하여 블로그에 게시하세요.",
            )
        except Exception as e:
            QMessageBox.critical(self, "오류", f"추출 중 에러 발생: {str(e)}")

    def select_media(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "매물 사진/영상 선택", "",
            "Media Files (*.png *.jpg *.jpeg *.gif *.mp4 *.mov *.avi *.mkv)",
        )
        if files:
            self.selected_media = files
            if self.btn_add_media:
                self.btn_add_media.setText(f"사진,영상링크 ({len(self.selected_media)})")

    def generate_marketing_text(self):
        type_name = self.cb_type.currentText()
        m_id = self.le_m_id.text() or "ID 미발급"
        is_public = self.chk_loc_public.isChecked()

        addr = f"{self.le_do.text()} {self.le_si_gun.text()} {self.le_emd_ri.text()}"
        if type_name == "아파트":
            complex_name = self.cb_complex.currentText()
            dong = self.cb_dong.currentText()
            ho = self.cb_ho.currentText()
            detail_loc = f"\n📍 위치: {complex_name} {dong}동 " + (f"{ho}호" if is_public else "(호실 비공개)")
        else:
            jibon = self.le_jibon.text()
            detail_loc = f"\n📍 위치: {addr} " + (f"{jibon}" if is_public else "(상세지번 비공개)")

        summary = []
        for group in ["G2", "G3"]:
            for k, v in self.all_fields[group].items():
                val = v.currentText() if isinstance(v, QComboBox) else v.text()
                if val and val != "-":
                    summary.append(f"▪ {k}: {val}")

        text = f"""[성균관 부동산 매물안내]
━━━━━━━━━━━━━━
🏠 매물종류: {type_name}
🆔 매물번호: {m_id}{detail_loc}

📋 주요정보
{chr(10).join(summary)}

━━━━━━━━━━━━━━
✨ {type_name} 전문 중개! 
궁금하신 사항은 언제든 문의주세요.\n{Config.SNS_FOOTER}"""
        self.output.clear()
        self.output.append(text)
        QMessageBox.information(self, "문구 생성 완료", "하단 출력창에 발송용 문구가 생성되었습니다. 복사해서 사용하세요.")

    def add_to_multi_list(self):
        type_name = self.cb_type.currentText()
        m_id = self.le_m_id.text().strip().upper() or "ID미발급"
        addr = self.le_emd_ri.text() or "주소미입력"

        area_w = self.all_fields["G2"].get("면적(㎡)") or self.all_fields["G2"].get("대지면적(㎡)")
        price_w = (self.all_fields["G3"].get("매매가") or self.all_fields["G3"].get("희망매매가(만원)")
                   or self.all_fields["G3"].get("매매가(만원)"))

        area = area_w.text() if area_w else "-"
        price = price_w.text() if price_w else "-"

        extra = ""
        if type_name == "토지":
            jm_w = self.all_fields["G2"].get("지목")
            jm = (jm_w.currentText() if isinstance(jm_w, QComboBox) else jm_w.text()) if jm_w else ""
            if jm and jm != "-":
                extra = f" / {jm}"

        current_text = self.output.toPlainText()
        header = "📢 [성균관 부동산 추천 매물 리스트]\n━━━━━━━━━━━━━━\n"
        footer = f"\n━━━━━━━━━━━━━━\n✨ 문의주시면 친절히 안내해 드립니다.{Config.SNS_FOOTER}"

        if "추천 매물 리스트" not in current_text:
            new_content = header
        else:
            new_content = current_text.split("━━━━━━━━━━━━━━\n✨")[0]

        item_line = f"📍 {type_name}({m_id}): {addr} | {area}㎡ | {price}만원{extra}\n"
        self.output.clear()
        self.output.setText(new_content + item_line + footer)
        self.tabs.setCurrentIndex(0)
        QMessageBox.information(self, "목록 추가", f"{m_id} 매물이 안내 목록에 추가되었습니다.")

    def handle_sns_action(self):
        mode = self.cb_sns_type.currentText()
        if "(개별)" in mode:
            self.generate_marketing_text()
        elif "(다중)" in mode:
            self.add_to_multi_list()

    # ------------------------------------------------------------------
    # Type change handler (rebuilds input fields)
    # ------------------------------------------------------------------
    def on_type_changed(self):
        while self.entry_panel.count():
            item = self.entry_panel.takeAt(0)
            if item.layout():
                while item.layout().count():
                    sub_item = item.layout().takeAt(0)
                    if sub_item.widget():
                        sub_item.widget().deleteLater()
            if item.widget():
                item.widget().deleteLater()

        type_name = self.cb_type.currentText()
        for w in self.addr_fields:
            w.setVisible(True)

        self._update_detail_visibility(type_name)

        cfg = self.category_config[type_name]
        self.box2, self.all_fields["G2"], _ = create_box(
            "상세정보", "#f1f8e9", cfg["G2"], self.selected_media, self.select_media
        )
        self.box3, self.all_fields["G3"], _ = create_box(
            "가격조건", "#fff3e0", cfg["G3"], self.selected_media, self.select_media
        )

        if "MKT" in cfg:
            self.box_mkt, self.all_fields["MKT"], btn = create_box(
                "마케팅 정보", "#e3f2fd", cfg["MKT"], self.selected_media, self.select_media
            )
            if btn:
                self.btn_add_media = btn
        else:
            self.box_mkt = None
            self.all_fields["MKT"] = {}

        self.box4, self.all_fields["G4"], _ = create_box(
            "미공개정보", "#fce4ec", self.g4_labels, self.selected_media, self.select_media
        )

        col1_vbox = QVBoxLayout()
        col1_vbox.addWidget(self.box2)
        col1_vbox.addWidget(self.box3)

        self.entry_panel.addLayout(col1_vbox)
        if self.box_mkt:
            self.entry_panel.addWidget(self.box_mkt)
        self.entry_panel.addWidget(self.box4)

        self._connect_formatters()

    def _update_detail_visibility(self, type_name):
        """Show/hide the detail table based on property type. Override if needed."""
        self.detail_group.setVisible(True)

    def _connect_formatters(self):
        """Connect comma/phone formatters to the current field set."""
        comma_fields_g2 = ["면적(㎡)", "대지면적(㎡)", "건축면적(㎡)", "연면적(㎡)"]
        comma_fields_g3 = [
            "거래가(만원)", "거래가", "보증금/월세", "개별공시지가(㎡)",
            "자기자본", "대출금", "보증금", "월세", "전력(kw/h)",
            "희망매매가(만원)", "매매가(만원)",
        ]
        comma_fields_g4 = ["조정가"]
        phone_fields_g4 = ["소유자연락처", "연락처", "관계인연락처"]

        for field in comma_fields_g2:
            if field in self.all_fields["G2"]:
                self.all_fields["G2"][field].textChanged.connect(
                    lambda _, f=field: format_comma(self.all_fields, "G2", f)
                )

        for field in comma_fields_g3:
            if field in self.all_fields["G3"]:
                self.all_fields["G3"][field].textChanged.connect(
                    lambda _, f=field: format_comma(self.all_fields, "G3", f)
                )

        for field in comma_fields_g4:
            if field in self.all_fields["G4"]:
                self.all_fields["G4"][field].textChanged.connect(
                    lambda _, f=field: format_comma(self.all_fields, "G4", f)
                )

        for field in phone_fields_g4:
            if field in self.all_fields["G4"]:
                self.all_fields["G4"][field].textChanged.connect(
                    lambda _, f=field: format_phone(self.all_fields, "G4", f)
                )

        self._connect_totals_triggers()

    def _connect_totals_triggers(self):
        """Connect price fields to the totals recalculation. Override in subclass."""
        pass

    # ------------------------------------------------------------------
    # Load / Save
    # ------------------------------------------------------------------
    def load_data(self):
        m_id = self.le_m_id.text().strip().upper()
        if not m_id:
            QMessageBox.warning(self, "알림", "조회할 매물ID를 입력해주세요.")
            return
        if not os.path.exists(self.db_path):
            QMessageBox.warning(self, "오류", "데이터베이스 파일이 존재하지 않습니다.")
            return

        try:
            found = False
            with pd.ExcelFile(self.db_path, engine='openpyxl') as reader:
                prefix = m_id.split('-')[0] if '-' in m_id else ""
                if prefix != self.prefix:
                    return
                logical_name = self.property_name
                db_sheet = logical_name

                if db_sheet in reader.sheet_names:
                    df = pd.read_excel(reader, sheet_name=db_sheet)
                    row = df[df['매물ID'] == m_id]
                    if not row.empty:
                        found = True
                        data = row.iloc[0].to_dict()
                        self.cb_type.setCurrentText(logical_name)
                        self.le_do.setText(str(data.get("도", "")))
                        self.le_si_gun.setText(str(data.get("시/군/구", "")))
                        self.le_emd_ri.setText(str(data.get("읍/면/동/리", "")))
                        self.le_jibon.setText(str(data.get("지번", "")))

                        for group_name in ["G2", "G3", "MKT", "G4"]:
                            for label, widget in self.all_fields[group_name].items():
                                if isinstance(widget, QPushButton):
                                    continue
                                val = str(data.get(label, ""))
                                if val == "nan":
                                    val = ""
                                if isinstance(widget, QComboBox):
                                    widget.setCurrentText(val)
                                else:
                                    widget.setText(val)

                self._load_detail_data(reader, m_id, db_sheet)

            if found:
                self.output.append(f"✅ {m_id} 데이터를 성공적으로 불러왔습니다.")
            else:
                QMessageBox.warning(self, "실패", f"ID '{m_id}'에 해당하는 매물을 찾을 수 없습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"조회 중 오류 발생: {str(e)}")

    def _load_detail_data(self, reader, m_id, db_sheet):
        """Load detail table data from the Excel reader. Override for custom headers."""
        detail_sheet = f"{db_sheet}상세"
        if detail_sheet in reader.sheet_names:
            df_detail = pd.read_excel(reader, sheet_name=detail_sheet)
            rows = df_detail[df_detail['매물ID'] == m_id]
            self.detail_table.setRowCount(0)
            headers = self.detail_config["headers"]
            for _, d_row in rows.iterrows():
                r_idx = self.detail_table.rowCount()
                self.detail_table.insertRow(r_idx)
                for c_idx, h in enumerate(headers):
                    val = str(d_row.get(h, ""))
                    if val == "nan":
                        val = ""
                    self.detail_table.setItem(r_idx, c_idx, QTableWidgetItem(val))

    def process_save(self):
        type_name = self.cb_type.currentText()
        pfx = Config.PREFIX_MAP.get(type_name, "ETC")
        m_id = self.le_m_id.text().strip().upper()

        if m_id:
            current_pfx = m_id.split('-')[0]
            if current_pfx in Config.PREFIX_MAP.values() and current_pfx != pfx:
                old_id = m_id
                parts = m_id.split('-')
                if len(parts) >= 3:
                    m_id = f"{pfx}-{parts[1]}-{parts[2]}"
                    self.output.append(f"🔄 카테고리 변경 감지: ID가 {old_id}에서 {m_id}로 변환되었습니다.")

        if not m_id:
            date_str = datetime.datetime.now().strftime("%y%m%d")
            seq = 1
            if os.path.exists(self.db_path):
                try:
                    with pd.ExcelFile(self.db_path, engine='openpyxl') as reader:
                        max_seq = 0
                        date_pattern = f"-{date_str}-"
                        for s_name in reader.sheet_names:
                            df_tmp = pd.read_excel(reader, sheet_name=s_name)
                            if '매물ID' in df_tmp.columns:
                                matched = df_tmp[df_tmp['매물ID'].astype(str).str.contains(date_pattern)]['매물ID'].tolist()
                                for mid in matched:
                                    try:
                                        num = int(str(mid).split('-')[-1])
                                        if num > max_seq:
                                            max_seq = num
                                    except Exception:
                                        continue
                        seq = max_seq + 1
                except Exception:
                    pass
            m_id = f"{pfx}-{date_str}-{seq:02d}"

        # Media handling
        media_path_info = "-"
        if self.selected_media:
            target_dir = os.path.join(self.photo_base_dir, m_id)
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)
            for idx, src_path in enumerate(self.selected_media):
                ext = os.path.splitext(src_path)[1]
                filename = f"{m_id}_{idx+1}{ext}"
                shutil.copy(src_path, os.path.join(target_dir, filename))
            media_path_info = os.path.abspath(target_dir)

        # Collect record
        record = {"매물ID": m_id, "등록일": datetime.datetime.now().strftime("%Y-%m-%d")}
        is_public = self.chk_loc_public.isChecked()
        record.update({
            "도": self.le_do.text(),
            "시/군/구": self.le_si_gun.text(),
            "읍/면/동/리": self.le_emd_ri.text(),
        })

        if is_public:
            record["지번"] = self.le_jibon.text()

        for group_key in ["G2", "G3", "MKT"]:
            group = self.all_fields[group_key]
            for label, widget in group.items():
                if isinstance(widget, QPushButton):
                    continue
                record[label] = widget.currentText() if isinstance(widget, QComboBox) else widget.text()

        record["미디어경로(사진/영상)"] = media_path_info

        if not is_public:
            record["지번"] = self.le_jibon.text()

        for label, widget in self.all_fields["G4"].items():
            record[label] = widget.currentText() if isinstance(widget, QComboBox) else widget.text()

        # Extract detail data
        detail_df = self._extract_detail_data(m_id)

        # Save to Excel
        try:
            self.output.append(f"⏳ {m_id} 데이터를 저장 중입니다...")
            QApplication.processEvents()

            new_df = pd.DataFrame([record])
            sheets = {}

            try:
                if os.path.exists(self.db_path):
                    with pd.ExcelFile(self.db_path, engine='openpyxl') as reader:
                        sheets = {sheet: pd.read_excel(reader, sheet_name=sheet) for sheet in reader.sheet_names}

                    if type_name in sheets:
                        df_main = sheets[type_name]
                        df_main = df_main[df_main['매물ID'] != m_id]
                        combined = pd.concat([df_main, new_df], ignore_index=True)
                        cols = list(record.keys())
                        for c in combined.columns:
                            if c not in cols:
                                cols.append(c)
                        sheets[type_name] = combined[cols]
                    else:
                        sheets[type_name] = new_df

                    if detail_df is not None and not detail_df.empty:
                        d_sheet = f"{type_name}상세"
                        if d_sheet in sheets:
                            sheets[d_sheet] = sheets[d_sheet][sheets[d_sheet]['매물ID'] != m_id]
                            sheets[d_sheet] = pd.concat([sheets[d_sheet], detail_df], ignore_index=True)
                        else:
                            sheets[d_sheet] = detail_df
                else:
                    sheets = {type_name: new_df}
                    if detail_df is not None:
                        sheets[f"{type_name}상세"] = detail_df

                with pd.ExcelWriter(self.db_path, engine='openpyxl') as writer:
                    for s_name, df in sheets.items():
                        df.to_excel(writer, sheet_name=s_name, index=False)
                        ws = writer.sheets[s_name]
                        apply_excel_styles(ws, df)
            except PermissionError:
                raise Exception(
                    f"엑셀 파일('{self.db_path}')이 이미 열려 있습니다.\n엑셀 프로그램을 닫고 다시 시도해 주세요."
                )

            loc_info = f"{record['시/군/구']} {record['읍/면/동/리']}"
            self.output.append(f"✅ {m_id} 저장 성공: {loc_info}")
            QMessageBox.information(self, "성공", f"{m_id} 저장 및 분석이 완료되었습니다.")
            self.clear_inputs()
        except Exception as e:
            self.output.append(f"❌ 에러 발생: {str(e)}")
            QMessageBox.critical(self, "저장 에러", f"데이터를 저장하지 못했습니다.\n원인: {str(e)}")

    def _extract_detail_data(self, m_id):
        """Extract detail table rows as a DataFrame. Override for custom logic."""
        headers = self.detail_config["headers"]
        details = []
        for r in range(self.detail_table.rowCount()):
            row_data = {"매물ID": m_id}
            for c in range(self.detail_table.columnCount()):
                item = self.detail_table.item(r, c)
                row_data[headers[c]] = item.text() if item else ""
            details.append(row_data)
        if details:
            return pd.DataFrame(details)
        return None

    def clear_inputs(self):
        self.le_m_id.clear()
        self.le_do.clear()
        self.le_si_gun.clear()
        self.le_emd_ri.clear()
        self.le_jibon.clear()
        for group in self.all_fields.values():
            for widget in group.values():
                if isinstance(widget, QLineEdit):
                    widget.clear()
                elif isinstance(widget, QComboBox):
                    widget.setCurrentIndex(0)
        self.detail_table.setRowCount(0)
        self.calculate_detail_totals()
