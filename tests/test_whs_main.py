"""Unit tests for whs_main.py – Warehouse (창고) management."""
import sys
import types
import pytest
import pandas as pd

# Stub the missing reception.estate_db module
_reception = types.ModuleType("reception")
_estate_db = types.ModuleType("reception.estate_db")
_estate_db.update_apartment_in_excel = lambda *a, **kw: None
_reception.estate_db = _estate_db
sys.modules.setdefault("reception", _reception)
sys.modules.setdefault("reception.estate_db", _estate_db)

from whs_main import Config, EstateMaster_V13_6


@pytest.fixture
def widget(qapp):
    original_init_ui = EstateMaster_V13_6.initUI

    def patched_init_ui(self):
        self.emd_ri_list = []
        self.emd_ri_to_full_reg_map = {}
        self.apt_master_db = {}
        self.complex_dong_ho_map = {}
        self.apt_unit_details = {}
        original_init_ui(self)

    EstateMaster_V13_6.initUI = patched_init_ui
    try:
        w = EstateMaster_V13_6()
    finally:
        EstateMaster_V13_6.initUI = original_init_ui
    return w


# ── Config ────────────────────────────────────────────────────────

class TestConfig:
    def test_prefix_map(self):
        assert Config.PREFIX_MAP == {"창고": "WHS"}

    def test_db_path(self):
        assert Config.DB_PATH.endswith("estate_db.xlsx")

    def test_photo_base_dir(self):
        assert Config.PHOTO_BASE_DIR.endswith("property_photos")

    def test_sns_footer(self):
        assert "성균관" in Config.SNS_FOOTER
        assert "https://pstarking01-ai.github.io/2026/" in Config.SNS_FOOTER


# ── category_config ───────────────────────────────────────────────

class TestCategoryConfig:
    def test_warehouse_category_exists(self, widget):
        assert "창고" in widget.category_config

    def test_warehouse_g2_fields(self, widget):
        g2 = widget.category_config["창고"]["G2"]
        assert "사용승인일" in g2
        assert "대지면적(㎡)" in g2
        assert "건축면적(㎡)" in g2
        assert "연면적(㎡)" in g2
        assert "층수" in g2

    def test_warehouse_g3_fields(self, widget):
        g3 = widget.category_config["창고"]["G3"]
        assert "거래가" in g3
        assert "보증금" in g3
        assert "월세" in g3
        assert "전력(kw/h)" in g3

    def test_warehouse_mkt_fields(self, widget):
        mkt = widget.category_config["창고"]["MKT"]
        assert "광고상황" in mkt
        assert "사진,영상링크" in mkt


# ── _find_private_col ─────────────────────────────────────────────

class TestFindPrivateCol:
    def test_detects_private_and_media(self, widget):
        df = pd.DataFrame(columns=["매물ID", "지번", "미디어경로(사진/영상)", "면적"])
        prv, med, nums = widget._find_private_col(df)
        assert prv == 2
        assert med == 3
        assert 4 in nums

    def test_owner_as_private(self, widget):
        df = pd.DataFrame(columns=["A", "B", "소유자"])
        prv, _, _ = widget._find_private_col(df)
        assert prv == 3

    def test_all_numeric_keywords(self, widget):
        df = pd.DataFrame(columns=["면적", "개별공시지가", "매매가", "보증금", "월세"])
        _, _, nums = widget._find_private_col(df)
        assert len(nums) == 5


# ── calculate_factory_totals (warehouse uses same method) ─────────

class TestCalculateFactoryTotals:
    def test_basic_yield(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(1, 6)
        table.setItem(0, 4, QTableWidgetItem("1000"))
        table.setItem(0, 5, QTableWidgetItem("50"))
        widget.table_factory = table

        price_edit = QLineEdit()
        price_edit.setText("10000")
        widget.all_fields["G3"]["매매가"] = price_edit
        widget.lbl_factory_totals = QLabel()

        widget.calculate_factory_totals()
        text = widget.lbl_factory_totals.text()
        # yield = (50*12)/(10000-1000)*100 = 6.67%
        assert "6.67%" in text
        assert "보증금 1,000만원" in text
        assert "월세 50만원" in text

    def test_no_price(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(1, 6)
        table.setItem(0, 4, QTableWidgetItem("0"))
        table.setItem(0, 5, QTableWidgetItem("100"))
        widget.table_factory = table

        price_edit = QLineEdit()
        price_edit.setText("")
        widget.all_fields["G3"]["매매가"] = price_edit
        widget.lbl_factory_totals = QLabel()

        widget.calculate_factory_totals()
        text = widget.lbl_factory_totals.text()
        assert "0.00%" in text


# ── calculate_land_totals ─────────────────────────────────────────

class TestCalculateLandTotals:
    def test_land_totals(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(1, 10)
        table.setItem(0, 7, QTableWidgetItem("1000"))
        table.setItem(0, 9, QTableWidgetItem("50000"))
        widget.table_land = table

        area_field = QLineEdit()
        price_field = QLineEdit()
        widget.all_fields["G2"]["면적(㎡)"] = area_field
        widget.all_fields["G3"]["매매가"] = price_field
        widget.lbl_land_totals = QLabel()

        widget.calculate_land_totals()
        text = widget.lbl_land_totals.text()
        assert "1,000" in text
        assert "50,000" in text


# ── format_comma ──────────────────────────────────────────────────

class TestFormatComma:
    def test_formats_number(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G3"]["거래가"] = w
        w.setText("99999")
        widget.format_comma("G3", "거래가")
        assert w.text() == "99,999"


# ── format_phone ──────────────────────────────────────────────────

class TestFormatPhone:
    def test_mobile(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G2"]["연락처"] = w
        w.setText("01055556666")
        widget.format_phone("G2", "연락처")
        assert w.text() == "010-5555-6666"
