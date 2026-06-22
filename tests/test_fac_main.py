"""Unit tests for fac_main.py – Factory (공장) management."""
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

from fac_main import Config, EstateMaster_V13_6


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
        assert Config.PREFIX_MAP == {"공장": "FCT"}

    def test_db_path(self):
        assert Config.DB_PATH.endswith("estate_db.xlsx")

    def test_photo_base_dir(self):
        assert Config.PHOTO_BASE_DIR.endswith("property_photos")

    def test_sns_footer(self):
        assert "성균관" in Config.SNS_FOOTER


# ── category_config ───────────────────────────────────────────────

class TestCategoryConfig:
    def test_factory_category_exists(self, widget):
        assert "공장" in widget.category_config

    def test_factory_g2_fields(self, widget):
        g2 = widget.category_config["공장"]["G2"]
        assert "사용승인일" in g2
        assert "대지면적(㎡)" in g2
        assert "건축면적(㎡)" in g2
        assert "층수" in g2

    def test_factory_g3_fields(self, widget):
        g3 = widget.category_config["공장"]["G3"]
        assert "거래가" in g3
        assert "보증금" in g3
        assert "전력(kw/h)" in g3
        assert "호이스트" in g3

    def test_factory_mkt_fields(self, widget):
        mkt = widget.category_config["공장"]["MKT"]
        assert "광고상황" in mkt
        assert "사진,영상링크" in mkt


# ── _find_private_col ─────────────────────────────────────────────

class TestFindPrivateCol:
    def test_detects_private_columns(self, widget):
        df = pd.DataFrame(columns=["매물ID", "소유자", "미디어경로(사진/영상)", "보증금"])
        prv, med, nums = widget._find_private_col(df)
        assert prv == 2
        assert med == 3
        assert 4 in nums

    def test_multiple_numeric_columns(self, widget):
        df = pd.DataFrame(columns=["대지면적(㎡)", "매매가", "월세"])
        _, _, nums = widget._find_private_col(df)
        assert len(nums) == 3


# ── calculate_factory_totals ──────────────────────────────────────

class TestCalculateFactoryTotals:
    def test_basic_calculation(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(2, 6)
        table.setItem(0, 4, QTableWidgetItem("3000"))
        table.setItem(0, 5, QTableWidgetItem("200"))
        table.setItem(1, 4, QTableWidgetItem("2000"))
        table.setItem(1, 5, QTableWidgetItem("100"))
        widget.table_factory = table

        price_edit = QLineEdit()
        price_edit.setText("50000")
        widget.all_fields["G3"]["매매가"] = price_edit

        widget.lbl_factory_totals = QLabel()
        widget.calculate_factory_totals()
        text = widget.lbl_factory_totals.text()
        assert "보증금 5,000만원" in text
        assert "월세 300만원" in text
        # yield = (300*12)/(50000-5000)*100 = 8.00%
        assert "8.00%" in text

    def test_empty_table(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QLineEdit, QLabel
        widget.table_factory = QTableWidget(0, 6)
        widget.all_fields["G3"]["매매가"] = QLineEdit()
        widget.all_fields["G3"]["매매가"].setText("10000")
        widget.lbl_factory_totals = QLabel()
        widget.calculate_factory_totals()
        assert "0.00%" in widget.lbl_factory_totals.text()


# ── format_comma ──────────────────────────────────────────────────

class TestFormatComma:
    def test_formats_large_number(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G3"]["거래가"] = w
        w.setText("1000000")
        widget.format_comma("G3", "거래가")
        assert w.text() == "1,000,000"

    def test_single_digit(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G3"]["거래가"] = w
        w.setText("5")
        widget.format_comma("G3", "거래가")
        assert w.text() == "5"


# ── format_phone ──────────────────────────────────────────────────

class TestFormatPhone:
    def test_mobile_number(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G2"]["연락처"] = w
        w.setText("01098765432")
        widget.format_phone("G2", "연락처")
        assert w.text() == "010-9876-5432"

    def test_landline(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G2"]["연락처"] = w
        w.setText("0311234567")
        widget.format_phone("G2", "연락처")
        assert w.text() == "031-1234-567"

    def test_short_number(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G2"]["연락처"] = w
        w.setText("02123")
        widget.format_phone("G2", "연락처")
        assert w.text() == "02123"
