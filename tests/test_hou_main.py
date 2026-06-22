"""Unit tests for hou_main.py – House (주택) management."""
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

from hou_main import Config, EstateMaster_V13_6


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
        assert Config.PREFIX_MAP == {"주택": "HOU"}

    def test_db_path(self):
        assert Config.DB_PATH.endswith("estate_db.xlsx")

    def test_photo_base_dir(self):
        assert Config.PHOTO_BASE_DIR.endswith("property_photos")

    def test_sns_footer(self):
        assert "성균관" in Config.SNS_FOOTER
        assert "https://pstarking01-ai.github.io/2026/" in Config.SNS_FOOTER


# ── category_config ───────────────────────────────────────────────

class TestCategoryConfig:
    def test_house_category_exists(self, widget):
        assert "주택" in widget.category_config

    def test_house_g2_fields(self, widget):
        g2 = widget.category_config["주택"]["G2"]
        assert "준공연도" in g2
        assert "대지면적(㎡)" in g2
        assert "거래유형" in g2
        assert "거주형태" in g2

    def test_house_g3_fields(self, widget):
        g3 = widget.category_config["주택"]["G3"]
        assert "거래가" in g3
        assert "전세가" in g3
        assert "보증금/월세" in g3

    def test_house_mkt_fields(self, widget):
        mkt = widget.category_config["주택"]["MKT"]
        assert "광고상황" in mkt
        assert "사진,영상링크" in mkt


# ── _find_private_col ─────────────────────────────────────────────

class TestFindPrivateCol:
    def test_detects_owner(self, widget):
        df = pd.DataFrame(columns=["매물ID", "소유자"])
        prv, _, _ = widget._find_private_col(df)
        assert prv == 2

    def test_detects_media_col(self, widget):
        df = pd.DataFrame(columns=["A", "미디어경로(사진/영상)"])
        _, med, _ = widget._find_private_col(df)
        assert med == 2


# ── calculate_house_totals ────────────────────────────────────────

class TestCalculateHouseTotals:
    def test_yield_calculation(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(1, 7)
        table.setItem(0, 4, QTableWidgetItem("3000"))  # jeonse deposit
        table.setItem(0, 5, QTableWidgetItem("0"))      # monthly deposit
        table.setItem(0, 6, QTableWidgetItem("50"))      # monthly rent
        widget.table_house = table

        price_edit = QLineEdit()
        price_edit.setText("20000")
        widget.all_fields["G3"]["매매가"] = price_edit
        widget.lbl_house_totals = QLabel()

        widget.calculate_house_totals()
        text = widget.lbl_house_totals.text()
        # yield = (50*12)/(20000-3000)*100 = 3.53%
        assert "3.53%" in text
        assert "보증금 3,000만원" in text
        assert "월세 50만원" in text

    def test_multiple_rows(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(2, 7)
        table.setItem(0, 4, QTableWidgetItem("1000"))
        table.setItem(0, 5, QTableWidgetItem("500"))
        table.setItem(0, 6, QTableWidgetItem("30"))
        table.setItem(1, 4, QTableWidgetItem("2000"))
        table.setItem(1, 5, QTableWidgetItem("0"))
        table.setItem(1, 6, QTableWidgetItem("40"))
        widget.table_house = table

        price_edit = QLineEdit()
        price_edit.setText("25000")
        widget.all_fields["G3"]["매매가"] = price_edit
        widget.lbl_house_totals = QLabel()

        widget.calculate_house_totals()
        text = widget.lbl_house_totals.text()
        # total_deposit = 1000+500+2000 = 3500
        assert "보증금 3,500만원" in text
        # total_monthly = 30+40 = 70
        assert "월세 70만원" in text

    def test_zero_price(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(1, 7)
        table.setItem(0, 4, QTableWidgetItem("0"))
        table.setItem(0, 5, QTableWidgetItem("0"))
        table.setItem(0, 6, QTableWidgetItem("50"))
        widget.table_house = table

        price_edit = QLineEdit()
        price_edit.setText("")
        widget.all_fields["G3"]["매매가"] = price_edit
        widget.lbl_house_totals = QLabel()

        widget.calculate_house_totals()
        text = widget.lbl_house_totals.text()
        assert "0.00%" in text


# ── format_comma ──────────────────────────────────────────────────

class TestFormatComma:
    def test_formats_number(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G3"]["거래가"] = w
        w.setText("25000")
        widget.format_comma("G3", "거래가")
        assert w.text() == "25,000"


# ── format_phone ──────────────────────────────────────────────────

class TestFormatPhone:
    def test_mobile(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G2"]["연락처"] = w
        w.setText("01011112222")
        widget.format_phone("G2", "연락처")
        assert w.text() == "010-1111-2222"

    def test_landline(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G2"]["연락처"] = w
        w.setText("0311234567")
        widget.format_phone("G2", "연락처")
        assert w.text() == "031-1234-567"
