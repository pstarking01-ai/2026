"""Unit tests for com_main.py – Commercial property (상가) management."""
import os
import sys
import types
import pytest
import pandas as pd

# Stub the missing reception.estate_db module before importing com_main
_reception = types.ModuleType("reception")
_estate_db = types.ModuleType("reception.estate_db")
_estate_db.update_apartment_in_excel = lambda *a, **kw: None
_reception.estate_db = _estate_db
sys.modules.setdefault("reception", _reception)
sys.modules.setdefault("reception.estate_db", _estate_db)

from com_main import Config, EstateMaster_V13_6


@pytest.fixture
def widget(qapp):
    """Create an EstateMaster instance for testing.

    The source code has a bug where initUI() accesses self.emd_ri_list
    before load_apt_master_db() sets it (and load_apt_master_db is never
    called from __init__). We patch it to unblock widget creation.
    """
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
        assert Config.PREFIX_MAP == {"상가": "COM"}

    def test_db_path_ends_with_xlsx(self):
        assert Config.DB_PATH.endswith("estate_db.xlsx")

    def test_photo_base_dir(self):
        assert Config.PHOTO_BASE_DIR.endswith("property_photos")

    def test_sns_footer_contains_url(self):
        assert "https://pstarking01-ai.github.io/2026/" in Config.SNS_FOOTER
        assert "성균관" in Config.SNS_FOOTER


# ── category_config ───────────────────────────────────────────────

class TestCategoryConfig:
    def test_commercial_category_exists(self, widget):
        assert "상가" in widget.category_config

    def test_commercial_g2_fields(self, widget):
        g2 = widget.category_config["상가"]["G2"]
        assert "준공연도" in g2
        assert "용도지역" in g2
        assert "거래유형" in g2

    def test_commercial_g3_fields(self, widget):
        g3 = widget.category_config["상가"]["G3"]
        assert "거래가(만원)" in g3
        assert "수익율(%)" in g3

    def test_commercial_mkt_fields(self, widget):
        mkt = widget.category_config["상가"]["MKT"]
        assert "광고상황" in mkt
        assert "사진,영상링크" in mkt


# ── format_comma ──────────────────────────────────────────────────

class TestFormatComma:
    def test_formats_number(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G3"]["거래가(만원)"] = w
        w.setText("50000")
        widget.format_comma("G3", "거래가(만원)")
        assert w.text() == "50,000"

    def test_already_formatted(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G3"]["거래가(만원)"] = w
        w.setText("50,000")
        widget.format_comma("G3", "거래가(만원)")
        assert w.text() == "50,000"

    def test_non_digit_unchanged(self, widget):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G3"]["거래가(만원)"] = w
        w.setText("abc")
        widget.format_comma("G3", "거래가(만원)")
        assert w.text() == "abc"


# ── format_phone ──────────────────────────────────────────────────

class TestFormatPhone:
    @pytest.mark.parametrize("raw, expected", [
        ("01012345678", "010-1234-5678"),
        ("0311234567", "031-1234-567"),
        ("021234567", "021-234-567"),
        ("012", "012"),
    ])
    def test_phone_formatting(self, widget, raw, expected):
        from PyQt6.QtWidgets import QLineEdit
        w = QLineEdit()
        widget.all_fields["G2"]["연락처"] = w
        w.setText(raw)
        widget.format_phone("G2", "연락처")
        assert w.text() == expected


# ── _find_private_col ─────────────────────────────────────────────

class TestFindPrivateCol:
    def test_finds_private_columns(self, widget):
        df = pd.DataFrame(columns=["매물ID", "단지명", "지번", "소유자", "미디어경로(사진/영상)", "면적"])
        prv_idx, med_idx, num_cols = widget._find_private_col(df)
        assert prv_idx == 3  # "지번" is at 1-based index 3
        assert med_idx == 5  # "미디어경로" at index 5
        assert 6 in num_cols  # "면적" at index 6

    def test_no_private_columns(self, widget):
        df = pd.DataFrame(columns=["매물ID", "단지명"])
        prv_idx, med_idx, num_cols = widget._find_private_col(df)
        assert prv_idx == 999
        assert med_idx == -1
        assert num_cols == []

    def test_numeric_keywords(self, widget):
        df = pd.DataFrame(columns=["대지면적(㎡)", "매매가", "보증금", "월세"])
        _, _, num_cols = widget._find_private_col(df)
        assert 1 in num_cols
        assert 2 in num_cols
        assert 3 in num_cols
        assert 4 in num_cols


# ── calculate_commercial_totals (yield calculation) ───────────────

class TestCalculateCommercialTotals:
    def _setup_table(self, widget, rows):
        """Set up commercial detail table with given row data."""
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(len(rows), 7)
        for r, row_data in enumerate(rows):
            for c, val in enumerate(row_data):
                table.setItem(r, c, QTableWidgetItem(str(val)))
        widget.table = table

        price_edit = QLineEdit()
        widget.all_fields["G3"]["희망매매가(만원)"] = price_edit

        yield_edit = QLineEdit()
        widget.all_fields["G3"]["수익율(%)"] = yield_edit

        widget.lbl_totals = QLabel()
        return price_edit

    def test_basic_yield(self, widget):
        # 전세보증금=1000, 월세보증금=0, 월세=50
        price_edit = self._setup_table(widget, [["", "", "", "", "1000", "0", "50"]])
        price_edit.setText("10000")  # 매매가 1억

        widget.calculate_commercial_totals()
        text = widget.lbl_totals.text()
        assert "보증금 1,000만원" in text
        assert "월세 50만원" in text
        # yield = (50*12) / (10000-1000) * 100 = 6.67%
        assert "6.67%" in text

    def test_zero_investment(self, widget):
        price_edit = self._setup_table(widget, [["", "", "", "", "10000", "0", "100"]])
        price_edit.setText("10000")  # price == deposit => investment=0
        widget.calculate_commercial_totals()
        text = widget.lbl_totals.text()
        assert "0.00%" in text

    def test_empty_table(self, widget):
        self._setup_table(widget, [])
        from PyQt6.QtWidgets import QLineEdit
        widget.all_fields["G3"]["희망매매가(만원)"].setText("5000")
        widget.calculate_commercial_totals()
        text = widget.lbl_totals.text()
        assert "보증금 0만원" in text


# ── calculate_factory_totals ──────────────────────────────────────

class TestCalculateFactoryTotals:
    def test_factory_yield(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(1, 6)
        table.setItem(0, 4, QTableWidgetItem("2000"))
        table.setItem(0, 5, QTableWidgetItem("100"))
        widget.table_factory = table

        price_edit = QLineEdit()
        price_edit.setText("20000")
        widget.all_fields["G3"]["매매가"] = price_edit

        widget.lbl_factory_totals = QLabel()
        widget.calculate_factory_totals()
        text = widget.lbl_factory_totals.text()
        # yield = (100*12)/(20000-2000)*100 = 6.67%
        assert "6.67%" in text
        assert "보증금 2,000만원" in text


# ── calculate_house_totals ────────────────────────────────────────

class TestCalculateHouseTotals:
    def test_house_yield(self, widget):
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QLabel
        table = QTableWidget(1, 7)
        table.setItem(0, 4, QTableWidgetItem("5000"))  # jeonse
        table.setItem(0, 5, QTableWidgetItem("0"))      # w_dep
        table.setItem(0, 6, QTableWidgetItem("80"))      # monthly
        widget.table_house = table

        price_edit = QLineEdit()
        price_edit.setText("30000")
        widget.all_fields["G3"]["매매가"] = price_edit
        widget.lbl_house_totals = QLabel()

        widget.calculate_house_totals()
        text = widget.lbl_house_totals.text()
        # yield = (80*12)/(30000-5000)*100 = 3.84%
        assert "3.84%" in text
