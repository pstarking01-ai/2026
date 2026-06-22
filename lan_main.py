"""Land (토지) property management widget.

lan_main.py uses a simpler single-page layout without the full tab set,
so it doesn't inherit BaseEstateWidget but still reuses shared utilities.
"""
import sys
import os
import datetime

import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QMessageBox,
    QGroupBox, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QScrollArea, QFileDialog, QCheckBox,
)
from PyQt6.QtCore import Qt

from shared.config import Config
from shared.excel_utils import apply_excel_styles
from shared.format_utils import format_comma, format_phone
from shared.ui_helpers import create_box, create_detail_table_ui


class LandEstateWidget(QWidget):
    """Land (토지) registration system."""

    LAND_DROPDOWN_DATA = {
        "광고상황": ["광고전", "진행중", "거래완료", "매물철회"],
        "광고사이트": ["한방", "이실장", "부동산써브"],
        "자체광고 채널": ["홈페이지", "블로그", "유튜브"],
        "용도지역": Config.DROPDOWN_DATA["용도지역"],
        "통신사": Config.DROPDOWN_DATA["통신사"],
        "연락처구분": ["소유자", "관계인", "기타"],
    }

    def __init__(self):
        super().__init__()
        self.db_path = Config.DB_PATH
        self.photo_base_dir = Config.PHOTO_BASE_DIR
        self.selected_media = []
        self.btn_add_media = None
        self.all_fields = {"G2": {}, "G3": {}, "G4": {}}

        self.category_config = {
            "토지": {
                "G2": ["용도지역", "지목", "면적(㎡)", "개별공시지가(㎡)", "매매가(만원))"],
                "G3": ["진입로(폭)", "토목공사", "상/하수도", "전기/통신"],
            }
        }
        self.initUI()

    def initUI(self):
        self.setWindowTitle('성균관 토지 매물 등록 시스템 v1.0')
        self.setGeometry(100, 100, 1450, 900)
        self.setStyleSheet("background-color: #f8f9fa; font: 12pt 'Malgun Gothic';")

        main_layout = QVBoxLayout()
        prop_layout = QVBoxLayout()
        prop_layout.setSpacing(15)

        top_group = QGroupBox("매물 분류 및 위치")
        top_group.setStyleSheet("""
            QGroupBox { border: 2px solid #2c3e50; border-radius: 10px; margin-top: 15px; font-weight: bold; background: white; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
        """)
        v_sel_main = QVBoxLayout()
        v_sel_main.setContentsMargins(15, 25, 15, 15)

        # Row 1: Property ID
        h_sel_1 = QHBoxLayout()
        self.le_m_id = QLineEdit(); self.le_m_id.setPlaceholderText("LND-YYMMDD-NN"); self.le_m_id.setFixedWidth(150)
        self.btn_load = QPushButton("매물 불러오기"); self.btn_load.setFixedWidth(120)
        self.btn_load.clicked.connect(self.load_data)
        self.btn_load.setStyleSheet("background-color: #3498db; color: white; font-weight: bold;")
        h_sel_1.addWidget(QLabel("매물 ID:")); h_sel_1.addWidget(self.le_m_id); h_sel_1.addWidget(self.btn_load)
        h_sel_1.addStretch()
        v_sel_main.addLayout(h_sel_1)

        # Row 2: Address
        h_sel_2 = QHBoxLayout()
        self.le_do = QLineEdit(); self.le_do.setPlaceholderText("특별/특/광/도"); self.le_do.setFixedWidth(150)
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

        top_group.setLayout(v_sel_main)
        prop_layout.addWidget(top_group)

        self.entry_panel = QHBoxLayout()
        prop_layout.addLayout(self.entry_panel)

        # Detail table
        self.land_detail_group, self.table_land, self.lbl_land_totals = create_detail_table_ui(
            title="토지접수 상세 (필지별 현황/단위:㎡/만원)",
            headers=["필지ID", "특별/특/광/도", "시/군/구", "읍/면/동/리", "지번",
                      "용도지역", "지목", "면적(㎡)", "개별공시지가(㎡,원)", "매매가(만원)",
                      "소유자", "메모", "비고"],
            col_count=13,
            add_callback=self._add_land_row,
            del_callback=self._del_land_row,
            item_changed_callback=self.on_land_table_item_changed,
            totals_label_text="합계: 총 면적 0㎡ / 총 매매가 0만원",
            totals_color="#27ae60",
            add_btn_text="+ 필지 추가",
            del_btn_text="- 필지 삭제",
        )
        prop_layout.addWidget(self.land_detail_group)

        # Markdown copy button
        control_h_layout = QHBoxLayout()
        self.btn_copy_markdown = QPushButton("📋 마크다운 테이블 복사")
        self.btn_copy_markdown.setFixedHeight(40)
        self.btn_copy_markdown.setStyleSheet("background-color: #9b59b6; color: white; font-weight: bold;")
        self.btn_copy_markdown.clicked.connect(self.copy_as_markdown)
        control_h_layout.addWidget(self.btn_copy_markdown)
        control_h_layout.addStretch()
        prop_layout.addLayout(control_h_layout)

        self.btn_save = QPushButton("매물등록")
        self.btn_save.clicked.connect(self.process_save)
        self.btn_save.setFixedHeight(50)
        self.btn_save.setStyleSheet(
            "background-color: #2c3e50; color: white; font-weight: bold; font-size: 16px; border-radius: 5px;"
        )

        self.output = QTextEdit()
        self.output.setFixedHeight(180)
        self.output.setStyleSheet("border: 1px solid #ced4da; border-radius: 5px; background: white; font-size: 13px;")

        btn_h_layout = QHBoxLayout()
        btn_h_layout.addWidget(self.btn_save)
        prop_layout.addLayout(btn_h_layout)
        prop_layout.addWidget(self.output)

        prop_scroll = QScrollArea()
        prop_container = QWidget()
        prop_container.setLayout(prop_layout)
        prop_scroll.setWidget(prop_container)
        prop_scroll.setWidgetResizable(True)
        prop_scroll.setStyleSheet("border: none;")

        main_layout.addWidget(prop_scroll)
        self.on_type_changed()
        self.setLayout(main_layout)

    # ------------------------------------------------------------------
    # Detail table helpers
    # ------------------------------------------------------------------
    def _add_land_row(self):
        self.table_land.insertRow(self.table_land.rowCount())

    def _del_land_row(self):
        curr = self.table_land.currentRow()
        if curr >= 0:
            self.table_land.removeRow(curr)

    def on_land_table_item_changed(self, item):
        col = item.column()
        if col in [7, 8, 9]:
            text = item.text().replace(",", "").replace(" ", "").strip()
            if text.isdigit():
                formatted = format(int(text), ",")
                if item.text() != formatted:
                    self.table_land.blockSignals(True)
                    item.setText(formatted)
                    self.table_land.blockSignals(False)
        self.calculate_land_totals()

    def calculate_land_totals(self):
        total_area = 0
        total_price = 0
        for r in range(self.table_land.rowCount()):
            try:
                area_item = self.table_land.item(r, 7)
                price_item = self.table_land.item(r, 9)
                if area_item and area_item.text().replace(',', '').strip().isdigit():
                    total_area += int(area_item.text().replace(',', '').strip())
                if price_item and price_item.text().replace(',', '').strip().isdigit():
                    total_price += int(price_item.text().replace(',', '').strip())
            except Exception:
                continue

        area_field = self.all_fields.get("G2", {}).get("면적(㎡)")
        price_field = self.all_fields.get("G3", {}).get("매매가(만원)")
        if area_field and total_area > 0:
            area_field.blockSignals(True)
            area_field.setText(format(total_area, ","))
            area_field.blockSignals(False)
        if price_field and total_price > 0:
            price_field.blockSignals(True)
            price_field.setText(format(total_price, ","))
            price_field.blockSignals(False)

        self.lbl_land_totals.setText(f"합계: 총 면적 {total_area:,}㎡ / 총 매매가 {total_price:,}만원")

    # ------------------------------------------------------------------
    # Field creation and type change
    # ------------------------------------------------------------------
    def select_media(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "매물 사진/영상 선택", "",
            "Media Files (*.png *.jpg *.jpeg *.gif *.mp4 *.mov *.avi *.mkv)",
        )
        if files:
            self.selected_media = files

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

        self.land_detail_group.setVisible(True)

        cfg = self.category_config["토지"]
        self.box2, self.all_fields["G2"], _ = create_box(
            "상세정보", "#f1f8e9", cfg["G2"], self.selected_media,
            self.select_media, dropdown_data=self.LAND_DROPDOWN_DATA,
        )
        self.box3, self.all_fields["G3"], _ = create_box(
            "기반시설", "#fff3e0", cfg["G3"], self.selected_media,
            self.select_media, dropdown_data=self.LAND_DROPDOWN_DATA,
        )

        g4_labels = ["소유자", "성별", "소유자연락처", "관계인연락처", "조정가", "중개사메모1", "중개사메모2"]
        self.box4, self.all_fields["G4"], _ = create_box(
            "미공개정보", "#fce4ec", g4_labels, self.selected_media,
            self.select_media, dropdown_data=self.LAND_DROPDOWN_DATA,
        )

        col1_vbox = QVBoxLayout()
        col1_vbox.addWidget(self.box2)
        col1_vbox.addWidget(self.box3)
        self.entry_panel.addLayout(col1_vbox)
        self.entry_panel.addWidget(self.box4)

        # Connect formatters
        if "매매가(만원)" in self.all_fields["G3"]:
            self.all_fields["G3"]["매매가(만원)"].textChanged.connect(
                lambda _: format_comma(self.all_fields, "G3", "매매가(만원)")
            )
            self.all_fields["G3"]["매매가(만원)"].textChanged.connect(self.calculate_land_totals)
        if "면적(㎡)" in self.all_fields["G2"]:
            self.all_fields["G2"]["면적(㎡)"].textChanged.connect(
                lambda _: format_comma(self.all_fields, "G2", "면적(㎡)")
            )
        if "개별공시지가(㎡)" in self.all_fields["G3"]:
            self.all_fields["G3"]["개별공시지가(㎡)"].textChanged.connect(
                lambda _: format_comma(self.all_fields, "G3", "개별공시지가(㎡)")
            )
        if "조정가" in self.all_fields["G4"]:
            self.all_fields["G4"]["조정가"].textChanged.connect(
                lambda _: format_comma(self.all_fields, "G4", "조정가")
            )
        if "소유자연락처" in self.all_fields["G4"]:
            self.all_fields["G4"]["소유자연락처"].textChanged.connect(
                lambda _: format_phone(self.all_fields, "G4", "소유자연락처")
            )
        if "관계인연락처" in self.all_fields["G4"]:
            self.all_fields["G4"]["관계인연락처"].textChanged.connect(
                lambda _: format_phone(self.all_fields, "G4", "관계인연락처")
            )

    # ------------------------------------------------------------------
    # Markdown export
    # ------------------------------------------------------------------
    def copy_as_markdown(self):
        try:
            m_id = self.le_m_id.text() or "ID미발급"
            addr = f"{self.le_do.text()} {self.le_si_gun.text()} {self.le_emd_ri.text()} {self.le_jibon.text()}"

            md = f"### 📍 토지 매물 상세 보고서 ({m_id})\n\n"
            md += "#### **[ 1. 기본 매물 정보 ]**\n"
            md += "| **구분 항목** | **상세 정보** |\n| :--- | :--- |\n"
            md += f"| **📍 소재지** | **{addr}** |\n"

            for group in ["G2", "G3"]:
                for k, v in self.all_fields[group].items():
                    val = v.currentText() if isinstance(v, QComboBox) else v.text()
                    if val and val != "-":
                        md += f"| **{k}** | {val} |\n"

            if self.table_land.rowCount() > 0:
                md += "\n#### **[ 2. 필지별 세부 현황 ]**\n"
                headers = ["**필지ID**", "**지번**", "**용도지역**", "**지목**",
                           "**면적(㎡)**", "**매매가(만원)**", "**소유자**", "**메모**", "**비고**"]
                md += "| " + " | ".join(headers) + " |\n"
                md += "| " + " | ".join(["---"] * len(headers)) + " |\n"
                for r in range(self.table_land.rowCount()):
                    row = []
                    for c_idx in [0, 4, 5, 6, 7, 9, 10, 11, 12]:
                        item = self.table_land.item(r, c_idx)
                        row.append(item.text() if item else "")
                    md += "| " + " | ".join(row) + " |\n"

            QApplication.clipboard().setText(md)
            QMessageBox.information(self, "복사 완료", "토지 매물 정보가 마크다운 형식으로 복사되었습니다.")
        except Exception as e:
            QMessageBox.warning(self, "오류", f"마크다운 생성 중 오류: {e}")

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
                if "토지" in reader.sheet_names:
                    df = pd.read_excel(reader, sheet_name="토지")
                    row = df[df['매물ID'] == m_id]
                    if not row.empty:
                        found = True
                        data = row.iloc[0].to_dict()
                        self.le_do.setText(str(data.get("특별/특/광/도", "")))
                        self.le_si_gun.setText(str(data.get("시/군/구", "")))
                        self.le_emd_ri.setText(str(data.get("읍/면/동/리", "")))
                        self.le_jibon.setText(str(data.get("지번", "")))
                        for group_name in ["G2", "G3", "G4"]:
                            for label, widget in self.all_fields[group_name].items():
                                val = str(data.get(label, ""))
                                if val == "nan":
                                    val = ""
                                if isinstance(widget, QComboBox):
                                    widget.setCurrentText(val)
                                else:
                                    widget.setText(val)

                if "토지상세" in reader.sheet_names:
                    df_detail = pd.read_excel(reader, sheet_name="토지상세")
                    rows = df_detail[df_detail['매물ID'] == m_id]
                    self.table_land.setRowCount(0)
                    headers = ["필지ID", "특별/특/광/도", "시/군/구", "읍/면/동/리", "지번",
                               "용도지역", "지목", "면적(㎡)", "개별공시지가(㎡,원)",
                               "매매가(만원)", "소유자", "메모", "비고"]
                    for _, d_row in rows.iterrows():
                        r_idx = self.table_land.rowCount()
                        self.table_land.insertRow(r_idx)
                        for c_idx, h in enumerate(headers):
                            val = str(d_row.get(h, ""))
                            if val == "nan":
                                val = ""
                            self.table_land.setItem(r_idx, c_idx, QTableWidgetItem(val))

            if found:
                self.output.append(f"✅ {m_id} 데이터를 성공적으로 불러왔습니다.")
            else:
                QMessageBox.warning(self, "실패", f"ID '{m_id}'에 해당하는 매물을 찾을 수 없습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"조회 중 오류 발생: {str(e)}")

    def process_save(self):
        type_name = "토지"
        pfx = "LND"
        m_id = self.le_m_id.text().strip().upper()

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

        media_path_info = "-"

        record = {"매물ID": m_id}
        is_public = self.chk_loc_public.isChecked()
        record.update({
            "특별/특/광/도": self.le_do.text(),
            "시/군/구": self.le_si_gun.text(),
            "읍/면/동/리": self.le_emd_ri.text(),
        })
        if is_public:
            record["지번"] = self.le_jibon.text()

        for group_key in ["G2", "G3"]:
            for label, widget in self.all_fields[group_key].items():
                if not isinstance(widget, QPushButton):
                    record[label] = widget.currentText() if isinstance(widget, QComboBox) else widget.text()

        record["미디어경로(사진/영상)"] = media_path_info
        if not is_public:
            record["지번"] = self.le_jibon.text()

        for label, widget in self.all_fields["G4"].items():
            record[label] = widget.currentText() if isinstance(widget, QComboBox) else widget.text()

        # Extract land detail
        detail_df_land = None
        if self.table_land.rowCount() > 0:
            land_details = []
            l_headers = ["필지ID", "특별/특/광/도", "시/군/구", "읍/면/동/리", "지번",
                         "용도지역", "지목", "면적(㎡)", "개별공시지가(㎡,원)",
                         "매매가(만원)", "소유자", "메모", "비고"]
            for r in range(self.table_land.rowCount()):
                sub_id = f"{m_id}-{r+1:02d}"
                row_data = {"매물ID": m_id}
                for c in range(self.table_land.columnCount()):
                    item = self.table_land.item(r, c)
                    val = sub_id if c == 0 else (item.text() if item else "")
                    row_data[l_headers[c]] = val
                land_details.append(row_data)
            detail_df_land = pd.DataFrame(land_details)

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

                    if detail_df_land is not None and not detail_df_land.empty:
                        l_sheet = "토지상세"
                        if l_sheet in sheets:
                            sheets[l_sheet] = sheets[l_sheet][sheets[l_sheet]['매물ID'] != m_id]
                            sheets[l_sheet] = pd.concat([sheets[l_sheet], detail_df_land], ignore_index=True)
                        else:
                            sheets[l_sheet] = detail_df_land
                else:
                    sheets = {type_name: new_df}
                    if detail_df_land is not None:
                        sheets["토지상세"] = detail_df_land

                with pd.ExcelWriter(self.db_path, engine='openpyxl') as writer:
                    for s_name, df in sheets.items():
                        df.to_excel(writer, sheet_name=s_name, index=False)
                        ws = writer.sheets[s_name]
                        apply_excel_styles(ws, df, numeric_keywords=["면적", "개별공시지가", "매매가", "조정가"])
            except PermissionError:
                raise Exception(
                    f"엑셀 파일('{self.db_path}')이 이미 열려 있습니다.\n엑셀 프로그램을 닫고 다시 시도해 주세요."
                )

            loc_info = f"{record['시/군/구']} {record['읍/면/동/리']} {record.get('지번', '')}"
            self.output.append(f"✅ {m_id} 저장 성공: {loc_info}")
            QMessageBox.information(self, "성공", f"{m_id} 저장 및 분석이 완료되었습니다.")
            self.clear_inputs()
        except Exception as e:
            self.output.append(f"❌ 에러 발생: {str(e)}")
            QMessageBox.critical(self, "저장 에러", f"데이터를 저장하지 못했습니다.\n원인: {str(e)}")

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
        self.table_land.setRowCount(0)
        self.calculate_land_totals()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = LandEstateWidget()
    window.show()
    sys.exit(app.exec())
