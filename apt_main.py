import sys
import pandas as pd
import os
from datetime import datetime
import uuid
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QComboBox, QDateEdit, QTableWidget,
                             QTableWidgetItem, QPushButton, QScrollArea, QFormLayout, QMessageBox, QHeaderView)
from PyQt6.QtCore import QDate, Qt
from openpyxl.styles import Font, Alignment

from shared.config import Config

class ApartmentApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("아파트 매물 등록 시스템 v1.0")
        self.setGeometry(100, 100, 900, 800)
        
        self.db_file = Config.DB_PATH
        self.info_file = self.db_file

        self.type_map = {"아파트": Config.PREFIX_MAP["아파트"]}
        
        self.init_ui()
        self.refresh_listing_id()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 검색 및 제어 바 (상단: 검색창 / 하단: 버튼)
        search_input_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("검색할 매물 ID 입력...")
        search_input_layout.addWidget(QLabel("매물 검색:"))
        search_input_layout.addWidget(self.search_input)
        main_layout.addLayout(search_input_layout)

        search_btn_layout = QHBoxLayout()
        search_btn = QPushButton("데이터 불러오기(수정)")
        search_btn.clicked.connect(self.load_data)
        reset_btn = QPushButton("신규 작성(초기화)")
        reset_btn.clicked.connect(self.reset_form)
        search_btn_layout.addWidget(search_btn)
        search_btn_layout.addWidget(reset_btn)
        main_layout.addLayout(search_btn_layout)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        self.layout = QVBoxLayout(content_widget)

        # 1. 공개 정보
        form1 = QFormLayout()

        self.prop_type_combo = QComboBox()
        self.prop_type_combo.addItems(list(self.type_map.keys()))
        self.prop_type_combo.currentTextChanged.connect(self.refresh_listing_id)

        self.listing_id = QLineEdit()
        self.listing_id.setReadOnly(True)
        self.listing_id.setStyleSheet("background-color: #f0f0f0; font-weight: bold;")

        self.complex_name = QLineEdit()
        self.dong = QLineEdit()
        self.ho = QLineEdit()
        self.living_type = QComboBox()
        self.living_type.addItems(["자가", "임대"])
        self.trade_type = QComboBox()
        self.trade_type.addItems(["매매", "전세", "월세"])
        self.trade_type.currentTextChanged.connect(self.toggle_price_fields)

        # 가격 입력 필드 (동적으로 표시/숨김)
        self.price_buy = QLineEdit(); self.price_buy.textChanged.connect(lambda: self.format_number(self.price_buy))
        self.price_jeonse = QLineEdit(); self.price_jeonse.textChanged.connect(lambda: self.format_number(self.price_jeonse))
        self.price_deposit = QLineEdit(); self.price_deposit.textChanged.connect(lambda: self.format_number(self.price_deposit))
        self.price_monthly = QLineEdit(); self.price_monthly.textChanged.connect(lambda: self.format_number(self.price_monthly))

        # 가격 필드 라벨
        self.price_buy_label = QLabel("거래가:")
        self.price_jeonse_label = QLabel("거래가:")
        self.price_deposit_label = QLabel("거래가(보증금):")
        self.price_monthly_label = QLabel("월세:")

        form1.addRow("부동산 종류:", self.prop_type_combo)
        form1.addRow("매물 ID (자동):", self.listing_id)
        form1.addRow("단지명:", self.complex_name)
        form1.addRow("동:", self.dong)
        form1.addRow("호:", self.ho)
        form1.addRow("거주 형태:", self.living_type)
        form1.addRow("거래 유형:", self.trade_type)
        form1.addRow(self.price_buy_label, self.price_buy)
        form1.addRow(self.price_jeonse_label, self.price_jeonse)
        form1.addRow(self.price_deposit_label, self.price_deposit)
        form1.addRow(self.price_monthly_label, self.price_monthly)

        # 2. 비공개 정보
        form2 = QFormLayout()
        self.owner_name = QLineEdit()
        self.owner_gender = QComboBox()
        self.owner_gender.addItems(["남", "여", "미상"])
        self.owner_contact = QLineEdit()
        self.owner_contact.textChanged.connect(lambda: self.auto_hyphen(self.owner_contact))
        self.owner_carrier = QComboBox()
        self.owner_carrier.addItems(["SKT", "KT", "LGU+", "SKT알뜰폰", "KT알뜰폰", "LGU+알뜰폰", "미확인"])
        self.owner_phone = QLineEdit() # 관계인 전화번호
        self.owner_phone.textChanged.connect(lambda: self.auto_hyphen(self.owner_phone))
        self.owner_password = QLineEdit() # 비밀번호 (현관, 집)
        self.negotiable_price = QLineEdit(); self.negotiable_price.textChanged.connect(lambda: self.format_number(self.negotiable_price))
        self.lease_end_date = QLineEdit()
        self.lease_end_date.setPlaceholderText("예: 20241231")
        self.lease_end_date.setMaxLength(10)
        self.lease_end_date.textChanged.connect(lambda: self.auto_hyphen_date(self.lease_end_date))
        self.memo1 = QLineEdit(); self.memo1.setMaxLength(100) # 메모1 길이 확장
        self.memo2 = QLineEdit(); self.memo2.setMaxLength(100) # 메모2 추가
        self.relationship_type = QComboBox()
        self.relationship_type.addItems(["의뢰인", "임차인", "기타"])

        form2.addRow("소유자명:", self.owner_name)
        form2.addRow("성별:", self.owner_gender)
        form2.addRow("소유자 연락처1:", self.owner_contact)
        form2.addRow("통신사:", self.owner_carrier)
        form2.addRow("관계인:", self.relationship_type)
        form2.addRow("관계인 연락처:", self.owner_phone)
        form2.addRow("비밀번호(현관,집):", self.owner_password)
        form2.addRow("조정가:", self.negotiable_price)
        self.lease_end_label = QLabel("만기일:")
        form2.addRow(self.lease_end_label, self.lease_end_date)
        form2.addRow("메모 1(100자):", self.memo1)
        form2.addRow("메모 2(100자):", self.memo2)

        # 공개 정보와 비공개 정보를 2열로 배치
        top_h_layout = QHBoxLayout()
        public_info_v_layout = QVBoxLayout()
        public_info_v_layout.addWidget(QLabel("<h3>1. 공개 정보</h3>"))
        public_info_v_layout.addLayout(form1)
        private_info_v_layout = QVBoxLayout()
        private_info_v_layout.addWidget(QLabel("<h3>2. 비공개 정보</h3>"))
        private_info_v_layout.addLayout(form2)
        top_h_layout.addLayout(public_info_v_layout)
        top_h_layout.addLayout(private_info_v_layout)
        self.layout.addLayout(top_h_layout)

        # 3. 상세 정보
        self.layout.addWidget(QLabel("<h3>3. 아파트 상세 정보</h3>"))

        detail_h_layout = QHBoxLayout()
        form3_left = QFormLayout()
        form3_right = QFormLayout()

        self.total_households = QLineEdit()
        self.approval_date = QLineEdit()  # 사용승인일 추가
        self.total_parking = QLineEdit()
        self.management_office_contact = QLineEdit()
        self.management_office_contact.textChanged.connect(lambda: self.auto_hyphen_334(self.management_office_contact))

        self.parking_per_unit = QLineEdit(); self.parking_per_unit.setReadOnly(True)
        self.total_households.textChanged.connect(self.calculate_parking)
        self.total_parking.textChanged.connect(self.calculate_parking)

        form3_left.addRow("세대수:", self.total_households)
        form3_right.addRow("사용승인일:", self.approval_date)
        form3_left.addRow("총 주차대수:", self.total_parking)
        form3_right.addRow("세대당 주차 (자동):", self.parking_per_unit)
        form3_left.addRow("관리사무소 연락처:", self.management_office_contact)

        detail_h_layout.addLayout(form3_left)
        detail_h_layout.addLayout(form3_right)
        self.layout.addLayout(detail_h_layout)

        # 주소 정보 (1열 유지)
        form_address = QFormLayout()
        self.land_address = QLineEdit()
        self.road_address = QLineEdit()
        form_address.addRow("지번주소:", self.land_address)
        form_address.addRow("도로명주소:", self.road_address)
        self.layout.addLayout(form_address)

        # 4. 동 정보 테이블
        self.layout.addWidget(QLabel("<h3>4. 동 정보</h3>"))
        headers = ["동", "라인", "층", "평형", "타입", "공급면적", "전용면적", "방수", "욕실수", "방향", "기타1", "기타2"]
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.layout.addWidget(self.table)
        
        btn_layout = QHBoxLayout()
        add_row_btn = QPushButton("동 정보 행 추가")
        add_row_btn.clicked.connect(lambda: self.table.insertRow(self.table.rowCount()))
        del_row_btn = QPushButton("선택 행 삭제")
        del_row_btn.clicked.connect(lambda: self.table.removeRow(self.table.currentRow()))
        copy_md_btn = QPushButton("매물 정보 마크다운 복사")
        copy_md_btn.clicked.connect(self.copy_as_markdown)
        btn_layout.addWidget(add_row_btn)
        btn_layout.addWidget(del_row_btn)
        btn_layout.addWidget(copy_md_btn)
        self.layout.addLayout(btn_layout)

        # 저장 버튼
        save_btn = QPushButton("매물 등록 / 수정사항 저장")
        save_btn.setFixedHeight(50)
        save_btn.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold;")
        save_btn.clicked.connect(self.save_to_excel)
        main_layout.addWidget(scroll)
        # 초기 가격 필드 상태 설정
        self.toggle_price_fields(self.trade_type.currentText())
        main_layout.addWidget(save_btn)
        scroll.setWidget(content_widget)

    def reset_form(self):
        """모든 입력 필드를 초기화하고 새로운 매물 ID를 생성합니다."""
        self.complex_name.clear()
        self.dong.clear()
        self.ho.clear()
        self.price_buy.clear()
        self.price_jeonse.clear()
        self.price_deposit.clear()
        self.price_monthly.clear()
        self.owner_name.clear()
        self.owner_contact.clear()
        self.owner_phone.clear()
        self.owner_password.clear()
        self.negotiable_price.clear()
        self.lease_end_date.clear()
        self.memo1.clear()
        self.memo2.clear()
        self.total_households.clear()
        self.approval_date.clear()
        self.total_parking.clear()
        self.management_office_contact.clear()
        self.land_address.clear()
        self.road_address.clear()
        self.table.setRowCount(0)
        self.refresh_listing_id()
        QMessageBox.information(self, "알림", "입력창이 초기화되었습니다. 신규 등록이 가능합니다.")

    def load_data(self):
        """입력된 ID를 바탕으로 엑셀에서 데이터를 찾아 화면에 표시합니다."""
        target_id = self.search_input.text().strip()
        if not target_id:
            QMessageBox.warning(self, "경고", "조회할 매물 ID를 입력해주세요.")
            return

        try:
            if os.path.exists(self.db_file):
                df = pd.read_excel(self.db_file, sheet_name="아파트")
                res = df[df["매물ID"].astype(str) == target_id]
                
                if not res.empty:
                    row = res.iloc[0]
                    # 데이터 채우기 (공백 제거 처리 추가)
                    complex_name_val = str(row.get("단지명", "")).strip()
                    self.listing_id.setText(str(row.get("매물ID", "")))
                    self.complex_name.setText(complex_name_val)
                    self.dong.setText(str(row.get("동", "")))
                    self.ho.setText(str(row.get("호", "")))
                    self.living_type.setCurrentText(str(row.get("거주형태", "")))
                    self.trade_type.setCurrentText(str(row.get("거래유형", "")))
                    
                    # 거래가 필드를 거래유형에 맞춰 복원 (통합 컬럼 '거래가' 대응)
                    trade_type_val = str(row.get("거래유형", ""))
                    deal_price_val = str(row.get("거래가", "")).replace("nan", "")
                    if trade_type_val == "매매": self.price_buy.setText(deal_price_val)
                    elif trade_type_val == "전세": self.price_jeonse.setText(deal_price_val)
                    elif trade_type_val == "월세": self.price_deposit.setText(deal_price_val)

                    self.price_monthly.setText(str(row.get("월세", "")))
                    self.owner_name.setText(str(row.get("소유자", "")))
                    self.owner_gender.setCurrentText(str(row.get("성별", "미상")))
                    self.owner_contact.setText(str(row.get("소유자연락처", "")))
                    self.owner_carrier.setCurrentText(str(row.get("통신사", "미확인")))
                    self.relationship_type.setCurrentText(str(row.get("연락처구분", "의뢰인")))
                    self.owner_phone.setText(str(row.get("연락처", "")))
                    self.owner_password.setText(str(row.get("비번(현관/집)", "")))
                    self.negotiable_price.setText(str(row.get("조정가", "")))
                    
                    lease_val = str(row.get("만기일", ""))[:10]
                    if lease_val and lease_val != "nan":
                        self.lease_end_date.setText(lease_val)
                    self.memo1.setText(str(row.get("중개사메모1", "")))
                    self.memo2.setText(str(row.get("중개사메모2", "")))

                    # 상세 정보 로드 (apt_complex_info.xlsx)
                    if os.path.exists(self.db_file):
                        # 모든 시트를 읽어 시트명 매칭을 유연하게 처리
                        all_info_sheets = pd.read_excel(self.db_file, sheet_name=None)

                        # 1. 아파트 상세 정보 (아파트상세 시트 검색)
                        detail_key = next((k for k in all_info_sheets.keys() if k.strip() == "아파트상세"), None)

                        if detail_key:
                            detail_df = all_info_sheets[detail_key]
                            # 1. 아파트 상세 정보 로드 (단지ID 우선, 단지명 차선)
                            safe_complex_id = complex_name_val.replace(' ', '')
                            detail_res = pd.DataFrame()
                            
                            for id_col in ["단지ID", "id"]:
                                if id_col in detail_df.columns:
                                    detail_res = detail_df[detail_df[id_col].astype(str).str.strip() == safe_complex_id]
                                    if not detail_res.empty: break
                            
                            if detail_res.empty and "단지명" in detail_df.columns:
                                detail_res = detail_df[detail_df["단지명"].astype(str).str.strip() == complex_name_val]

                            if not detail_res.empty:
                                d_row = detail_res.iloc[0]
                                # nan 값 방지 처리 함수
                                # nan 값 방지 처리 함수 (단지ID는 UI에 표시되지 않으므로 제외)
                                def get_v(v): return "" if pd.isna(v) or str(v).lower() == "nan" else str(v)
                                
                                self.total_households.setText(get_v(d_row.get("세대수", "")))
                                self.approval_date.setText(get_v(d_row.get("사용승인일", "")))
                                self.total_parking.setText(get_v(d_row.get("총주차대수", "")))
                                self.management_office_contact.setText(get_v(d_row.get("관리사무소연락처", "")))
                                self.land_address.setText(get_v(d_row.get("지번주소", "")))
                                self.road_address.setText(get_v(d_row.get("도로명주소", "")))

                        # 2. 동 정보 테이블 로드 (시트명: [단지명]_dong 검색 - 공백 제거 매칭으로 정확도 향상)
                        target_dong_name = f"{complex_name_val}_dong".replace(" ", "")
                        dong_key = next((k for k in all_info_sheets.keys() if k.strip().replace(" ", "") == target_dong_name), None)
                        
                        if dong_key:
                            dong_df = all_info_sheets[dong_key]
                            self.table.setRowCount(len(dong_df))
                            for r_idx, r_data in dong_df.iterrows():
                                for c_idx in range(self.table.columnCount()):
                                    header = self.table.horizontalHeaderItem(c_idx).text()
                                    val = r_data.get(header, "")
                                    self.table.setItem(r_idx, c_idx, QTableWidgetItem("" if pd.isna(val) else str(val)))
                    
                    QMessageBox.information(self, "성공", f"매물[{target_id}] 데이터를 불러왔습니다. 수정 후 저장하세요.")
                else:
                    QMessageBox.warning(self, "실패", "해당 ID의 매물을 찾을 수 없습니다.")
            else:
                QMessageBox.warning(self, "오류", "데이터베이스 파일이 존재하지 않습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"데이터 로드 중 오류 발생: {e}")

    def refresh_listing_id(self):
        """요청하신 규칙에 따라 매물 ID를 생성합니다: aptYYMMDD-NN"""
        prop_name = self.prop_type_combo.currentText()
        type_code = self.type_map.get(prop_name, "etc")
        date_str = datetime.now().strftime("%y%m%d") # 예: 260611
        seq = 1
        if os.path.exists(self.db_file):
            try:
                # 사용자 요청에 따라 '아파트' 시트에서 데이터를 관리합니다.
                df = pd.read_excel(self.db_file, sheet_name="아파트")
                if "매물ID" in df.columns:
                    prefix = f"{type_code}{date_str}-"
                    # 해당 접두어로 시작하는 ID들만 필터링
                    relevant_ids = df[df["매물ID"].astype(str).str.startswith(prefix)]["매물ID"].tolist()
                    if relevant_ids:
                        seq_list = []
                        for rid in relevant_ids:
                            try:
                                # 하이픈 뒤의 숫자 추출
                                seq_list.append(int(rid.split('-')[-1]))
                            except: continue
                        if seq_list:
                            seq = max(seq_list) + 1
            except Exception:
                pass # 파일이 없거나 시트가 없는 경우 순번 1 유지
        
        self.listing_id.setText(f"{type_code}{date_str}-{str(seq).zfill(2)}")

    def toggle_price_fields(self, text):
        # 모든 가격 필드와 라벨을 숨김
        self.price_buy_label.hide()
        self.price_buy.hide()
        self.price_jeonse_label.hide()
        self.price_jeonse.hide()
        self.price_deposit_label.hide()
        self.price_deposit.hide()
        self.price_monthly_label.hide()
        self.price_monthly.hide()

        # 선택된 거래 유형에 따라 필드 표시
        if text == "매매":
            self.price_buy_label.show()
            self.price_buy.show()
        elif text == "전세":
            self.price_jeonse_label.show()
            self.price_jeonse.show()
        elif text == "월세":
            self.price_deposit_label.show()
            self.price_deposit.show()
            self.price_monthly_label.show()
            self.price_monthly.show()

    def auto_hyphen(self, widget):
        text = widget.text().replace("-", "")
        formatted = text
        if len(text) > 3 and len(text) <= 7:
            formatted = f"{text[:3]}-{text[3:]}"
        elif len(text) > 7:
            formatted = f"{text[:3]}-{text[3:7]}-{text[7:11]}"
        
        if widget.text() != formatted:
            pos = widget.cursorPosition()
            old_len = len(widget.text())
            widget.setText(formatted)
            new_pos = pos + (len(formatted) - old_len)
            widget.setCursorPosition(max(0, new_pos))

    def auto_hyphen_334(self, widget):
        """3-3-4 자릿수 기준 하이픈 자동 삽입"""
        text = widget.text().replace("-", "")
        formatted = text
        if len(text) > 3 and len(text) <= 6:
            formatted = f"{text[:3]}-{text[3:]}"
        elif len(text) > 6:
            formatted = f"{text[:3]}-{text[3:6]}-{text[6:10]}"
            
        if widget.text() != formatted:
            pos = widget.cursorPosition()
            old_len = len(widget.text())
            widget.setText(formatted)
            new_pos = pos + (len(formatted) - old_len)
            widget.setCursorPosition(max(0, new_pos))

    def auto_hyphen_date(self, widget):
        """날짜 입력 시 8자리 숫자를 YYYY-MM-DD 형식으로 변환"""
        text = "".join(filter(str.isdigit, widget.text()))
        if len(text) > 8: text = text[:8]
        formatted = text
        if len(text) > 4 and len(text) <= 6:
            formatted = f"{text[:4]}-{text[4:]}"
        elif len(text) > 6:
            formatted = f"{text[:4]}-{text[4:6]}-{text[6:8]}"
            
        if widget.text() != formatted:
            pos = widget.cursorPosition()
            old_len = len(widget.text())
            widget.setText(formatted)
            new_pos = pos + (len(formatted) - old_len)
            widget.setCursorPosition(max(0, new_pos))

    def format_number(self, widget):
        """숫자 입력 시 천 단위 콤마 표시"""
        text = widget.text().replace(",", "")
        if text.isdigit():
            formatted = f"{int(text):,}"
            if widget.text() != formatted:
                pos = widget.cursorPosition()
                old_len = len(widget.text())
                widget.setText(formatted)
                # 콤마 추가로 인한 커서 위치 조정
                new_pos = pos + (len(formatted) - old_len)
                widget.setCursorPosition(max(0, new_pos))
        elif not text:
            widget.setText("")

    def keyPressEvent(self, event):
        """테이블 복사/붙여넣기 단축키 처리"""
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_V:
                self.paste_data()
            elif event.key() == Qt.Key.Key_C:
                self.copy_data()
        super().keyPressEvent(event)

    def copy_data(self):
        """테이블 선택 영역 복사"""
        selection = self.table.selectedRanges()
        if not selection: return
        res = ""
        for r in range(selection[0].topRow(), selection[0].bottomRow() + 1):
            for c in range(selection[0].leftColumn(), selection[0].rightColumn() + 1):
                item = self.table.item(r, c)
                res += (item.text() if item else "") + "\t"
            res = res.rstrip("\t") + "\n"
        QApplication.clipboard().setText(res)

    def paste_data(self):
        """클립보드 데이터를 파싱하여 아파트 상세 필드 또는 동 정보 테이블에 자동 입력"""
        clipboard = QApplication.clipboard()
        text = clipboard.text().strip()
        if not text: return

        # 아파트 상세 정보 필드 매핑용
        detail_field_map = {
            "단지명": self.complex_name, "세대수": self.total_households,
            "사용승인일": self.approval_date, "총주차대수": self.total_parking,
            "관리사무소연락처": self.management_office_contact,
            "지번주소": self.land_address, "도로명주소": self.road_address
        }

        rows = text.strip().split('\n')
        curr_row = self.table.currentRow()
        curr_col = self.table.currentColumn()
        if curr_row == -1: curr_row = 0
        if curr_col == -1: curr_col = 0

        filled_detail = False
        for r_idx, row_text in enumerate(rows):
            # 1. 마크다운 상세 정보 파싱 (| 항목 | 내용 | 형태)
            if '|' in row_text:
                cols = [c.strip() for c in row_text.split('|') if c.strip()]
                
                # 상세 정보 필드 매칭 (예: | 세대수 | 1000 |)
                if len(cols) >= 2 and cols[0] in detail_field_map:
                    detail_field_map[cols[0]].setText(cols[1])
                    filled_detail = True
                    continue # 상세 필드로 처리했으면 다음 줄로

                # 마크다운 테이블 구분선(|---|) 건너뛰기
                if all(all(char in '-: ' for char in cell) for cell in cols):
                    continue
                
                # 테이블 데이터로 간주하고 진행
            else:
                cols = row_text.split('\t')

            row_to_fill = curr_row + r_idx
            if row_to_fill >= self.table.rowCount():
                self.table.insertRow(self.table.rowCount())
                
            for c_idx, col_text in enumerate(cols):
                col_to_fill = curr_col + c_idx
                if col_to_fill < self.table.columnCount():
                    self.table.setItem(row_to_fill, col_to_fill, QTableWidgetItem(col_text.strip()))

        if filled_detail: self.calculate_parking()
        self.table.resizeColumnsToContents()

    def copy_as_markdown(self):
        """현재 입력된 정보를 마크다운 테이블 형식으로 복사하여 외부에서 수정 가능하도록 함"""
        try:
            md = f"### 매물 상세 보고서: {self.complex_name.text()} {self.dong.text()}동 {self.ho.text()}호\n\n"
            md += "| 항목 | 내용 |\n| :--- | :--- |\n"
            md += f"| 매물ID | {self.listing_id.text()} |\n"
            md += f"| 거래유형 | {self.trade_type.currentText()} |\n"
            md += f"| 거래가 | {self.price_buy.text() or self.price_jeonse.text() or self.price_deposit.text() + (f'/{self.price_monthly.text()}' if self.price_monthly.text() else '')} |\n"
            md += f"| 소유자 | {self.owner_name.text()} ({self.owner_gender.currentText()}) |\n"
            md += f"| 연락처 | {self.owner_contact.text()} |\n"
            md += f"| 조정가 | {self.negotiable_price.text()} |\n"
            md += f"| 만기일 | {self.lease_end_date.text()} |\n"
            md += f"| 메모1 | {self.memo1.text()} |\n"
            md += f"| 메모2 | {self.memo2.text()} |\n\n"
            
            md += "#### 단지 및 주소 정보\n"
            md += f"- 주소: {self.road_address.text()} ({self.land_address.text()})\n"
            md += f"- 세대수/주차: {self.total_households.text()}세대 / 총 {self.total_parking.text()}대 (세대당 {self.parking_per_unit.text()})\n"
            
            QApplication.clipboard().setText(md)
            QMessageBox.information(self, "복사 완료", "매물 정보가 마크다운 형식으로 클립보드에 복사되었습니다.\n메모장이나 메신저에 붙여넣으세요.")
        except Exception as e:
            QMessageBox.warning(self, "오류", f"마크다운 생성 중 오류 발생: {e}")

    def calculate_parking(self):
        try:
            h = float(self.total_households.text())
            p = float(self.total_parking.text())
            self.parking_per_unit.setText(f"{p/h:.2f}")
        except: self.parking_per_unit.clear()

    def save_sheet(self, file_path, sheet_name, data_dict_list):
        """모든 시트를 유지하며 특정 시트의 데이터를 업데이트하는 공통 함수"""
        # 저장될 폴더가 없으면 자동으로 생성합니다.
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        new_df = pd.DataFrame(data_dict_list)
        
        if os.path.exists(file_path):
            all_sheets = pd.read_excel(file_path, sheet_name=None)
            
            if sheet_name in all_sheets:
                existing_df = all_sheets[sheet_name]

                # 업데이트를 위한 키 컬럼 결정 (단지ID > id > 매물ID > 단지명)
                key_col = None
                for col in ["단지ID", "id", "매물ID", "단지명"]:
                    if col in new_df.columns and col in existing_df.columns:
                        key_col = col
                        break

                if key_col:
                    df = pd.concat([existing_df[~existing_df[key_col].isin(new_df[key_col])], new_df], ignore_index=True)
                else: # 키 컬럼이 없거나 지정되지 않았다면 단순히 추가
                    df = pd.concat([existing_df, new_df], ignore_index=True)
            else:
                df = new_df
            
            # 헤더 순서 배치: 기존 시트의 헤더 순서를 유지하되, 새로운 컬럼이 있으면 뒤에 추가
            if not df.empty:
                # 입력 데이터(dict)의 키 순서를 기준으로 하되, 기존 파일이 있다면 기존 헤더 순서 우선
                current_cols = df.columns.tolist()
                input_cols = list(data_dict_list[0].keys())
                
                # 헤더 정렬 (입력 데이터의 키 순서로 재배치)
                ordered_cols = [c for c in input_cols if c in current_cols] + \
                               [c for c in current_cols if c not in input_cols]
                df = df[ordered_cols]

            all_sheets[sheet_name] = df

            try:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    for s_name, s_df in all_sheets.items():
                        s_df.to_excel(writer, sheet_name=s_name, index=False)
                        ws = writer.sheets[s_name]
                        self._apply_excel_styles(ws, s_df)
            except PermissionError:
                raise Exception(f"파일({os.path.basename(file_path)})이 이미 열려 있습니다. 엑셀을 닫고 다시 저장하세요.")
        else:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df = pd.DataFrame(data_dict_list)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                self._apply_excel_styles(ws, df)

    def _apply_excel_styles(self, ws, df):
        """엑셀 워크시트에 폰트 16, 행높이 30, 좌측/중앙 정렬 및 자동 열너비 적용"""
        font_normal = Font(name='맑은 고딕', size=16)
        font_bold = Font(name='맑은 고딕', size=16, bold=True)
        
        for r_idx, row in enumerate(ws.iter_rows(), 1):
            ws.row_dimensions[r_idx].height = 30
            is_header = (r_idx == 1)
            for cell in row:
                cell.font = font_bold if is_header else font_normal
                cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # 자동 열 너비 조정
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    current_len = sum(2.5 if ord(c) > 128 else 1.2 for c in str(cell.value))
                    if current_len > max_len: max_len = current_len
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    def save_to_excel(self):
        try:
            c_name = self.complex_name.text().strip()
            if not c_name:
                QMessageBox.warning(self, "경고", "단지명을 입력해야 저장이 가능합니다.")
                return

            # 동 정보 테이블에서 현재 입력된 동과 일치하는 상세 정보(평형, 타입 등) 자동 수집
            # 사용자 요청: estate_db.xlsx의 해당 단지_dong 시트에서 정보 가져오기
            pyung_val, type_val, supp_area, priv_area = "", "", "", ""
            target_dong = self.dong.text().strip()

            if os.path.exists(self.db_file):
                try:
                    all_info_sheets = pd.read_excel(self.db_file, sheet_name=None)
                    # 단지명으로 동 정보 시트 이름 구성 (공백 제거)
                    dong_sheet_name_candidate = f"{c_name.replace(' ', '')}_dong"
                    dong_key = next((k for k in all_info_sheets.keys() if k.strip().replace(" ", "") == dong_sheet_name_candidate), None)

                    if dong_key:
                        dong_df = all_info_sheets[dong_key]
                        # "동" 컬럼에서 현재 입력된 동과 일치하는 행 찾기
                        if "동" in dong_df.columns:
                            matching_dong_rows = dong_df[dong_df["동"].astype(str).str.strip() == target_dong]
                        else:
                            matching_dong_rows = pd.DataFrame()
                        if not matching_dong_rows.empty:
                            # 첫 번째 일치하는 행에서 정보 추출
                            dong_row = matching_dong_rows.iloc[0]
                            pyung_val = str(dong_row.get("평형", "")) if not pd.isna(dong_row.get("평형")) else ""
                            type_val = str(dong_row.get("타입", "")) if not pd.isna(dong_row.get("타입")) else ""
                            supp_area = str(dong_row.get("공급면적", "")) if not pd.isna(dong_row.get("공급면적")) else ""
                            priv_area = str(dong_row.get("전용면적", "")) if not pd.isna(dong_row.get("전용면적")) else ""
                except Exception as e:
                    print(f"estate_db.xlsx에서 동 정보 로드 중 오류 발생: {e}")
                    # 오류 발생 시 기본값(빈 문자열) 유지

            trade_type_val = self.trade_type.currentText()
            # 거래유형에 따른 통합 거래가 결정
            deal_price = self.price_buy.text() or self.price_jeonse.text() or self.price_deposit.text()

            # 1. 공개/비공개 데이터 저장 (표준화된 헤더로 정리)
            pub_priv_data = {
                "매물ID": self.listing_id.text(),
                "등록일": datetime.now().strftime("%Y-%m-%d"),
                "단지명": c_name, 
                "동": self.dong.text(), 
                "호": self.ho.text(),
                "평형": pyung_val, 
                "타입": type_val,
                "공급면적": supp_area, 
                "전용면적": priv_area,
                "거주형태": self.living_type.currentText(), 
                "거래유형": trade_type_val,
                "거래가": deal_price,
                "월세": self.price_monthly.text(),
                "소유자": self.owner_name.text(), 
                "성별": self.owner_gender.currentText(),
                "소유자연락처": self.owner_contact.text(), 
                "통신사": self.owner_carrier.currentText(),
                "연락처구분": self.relationship_type.currentText(),
                "연락처": self.owner_phone.text(),
                "비번(현관/집)": self.owner_password.text(), 
                "조정가": self.negotiable_price.text(),
                "만기일": self.lease_end_date.text(), 
                "중개사메모1": self.memo1.text(), 
                "중개사메모2": self.memo2.text(),
            }
            self.save_sheet(self.db_file, "아파트", [pub_priv_data])

            # 2 & 3. 아파트 상세 정보 및 동 정보 중복 체크 및 저장
            safe_complex_id = c_name.replace(' ', '') # 단지ID로 사용할 값
            complex_exists = False
            if os.path.exists(self.db_file):
                try:
                    all_info = pd.read_excel(self.db_file, sheet_name=None)
                    detail_key = next((k for k in all_info.keys() if k.strip() == "아파트상세"), None)
                    if detail_key:
                        check_df = all_info[detail_key]
                        # '단지ID' 컬럼이 있다면 '단지ID'로 존재 여부 확인
                        id_col = "단지ID" if "단지ID" in check_df.columns else ("id" if "id" in check_df.columns else "단지명")
                        search_val = safe_complex_id if id_col in ["단지ID", "id"] else c_name
                        if search_val in check_df[id_col].astype(str).str.strip().values:
                            complex_exists = True
                except Exception:
                    pass

            save_details = False
            if complex_exists:
                reply = QMessageBox.question(self, "정보 수정 확인", 
                                           f"'{c_name}'의 상세 정보가 이미 존재합니다.\n입력된 내용으로 정보를 수정하시겠습니까?",
                                           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.Yes:
                    save_details = True
            else:
                save_details = True

            if save_details:
                # 2. 아파트 상세 정보 저장 (estate_db.xlsx -> '아파트상세' 시트)
                detail_data = {
                    "단지ID": safe_complex_id,
                    "단지명": c_name, 
                    "세대수": self.total_households.text(),
                    "사용승인일": self.approval_date.text(), "총주차대수": self.total_parking.text(),
                    "세대당주차": self.parking_per_unit.text(), "관리사무소연락처": self.management_office_contact.text(),
                    "지번주소": self.land_address.text(), "도로명주소": self.road_address.text(),
                }
                self.save_sheet(self.db_file, "아파트상세", [detail_data])

                # 3. 동 정보 저장 (estate_db.xlsx -> '[단지명]_dong' 시트)
                table_rows = []
                for r in range(self.table.rowCount()):
                    row_data = {}
                    for c in range(self.table.columnCount()):
                        header = self.table.horizontalHeaderItem(c).text()
                        item = self.table.item(r, c)
                        row_data[header] = item.text() if item else ""
                    table_rows.append(row_data)
                
                if table_rows:
                    # 저장 시에도 공백을 제거한 ID 형태의 시트명 사용
                    dong_sheet_name = f"{c_name.replace(' ', '')}_dong"
                    self.save_sheet(self.db_file, dong_sheet_name, table_rows)

            # 저장 위치를 명확히 알 수 있도록 전체 경로를 표시합니다.
            QMessageBox.information(self, "저장 완료", f"모든 데이터가 '{self.db_file}'의 각 시트에 통합 정리되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"저장 실패: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ApartmentApp()
    window.show()
    sys.exit(app.exec())
    