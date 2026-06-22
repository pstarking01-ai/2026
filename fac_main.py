import sys

from PyQt6.QtWidgets import QApplication

from shared.base_estate_widget import BaseEstateWidget


class FactoryEstateWidget(BaseEstateWidget):
    """Factory (공장) management widget."""

    property_name = "공장"
    prefix = "FCT"

    category_config = {
        "공장": {
            "G2": ["사용승인일", "대지면적(㎡)", "건축면적(㎡)", "연면적(㎡)",
                    "용도지역/지구/구역", "지목", "주구조", "주용도", "건물높이", "층수"],
            "G3": ["거래가", "보증금", "월세", "화장실", "승강기", "주차장(대)",
                    "호이스트", "도로조건", "하수처리", "전력(kw/h)"],
            "MKT": ["광고상황", "광고사이트", "자체광고 채널", "사진,영상링크"],
        }
    }

    detail_config = {
        "title": "공장/창고 상세 (층별 현황/단위:㎡/만원)",
        "headers": ["층", "구조", "용도", "면적(㎡)", "보증금(만원)", "월세(만원)", "비고"],
        "numeric_cols": [3, 4, 5],
        "totals_label": "합계: 보증금 0만원 / 월세 0만원 (수익률: 0.00%)",
        "totals_color": "#8e44ad",
        "add_btn_text": "+ 층 추가",
        "del_btn_text": "- 층 삭제",
    }

    g4_labels = ["소유자", "성별", "소유자연락처", "통신사", "연락처구분", "연락처",
                 "비번", "조정가", "중개사메모1", "중개사메모2"]

    def _update_detail_visibility(self, type_name):
        self.detail_group.setVisible(type_name in ["창고", "공장"])

    def _connect_totals_triggers(self):
        if "거래가" in self.all_fields["G3"]:
            self.all_fields["G3"]["거래가"].textChanged.connect(self.calculate_detail_totals)
        if "희망매매가(만원)" in self.all_fields["G3"]:
            self.all_fields["G3"]["희망매매가(만원)"].textChanged.connect(self.calculate_detail_totals)

    def calculate_detail_totals(self):
        total_dep = 0
        total_mon = 0

        price_field = self.all_fields.get("G3", {}).get("매매가")
        if not price_field:
            price_field = self.all_fields.get("G3", {}).get("거래가")
        price_text = price_field.text().replace(",", "") if price_field else ""
        price_val = int(price_text) if price_text.isdigit() else 0

        for r in range(self.detail_table.rowCount()):
            try:
                dep_item = self.detail_table.item(r, 4)
                mon_item = self.detail_table.item(r, 5)
                if dep_item and dep_item.text().replace(',', '').strip().isdigit():
                    total_dep += int(dep_item.text().replace(',', '').strip())
                if mon_item and mon_item.text().replace(',', '').strip().isdigit():
                    total_mon += int(mon_item.text().replace(',', '').strip())
            except Exception:
                continue

        total_yearly = total_mon * 12
        yield_rate = 0
        investment = price_val - total_dep
        if investment > 0:
            yield_rate = (total_yearly / investment) * 100

        self.lbl_detail_totals.setText(
            f"합계: 보증금 {total_dep:,}만원 / 월세 {total_mon:,}만원 "
            f"(연 {total_yearly:,}만원) 수익률: {yield_rate:.2f}%"
        )


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FactoryEstateWidget()
    window.show()
    sys.exit(app.exec())
