import sys

from PyQt6.QtWidgets import QApplication

from shared.base_estate_widget import BaseEstateWidget


class HouseEstateWidget(BaseEstateWidget):
    """House (주택) management widget."""

    property_name = "주택"
    prefix = "HOU"

    category_config = {
        "주택": {
            "G2": ["준공연도", "대지면적(㎡)", "연면적(㎡)", "주구조", "거래유형", "거주형태"],
            "G3": ["거래가", "전세가", "보증금/월세"],
            "MKT": ["광고상황", "광고사이트", "자체광고 채널", "사진,영상링크"],
        }
    }

    detail_config = {
        "title": "주택접수 상세 (층별/호실별 현황/단위:㎡/만원)",
        "headers": ["층", "호실", "용도", "면적", "전세보증금", "월세보증금", "월세", "옵션", "현재상태", "임차만기일"],
        "numeric_cols": [3, 4, 5, 6],
        "totals_label": "합계: 보증금 0만원 / 월세 0만원 (연 0만원)",
        "totals_color": "#e67e22",
    }

    g4_labels = ["소유자", "성별", "소유자연락처", "통신사", "연락처구분", "연락처",
                 "비번(현관/집)", "조정가", "중개사메모1", "중개사메모2"]

    def _update_detail_visibility(self, type_name):
        self.detail_group.setVisible(type_name == "주택")

    def _connect_totals_triggers(self):
        if "거래가" in self.all_fields["G3"]:
            self.all_fields["G3"]["거래가"].textChanged.connect(self.calculate_detail_totals)
        if "희망매매가(만원)" in self.all_fields["G3"]:
            self.all_fields["G3"]["희망매매가(만원)"].textChanged.connect(self.calculate_detail_totals)

    def calculate_detail_totals(self):
        total_deposit = 0
        total_monthly = 0
        for r in range(self.detail_table.rowCount()):
            try:
                jeonse_item = self.detail_table.item(r, 4)
                w_dep_item = self.detail_table.item(r, 5)
                mon_item = self.detail_table.item(r, 6)
                for item in [jeonse_item, w_dep_item]:
                    if item and item.text().replace(',', '').strip().isdigit():
                        total_deposit += int(item.text().replace(',', '').strip())
                if mon_item and mon_item.text().replace(',', '').isdigit():
                    total_monthly += int(mon_item.text().replace(',', ''))
            except Exception:
                continue

        total_yearly = total_monthly * 12
        price_field = self.all_fields.get("G3", {}).get("매매가")
        if not price_field:
            price_field = self.all_fields.get("G3", {}).get("거래가")
        price_text = price_field.text().replace(",", "") if price_field else ""
        price_val = int(price_text) if price_text.isdigit() else 0
        yield_rate = 0
        investment = price_val - total_deposit
        if investment > 0:
            yield_rate = (total_yearly / investment) * 100
        self.lbl_detail_totals.setText(
            f"합계: 보증금 {total_deposit:,}만원 / 월세 {total_monthly:,}만원 "
            f"(연 {total_yearly:,}만원) 수익률: {yield_rate:.2f}%"
        )


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = HouseEstateWidget()
    window.show()
    sys.exit(app.exec())
