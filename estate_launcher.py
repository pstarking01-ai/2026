"""
estate_launcher.py - 부동산 중개업무자동화시스템 통합 런처
모든 매물 종류(아파트, 상가, 공장, 주택, 창고, 토지) 접수 시스템을 
하나의 메뉴에서 선택하여 실행할 수 있는 통합 런처입니다.
"""
import sys
import os
import subprocess

from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QGroupBox, QGridLayout)
from PyQt6.QtCore import Qt


class EstateLauncher(QWidget):
    def __init__(self):
        super().__init__()
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.initUI()

    def initUI(self):
        self.setWindowTitle("부동산 중개업무자동화시스템 - 통합 런처")
        self.setGeometry(300, 200, 600, 400)
        self.setStyleSheet("""
            QWidget { background-color: #f8f9fa; font: 12pt 'Malgun Gothic'; }
            QPushButton { 
                font-size: 14pt; font-weight: bold; 
                padding: 20px; border-radius: 10px; 
                min-height: 60px;
            }
            QPushButton:hover { opacity: 0.8; }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        title = QLabel("부동산 중개업무자동화시스템")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 20pt; font-weight: bold; color: #2c3e50; padding: 10px;")
        layout.addWidget(title)

        subtitle = QLabel("성균관 공인중개사사무소")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("font-size: 12pt; color: #7f8c8d; margin-bottom: 20px;")
        layout.addWidget(subtitle)

        group = QGroupBox("매물 종류 선택")
        group.setStyleSheet("""
            QGroupBox { border: 2px solid #2c3e50; border-radius: 10px; margin-top: 15px; font-weight: bold; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
        """)
        grid = QGridLayout()
        grid.setSpacing(15)
        grid.setContentsMargins(20, 30, 20, 20)

        buttons = [
            ("apt_main.py", "아파트 (APT)", "#3498db"),
            ("com_main.py", "상가 (COM)", "#e67e22"),
            ("fac_main.py", "공장 (FCT)", "#9b59b6"),
            ("hou_main.py", "주택 (HOU)", "#27ae60"),
            ("whs_main.py", "창고 (WHS)", "#e74c3c"),
            ("lan_main.py", "토지 (LND)", "#16a085"),
        ]

        for i, (script, label, color) in enumerate(buttons):
            btn = QPushButton(label)
            btn.setStyleSheet(f"background-color: {color}; color: white;")
            btn.clicked.connect(lambda checked, s=script: self.launch_app(s))
            grid.addWidget(btn, i // 3, i % 3)

        group.setLayout(grid)
        layout.addWidget(group)
        self.setLayout(layout)

    def launch_app(self, script_name):
        script_path = os.path.join(self.base_dir, script_name)
        if not os.path.exists(script_path):
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "오류", f"{script_name} 파일을 찾을 수 없습니다.")
            return
        subprocess.Popen([sys.executable, script_path], cwd=self.base_dir)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = EstateLauncher()
    window.show()
    sys.exit(app.exec())
