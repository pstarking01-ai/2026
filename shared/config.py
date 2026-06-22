import os


class Config:
    DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "estate_db.xlsx")
    PHOTO_BASE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "property_photos")
    APT_MASTER_EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "apt_complex_info.xlsx")

    PREFIX_MAP = {
        "상가": "COM",
        "공장": "FCT",
        "주택": "HOU",
        "창고": "WHS",
        "토지": "LND",
        "아파트": "APT",
    }

    SNS_FOOTER = "성균관 공인중개사사무소\n성균관전자명함안내 https://pstarking01-ai.github.io/2026/"

    EXCEL_STYLE = {
        "font_size": 16,
        "row_height": 30,
        "public_color": "D9EAD3",
        "private_color": "FCE4EC",
    }
    WIDTH_KOR = 2.5
    WIDTH_ENG = 1.2
    MAX_WIDTH = 50

    DROPDOWN_DATA = {
        "거주형태": ["자가", "임대차"],
        "거래유형": ["매매", "전세", "월세"],
        "성별": ["남", "여"],
        "광고상황": ["광고전", "진행중", "거래완료", "매물철회"],
        "광고사이트": ["한방", "이실장", "부동산써브"],
        "자체광고 채널": ["홈페이지", "블로그", "유튜브"],
        "현재상태": ["거주", "공실"],
        "용도지역": [
            "제1종전용주거", "제2종전용주거", "제1종일반주거", "제2종일반주거",
            "제3종일반주거", "준주거", "중심상업", "일반상업", "근린상업", "유통상업",
            "전용공업지역", "일반공업지역", "준공업지역", "보전관리지역", "생산관리지역",
            "계획관리지역", "농림지역", "자연환경보전지역",
        ],
        "통신사": ["SKT", "KT", "LGU+", "SKT알뜰폰", "KT알뜰폰", "LGU+알뜰폰", "미확인"],
        "연락처구분": ["의뢰인", "임차인"],
    }
