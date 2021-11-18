from openpyxl import load_workbook
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
from urllib.parse import unquote_plus
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from functions import login
from functions import openChrome

# selenium 실행 및 driver 정보 가져오기
driver = openChrome.openChrome()

# 엑셀파일 불러오기
load_wb = load_workbook("./wineAutoBid.xlsx", data_only=True)

# Sheet 라는 이름의 시트 선택
ws2 = load_wb['Sheet2']
