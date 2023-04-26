from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By

import requests
import datetime
import json
import time

################################################################################
# 설정
################################################################################
ID = "mskang"
PW = "aproele0320!@"
MENU_ID = '2002090000'      # 전자결재 > 결재문서 > 수신참조함

# 문서에 따라, 2, 3, 4, 5, 7 를 선택적으로 해야 함.
# 1: 본문, 2: 결재문서, 3:결재문서 + 결재의견, 4: 결재문서 + 첨부파일,
# 5: 결재문서 + 결재의견 + 첨부파일리스트, 6: 결재문서 + 추가항목
# 7: 결재문서 + 결재의견 + 첨부파일리스트 + 추가항목
PDF_TYPE = '5'
PAGE_SIZE = '1000'          # 한 번에 조회 할 데이터 수
READ_TYPE = ''              # 열람구분 ('': 전체, 10: 열람, 20: 미열람)

DRIVER_PATH = '{/opt/homebrew/bin/chromedriver}'
#DRIVER_PATH = '{C:/tools/chromedriver_win32}'
driver = webdriver.Chrome(DRIVER_PATH)

# 로그인
def login():
    driver.get("https://mail.aproele.com/gw/uat/uia/egovLoginUsr.do")

    id_box = driver.find_element(By.ID, 'userId')
    pw_box = driver.find_element(By.ID, 'userPw')
    btn_login = driver.find_element(By.CLASS_NAME, 'log_btn')

    id_box.click()
    id_box.send_keys(ID)

    pw_box.click()
    pw_box.send_keys(PW)

    btn_login.click()


def downloadPdf():
    driver.get('https://mail.aproele.com/eap/ea/eadoc/EaDocList.do?menu_no=' + MENU_ID)
    driver.implicitly_wait(10)

    headers = {
        'Accept': '*/*',
        'Content-Type': 'application/x-www-form-urlencoded',
    }

    cookies = driver.get_cookies()
    newCookie = ''
    for cookie in cookies:
        if len(newCookie) > 0:
            newCookie = newCookie + '; '
        newCookie = newCookie + '{}={}'.format(cookie['name'], cookie['value'])
    headers['Cookie'] = newCookie + '; Path=/; Domain=aproele.com; Secure; HttpOnly;'

    data = {
        'page': '1',
        'pageSize': PAGE_SIZE,
        'nMenuID': MENU_ID,
        'sfrDt': '2015-01-01',
        'stoDt': datetime.datetime.now().strftime('%Y-%m-%d'),
        'sReadYN': READ_TYPE,
        'sTitleTP': '0',
        'hiddenSelect': PAGE_SIZE,
    }

    response = requests.post("https://mail.aproele.com/eap/ea/eadoc/GetEaDocList.do", headers=headers, data=data)

    # print("resonsee", response)
    # print("Status Code", response.status_code)
    # print("JSON Response ", response.json())

    resData = json.loads(response.content)
    lists = resData['list']

    print("list", lists)
    print("  ----- download --------")

    index = 1
    for list in lists:
        print("{}. {}".format(index, list['DOC_TITLE']))
        driver.get("https://mail.aproele.com/eap/ea/docpop/EAAppDocPrintPop.do?doc_id={}&form_id={}&p_doc_id=0&mode=PDF&doc_auth=-1&type=1&area={}&spDocId={};0".format(list['DOC_ID'], list['FORM_ID'], PDF_TYPE, list['DOC_ID']))
        time.sleep(5)
        index = index + 1


print("# 1. Login  ---------------------------------------------")
login()

print("# 2. Download PDF  --------------------------------------")

downloadPdf()

print("# 3. End ------------------------------------------------")
time.sleep(10)
