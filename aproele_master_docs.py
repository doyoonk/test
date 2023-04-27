from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import os
import requests
import datetime
import json
import time
import shutil

################################################################################
# 설정
################################################################################
ID = "mskang"
PW = "aproele0320!@"

# PDF_TYPE = '5'  ##---------->>  사용안함. 첨부파일이 있으면 5, 없으면 3으로 자동 설정됨.

START_PAGE = 1              # 데이터를 다운
PAGE_SIZE = 50               # 한 번에 조회 할 데이터 수

DOWNLOAD_PATH = '/Users/dbclose/Download/temp'      # 파일 다운로드 경로 - 임시 저장 용  (임시 저장 후 SAVE_PATH의 서브 경로로 파일이 이동됨)
SAVE_PATH = '/Users/dbclose/Download/menu'          # 최종 파일 저장 경로 (이 경로에 서브 디렉토리가 생성됨 - [문서 등록날짜_문서번호], 예시) /20230426_16435/ )



option = Options()
option.add_experimental_option('prefs', {
    'download.default_directory': DOWNLOAD_PATH,
    'download.prompt_for_download': False,
})


DRIVER_PATH = '{/opt/homebrew/bin/chromedriver}'
#DRIVER_PATH = '{C:/tools/chromedriver_win32}'
driver = webdriver.Chrome(DRIVER_PATH, options=option)


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def moveFiles(source, target):
    try:
        files = os.listdir(source)

        downloging_files = 0
        for file in files:
            if (file.find('crdownload') > -1):
                downloging_files = downloging_files + 1

        if (downloging_files > 0):
            time.sleep(1)
            moveFiles(source, target)

        #print("source:", source)
        for file in files:
            #print('- file: ', source + file)
            shutil.move(source + '/' + file, target)

    except OSError:
        time.sleep(1)
        moveFiles(source, target)

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

    # 마스터 페이지로 이동
    driver.execute_script("changeMode('MASTER')")



def downloadPdf():
    # driver.get('https://mail.aproele.com/eap/admin/eadoc/EADocMngList.do?menu_no=' + MENU_ID)
    # driver.implicitly_wait(10)

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
        'pageSize': PAGE_SIZE,
        'p_fr_dt': '2015-01-01',
        'p_to_dt': datetime.datetime.now().strftime('%Y-%m-%d')
    }

    has_next = True
    page = START_PAGE

    while has_next:
        data['page'] = page

        response = requests.post("https://mail.aproele.com/eap/admin/eadoc/GetEADocMngList.do", headers=headers, data=data)

        #print('current_page:', data['page'])
        # print("resonsee", response)
        # print("Status Code", response.status_code)
        # print("JSON Response ", response.json())

        resData = json.loads(response.content)
        lists = resData['list']
        startCount = resData['startCount']
        totalCount = resData['totalCount']

        #print("startCount: {}, totalCount: {}".format(startCount, totalCount))

        #print("list", lists)
        print("----- download: " + str(page) + ' page... ------------')

        index = 1
        for list in lists:

            print("{}. {}".format(index, list['DOC_TITLE']))

            if list['FILE_CNT'] == 0:
                PDF_TYPE = '3'
            else:
                PDF_TYPE = '5'

            # pdf 다운로드
            driver.get("https://mail.aproele.com/eap/ea/docpop/EAAppDocPrintPop.do?doc_id={}&form_id={}&p_doc_id=0&mode=PDF&doc_auth=1&type=1&area={}&spDocId={};0".format(list['DOC_ID'], list['FORM_ID'], PDF_TYPE, list['DOC_ID']))
            time.sleep(1)

            # 첨부파일 다운로드
            if list['FILE_CNT'] > 0:
                driver.get("https://mail.aproele.com/eap/ea/docpop/EAAppDocViewPop.do?doc_id={}&form_id={}&doc_auth=1".format(list['DOC_ID'], list['FORM_ID']))
                driver.implicitly_wait(10)

                driver.switch_to.frame('uploaderView')

                # 파일 전체 다운로드
                # driver.execute_script("fnDownloadAll()")

                # 파일 개별 다운로드
                file_area = driver.find_element(By.ID, "dzDownloaderDiv")
                items = file_area.find_elements(By.CSS_SELECTOR, "li.file_name")

                for item in items:
                    try:
                        print('  - file:', item.find_element(By.NAME, 'fileNm').text)
                        item.click()
                    except OSError:
                        time.sleep(1)
                        item.click()

                time.sleep(1)


            # 파일 이동
            save_path = SAVE_PATH + '/' + list['CREATED_DT'][0:4] + '/' + list['CREATED_DT'][0:7] + '/' + list['CREATED_DT'].replace('-', '') + '_' + str(list['DOC_ID'])
            print('save_page: ', save_path)
            createFolder(save_path)
            moveFiles(DOWNLOAD_PATH, save_path)


            index = index + 1


        if (startCount < PAGE_SIZE):
            has_next = False
            break
        page = page + 1


# 저장 경로 생성.
shutil.rmtree(DOWNLOAD_PATH)
shutil.rmtree(SAVE_PATH)
createFolder(DOWNLOAD_PATH)
createFolder(SAVE_PATH)

print('################################################################################')
print('# ')
print('# 문서 파일 다운로드: 마스터 > 전자결재관리 > 결제문서관리 > 결제문서목록')
print('# ')
print('################################################################################')
print('')

print("# 1. Login  ---------------------------------------------")
login()

print("# 2. Download PDF  --------------------------------------")
downloadPdf()

print("# 3. End ------------------------------------------------")
time.sleep(10)
