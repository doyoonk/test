#!/Users/dykim/src/python/test/venv/bin/python
# -*- coding: utf-8 -*-

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import os
import sys
import requests
import datetime
import json
import time
import shutil
import traceback
import subprocess

################################################################################
# 설정
################################################################################
WHO = "dykim" # "dbclose or dykim"
ID = "mskang"
PW = "aproele0320!@"

# DOWNLOAD_PATH: 파일 다운로드 경로 - 임시 저장 용  (임시 저장 후 SAVE_PATH의 서브 경로로 파일이 이동됨)
# SAVE_PATH: 최종 파일 저장 경로 (이 경로에 서브 디렉토리가 생성됨 - [문서 등록날짜_문서번호], 예시) /20230426_16435/ )
if WHO == "dbclose":
    DOWNLOAD_PATH = '/Users/dbclose/Download/temp'
    SAVE_PATH = '/Users/dbclose/Download/menu'
    DRIVER_PATH = '{/opt/homebrew/bin/chromedriver}'
else:
    DOWNLOAD_PATH = '/Users/dykim/Downloads/temp'
    SAVE_PATH = '/Users/dykim/Downloads/docs'
    DRIVER_PATH = '{/opt/local/bin/chromedriver}'
download_item = 1

#DRIVER_PATH = '{C:/tools/chromedriver_win32}'
# PDF_TYPE = '5'  ##---------->>  사용안함. 첨부파일이 있으면 5, 없으면 3으로 자동 설정됨.
START_PAGE = 1              # 데이터를 다운
START_INDEX = 1
# 2023/05/26 13:33 33-26 --> 34-1 33-26
# 2023/05/26 17:47 83-14 --> 84-1
# 2023/05/26 19:08 89-49 --> 90-1
# 2023/05/28 11:59 129-28 --> 129-29
# 2023/05/28 17:32 139-36 -->
# 2023/05/29 11:49 421-48 -->
# 2023/05/29 13:29 455-21 -->

PAGE_SIZE = 50               # 한 번에 조회 할 데이터 수

option = Options()
option.add_experimental_option('prefs', {
    'download.default_directory': DOWNLOAD_PATH,
    'download.prompt_for_download': False,
    'profile.default_content_setting_values.automatic_downloads': 1     # 여러 파일 다운로드 권한 허용.
})

#driver = webdriver.Chrome(DRIVER_PATH, options=option)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=option)

now = datetime.datetime.now()
LOG_FILE_PATH = SAVE_PATH + '/' + now.strftime("%Y-%m-%d_%H%M%S") + '.txt'
with open(LOG_FILE_PATH, 'w') as log:
    log.write("========= 파일 카운드가 일치하지 않는 데이터 (" + now.strftime("%Y-%m-%d %H:%M:%S") + ") =================\n")

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def latest_download_file(source):
    files = sorted(os.listdir(source))#, key=os.path.getmtime)
    #newest = files[-1]
    newest = ""
    for file in files:
        newest = newest + file
    return newest

def moveFiles(source, target):
    try:
        fileends = "crdownload"
        while "crdownload" == fileends:
            time.sleep(1)
            newest_file = latest_download_file(source)
            if "crdownload" in newest_file:
                fileends = "crdownload"
            else:
                fileends = "none"
        # 파일 이동.
        #shutil.move(source + '/*', target)
        subprocess.call("mv " + source + "/* " + target, shell=True)

        """
        files = os.listdir(source)

        downloging_files = 0
        for file in files:
            if (file.find('crdownload') > -1):
                downloging_files = downloging_files + 1

        if (downloging_files > 0):
            time.sleep(0.5)
            moveFiles(source, target)


        #print("source:", source)
        for file in files:
            #print('- file: ', source + '/' +  file)
            if os.path.isfile(target + '/' + file):
                os.remove(target + '/' + file)

            # 파일 이동.
            shutil.move(source + '/' + file, target)
        """
    except:
        print('moveFiles error!')
        traceback.print_exc()
        time.sleep(0.5)
        moveFiles(source, target)


# 로컬에 저장된 파일이 있는가?
def existsLocalFile(file_path):
    if os.path.exists(file_path):
        file_size = os.path.getsize(file_path)
        return file_size > 0
    else:
        return False


# 다시 다운로드 해야 하는가?
def needReDownload(file_path):
    return download_item == 0 and not existsLocalFile(file_path)


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
        'p_to_dt': '2023-04-30' # datetime.datetime.now().strftime('%Y-%m-%d')
    }

    has_next = True
    page = START_PAGE
    count = 0
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
        print('')
        print('------------------------------------------------------')
        print("----- download: " + str(page) + ' page... ')
        print('------------------------------------------------------')
        index = 1

        for list in lists:
            item_count = 0
            if page < START_PAGE or (page == START_PAGE and index < START_INDEX):
                index = index + 1
                continue

            save_path = SAVE_PATH + '/' + list['CREATED_DT'][0:4] + '/' + list['CREATED_DT'][0:7] + '/' + list['CREATED_DT'].replace('-', '') + '_' + str(list['DOC_ID'])
            docFilename = '[' + list['DOC_NO'] + ']' + list['DOC_TITLE'].replace(':', '_').replace('/', '_') + '.pdf'
            docFilepath = save_path + '/' + docFilename

            if list['FILE_CNT'] == 0:
                PDF_TYPE = '3'
            else:
                PDF_TYPE = '5'

            docFileMessage = "{}-{}:{}".format(page, index, docFilename, save_path)
            print(docFileMessage)
            print(" - 경로: {} ".format(save_path))

            """
            if download_item == 1:
                print(docFileMessage)
                print(" - 경로: {} ".format(save_path))

            elif download_item == 0 and existsLocalFile(docFilepath):
                print(docFileMessage + ' (파일이 존재함)')
                print(" - 경로: {} ".format(save_path))

            else:
                print(docFileMessage + ' (파일이 존재하지 않음)', file=sys.stderr)
                print(" - 경로: {} ".format(save_path), file=sys.stderr)
            """

            # pdf 다운로드
            if (download_item == 1 or needReDownload(docFilepath)):
                driver.get("https://mail.aproele.com/eap/ea/docpop/EAAppDocPrintPop.do?doc_id={}&form_id={}&p_doc_id=0&mode=PDF&doc_auth=1&type=1&area={}&spDocId={};0".format(list['DOC_ID'], list['FORM_ID'], PDF_TYPE, list['DOC_ID']))
                time.sleep(1)

            item_count = item_count + 1

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
                    item_count = item_count + 1

                    attachFilename = item.find_element(By.NAME, 'fileNm').text
                    attachFilepath = save_path + '/' + attachFilename

                    try:
                        #print(' - item: ', item)
                        #print('save_path: ' + save_path)
                        #print('attachFilename: ' + attachFilename)

                        attachFileMessage = ' - 첨부파일: ' + attachFilename

                        if download_item == 1:
                            print(attachFileMessage)
                        elif download_item == 0 and existsLocalFile(attachFilepath):
                            print(attachFileMessage + ' (파일 존재함)')
                        else:
                            print(attachFileMessage + ' (파일 존재하지 않음)', file=sys.stderr)

                        if (download_item == 1 or needReDownload(attachFilepath)):
                            #time.sleep(1)
                            item.click()
                            time.sleep(1)
                            #driver.execute_script("arguments[0].click();", item)
                    except:
                        print(' ㄴ error: item.click() ==> javascript.click()')
                        if (download_item == 1 or needReDownload(attachFilepath)):
                            #time.sleep(1)
                            #item.click()
                            driver.execute_script("arguments[0].click();", item)
                            time.sleep(1)


            # 파일 이동
            # print(' # 파일 이동: TEMP -> ', save_path, flush=True)

            createFolder(save_path)
            moveFiles(DOWNLOAD_PATH, save_path)


            # 파일 이동 후 파일 갯수 검증.
            file_count = 0
            for filename in os.listdir(save_path):
                if os.path.isfile(os.path.join(save_path, filename)):
                    file_count += 1

            if item_count <= file_count:
                print(f' ===> {page}-{index}: 파일 수가 정상입니다. ({item_count} <= {file_count})')
                # with open(LOG_FILE_PATH, 'a') as log:
                #     log.write('[{}] {}-{}: {}\n'.format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), page, index, list['DOC_TITLE'], "파일 수가 일치하지 않음" ))

            else:
                print(f' ===> {page}-{index}: 파일 수가 일치하지 않습니다. ({item_count} > {file_count})', file=sys.stderr)

                with open(LOG_FILE_PATH, 'a') as log:
                    log.write('[{}] {}-{}: {}\n'.format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), page, index, list['DOC_TITLE'], "파일 수가 일치하지 않음" ))


            index = index + 1
            count = count + item_count

            print('')


        if (startCount < PAGE_SIZE):
            has_next = False
            break
        page = page + 1

# 저장 경로 생성.
shutil.rmtree(DOWNLOAD_PATH)
#shutil.rmtree(SAVE_PATH)
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
