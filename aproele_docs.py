
import requests
import json

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import MoveTargetOutOfBoundsException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException

from urllib3._collections import HTTPHeaderDict

def wait_for_ajax(driver):
    wait = WebDriverWait(driver, 15)
    try:
        wait.until(lambda driver: driver.execute_script('return jQuery.active') == 0)
        wait.until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
    except Exception as e:
        pass

with requests.Session() as s:
    DRIVER_PATH = '{C:/tools/chromedriver_win32}'
    driver = webdriver.Chrome(DRIVER_PATH)
    driver.get("https://mail.aproele.com")
    wait_for_ajax(driver)
    headers = {
        'Accept': '*/*',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.9,ko;q=0.8',
        'Connection': 'keep-alive',
        'Content-Length': '183',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Cookie': 'JSESSIONID=0419A05B6C2F5426F080AA076003A417',
        'Host': 'mail.aproele.com',
        'Origin': 'https://mail.aproele.com',
        'Referer': 'https://mail.aproele.com/eap/ea/eadoc/EaDocList.do?menu_no=2002060000',
        'sec-ch-ua': '"Chromium";v="112", "Google Chrome";v="112", "Not:A-Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': 'Windows',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36'
    }

    cookies = driver.get_cookies()
    newCookie = ''
    for cookie in cookies:
        s.cookies.set(cookie['name'], cookie['value'])
        if len(newCookie) > 0:
            newCookie = newCookie + '; '
        newCookie = newCookie + '{}={}'.format(cookie['name'], cookie['value'])
    headers['Cookie'] = newCookie

    data = {
        'page': 0, 'pageSize': '10', 'sortField': '', 'sortType': '', 'nMenuID': '2002060000',
        'sfrDt': '2023-01-09', 'stoDt': '2023-04-09',
        'sDocTp': '', 'sTitle': '', 'sContents': '', 'sUserNm': '', 'sDocSts': '', 'sDocNo': '',
        'sDeptNm': '', 'sReadYN': '', 'sExecYN': '', 'sTitleTP': '0'
        }
    for x in range(1,9999):
        data['page'] = x
        response = s.post("https://mail.aproele.com/eap/ea/eadoc/GetEaDocList.do", headers=headers, data=data)
        resData = json.loads(response.content)
        lists = resData['list']
        if len(lists) == 0:
            break
        for list in lists:
            pcSaveLinks = {}
            #print('DOC_ID={}, FORM_ID={}'.format(list['DOC_ID'], list['FORM_ID']))
            #response = s.get("https://mail.aproele.com/eap/ea/docpop/EAAppDocViewPop.do?doc_id={}&form_id={}".format(list['DOC_ID'], list['FORM_ID']))
            #print(response.content)
            driver.get("https://mail.aproele.com/eap/ea/docpop/EAAppDocViewPop.do?doc_id={}&form_id={}".format(list['DOC_ID'], list['FORM_ID']))
            wait_for_ajax(driver)
            try:
                selPDF = driver.find_element(By.ID, 'pop_sel_pdf')
                try:
                    docPDF = selPDF.find_element(By.CLASS_NAME, 'li_mode_4')
                except Exception as e:
                    docPDF = selPDF.find_element(By.CLASS_NAME, 'li_mode_2')
                ## FIXME: PDF save here
                
                # get the local storage
                #storage = LocalStorage(driver)
                #attachFile = storage['attachFileList']
                attachFile = driver.execute_script("return attachFileList")
                ## FIXME: FILE save here

                #pcSaveLinks = driver.find_element(By.LINK_TEXT, 'PC????ž¥')
                #WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.LINK_TEXT, "PC????ž¥")))
                for link in pcSaveLinks:
                    link.click()
            except Exception as e:
                print(e)
                pass
            except NoSuchElementException:  #spelling error making this code not work as expected
                pass
            #pageSource = driver.page_source
            
            #r = s.get(url, stream=True)
            #with open("dest_path", "wb") as f:
            #    for chunk in r.iter_content(chunk_size=1024):
            #        f.write(chunk)
