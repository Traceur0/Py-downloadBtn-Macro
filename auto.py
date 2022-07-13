import re # trimming string
import json # detaching key value with code
# SELENIUM MODULE
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.ie.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as Condition
import pyautogui as gui

ie_options = webdriver.IeOptions()
ie_options.attach_to_edge_chrome = True
ie_options.edge_executable_path = r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
ser = Service(r"C:/browserDriver/IEDriverServer(x86).exe")
browser = webdriver.Ie(options=ie_options, service=ser)
wait = WebDriverWait(browser, 60)

with open("./key/text.json", "rt", encoding="UTF8") as text_json:
    text = json.load(text_json)
LOGIN_PAGE = text["pages"]["login"]
MAIN_PAGE = text["pages"]["main"]
SEARCH_PAGE = text["pages"]["search"]
STATEMENT = text["reason_d"]


def load_page():
    browser.get(LOGIN_PAGE)
    status_set = gui.alert(text='조회 날짜 등 세팅을 최종적으로 완료, 다운로드 준비가 되면 확인을 눌러주세요', title='Confirmation', button='Done')
    if status_set == "Done":
        return


def pagination(): # extract last page number
    # route : div[@id='gridPaging01']/div[@class='paginate']/ol/li[@class='last']
    try:
        wait.until(Condition.element_to_be_clickable(browser.find_element(By.XPATH, "//*[@class='last']/a")))
    except:
        return 1
    p_num = browser.find_element(By.XPATH, "//div[@id='pageBtns']/div/b").text
    # last_page = re.sub(r"[가-힣/:]", "", p_num)
    last_page = re.search('페이지 : 1 / (.+)', p_num)
    if last_page:
        last_page = last_page.group(1)
    return int(last_page)


def describe_DR() :
    download_reason = '''
    다운로드 사유를 입력해 주세요
    *입력하신 사유는 페이지 다운로드마다 반복 입력됩니다*
    '''
    reason_text = gui.prompt(text=download_reason, title='다운로드 사유 입력창')
    return reason_text


def xlsx_extraction(DR_reason): 
    download_btn = browser.find_element(By.XPATH, "//a[@id='btnExcel']")
    wait.until(Condition.element_to_be_clickable(download_btn))
    download_btn.click()
    browser.implicitly_wait(time_to_wait=60)
    browser.switch_to.window(browser.window_handles[-1]) # switch to pop-up browser(tab)
    download_reason = browser.find_element(By.XPATH, "//textarea[@id='conts']")
    download_reason.send_keys(DR_reason)
    confirm_btn = browser.find_element(By.XPATH, "//a[@class='btn_save']") ###
    print(f"Found {confirm_btn}")
    confirm_btn.click()


def switch_page(page):
    browser.implicitly_wait(time_to_wait=60)
    browser.switch_to.window(browser.window_handles[0]) # back to original browser
    page_btn = browser.find_element(By.LINK_TEXT, f"{page}") ### ?
    wait.until(Condition.element_to_be_clickable(page_btn))
    page_btn.click()


def auto_download(last_page):
    reason = describe_DR()
    for page in range(1, last_page):
        xlsx_extraction(reason)
        print(f"Downloading page {page} in excel format...")
        switch_page(page)
        browser.implicitly_wait(time_to_wait=45)
        if page == last_page:
            browser.quit()

# XPATH가 유효한지 다시 확인 -> 확인 필요한 부분 주석으로 표시