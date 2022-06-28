# SELENIUM NOMDULE
from codecs import ascii_decode
import string
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.ie.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as Condition
# HTTP LIBRARY MODULE
# import requests
# from bs4 import BeautifulSoup
import re # trimming string
import json # departing key value
import time


ie_options = webdriver.IeOptions()
ie_options.attach_to_edge_chrome = True
ie_options.edge_executable_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
ser = Service("C:/Scrapper/browserDriver/IEDriverServer(x86).exe")
browser = webdriver.Ie(options=ie_options, service=ser)
wait = WebDriverWait(browser, 60)

# request = requests.get(url)
# parsing = BeautifulSoup(request.text ,"html.parser")

def waiting_load_n_click(path):
        tag_name = browser.find_element(By.XPATH, path)
        wait.until(Condition.element_to_be_clickable(tag_name))
        tag_name.click()

def dynamicSearch():
    ## 변경 부분
    # id_input = browser.find_element(By.CSS_SELECTOR, value='fieldset > form > ul > li > #userId')
    # id_input = browser.find_element(By.XPATH, "//input[@id='userId']")
    # id_input.send_keys("admin")
    with open("./key/site_info.json", "r") as site_json:
        site = json.load(site_json)
    ID = site["login_id"]
    # DISK_NAME = site["disk_name"]
    LOGIN_PAGE = site["pages"]["login"]
    MAIN_PAGE = site["pages"]["main"]
    SEARCH_PAGE = site["pages"]["search"]

    browser.get(LOGIN_PAGE)
    # browser.implicitly_wait(time_to_wait=10) # 명시적 대기 / 로드 될때까지 대기할 수 있도록 설정

    # browser.execute_script(f"document.querySelector('#userId').value=\'{ID}\'")
    # browser.implicitly_wait(time_to_wait=5)

    ## 대체 부분
    # browser.switch_to.frame('') # iframe 으로 전환

    # login_btn = browser.find_element(By.XPATH, "//input[@class='btn_login']")
    # login_btn.click()
    
    # 공인인증서 입력

    # login_btn.click() # 공인인증서 입력 후
    
    # waiting_load_n_click("//button[@id='xwup_media_removable']")

    # key_disk = browser.find_element(By.NAME, DISK_NAME) #수정 필요
    # wait.until(Condition.element_to_be_clickable(key_disk))
    # key_disk.click()
    # waiting_load_n_click("select_disk", "//div[@class='content-menu-layout']/ul/li[@class='context-menu-item-unfocused']")

    browser.implicitly_wait(time_to_wait=50)
    url_now = browser.current_url
    wait.until(Condition.element_to_be_clickable(browser.find_element(By.XPATH, "//img[@id='iconMnu_06']")))
    if url_now == MAIN_PAGE :
        browser.execute_script(f'window.open(\"{SEARCH_PAGE}\");')
    # browser.execute_script(f'window.open(\"{SEARCH_PAGE}\");') # 새 브라우저 열기
    browser.implicitly_wait(time_to_wait=10)
    scrapping()


    
    # waiting_load_n_click("통계및조회", "//img[@id='iconMnu_06']")
    # waiting_load_n_click("조회", "//img[@id='iconMnu_06']")
    # waiting_load_n_click("서비스제공조회", "//div[@class='left3bg']/div/a")
    
    # browser.quit()

def scrapping(): # 동작 횟수 = 1 * 10 * pagination 수
    with open("./key/site_info.json", "r") as site_json:
        site_info = json.load(site_json)
    SEARCH_PAGE = site_info["pages"]["search"]
    browser.get(SEARCH_PAGE)
    browser.implicitly_wait(time_to_wait=10)
    waiting_load_n_click("//a[@id='btnSearch']")
    extracting_row()

def extracting_row(): # 한 페이지에서 정보 받아오기 (10줄의 정보) or (1줄)
    row_num = -1 # default : 0, range : 0~9
    one_row = []
    for row_num in range(9):
        # div[@class='jqx-grid-content']/div[@id='contnetablejqxgrid01']/div[@id='row{row_num}jqxgrid01']/div[@id='jqx-grid-cell']
        browser.implicitly_wait(time_to_wait=5)
        extracting = browser.find_element_by_css_selector(f"div#row{row_num}jqxgrid01").get_attribute("innerHTML")
        adjusted = re.sub(r"[A-Za-z]", "", extracting)
        one_row.append(adjusted)
    
    # num = extracting.select_one("")
    # apparatusName = extracting.select_one("")
    # caseNum = extracting.select_one("")
    # victim = extracting.select_one("")
    # assailant = extracting.select_one("")
    # s_Date = extracting.select_one("")
    # counselor_in_s = extracting.select_one("")
    # facility_assailant = extracting.select_one("")
    # goal_s = extracting.select_one("")
    # served_s = extracting.select_one("")
    # s_category = extracting.select_one("")
    # method_of_serving = extracting.select_one("")
    # number_of_s = extracting.select_one("")
    # counselor_in_charge = extracting.select_one("")
    # case_counselor = extracting.select_one("")
    # { # 15항목
    #     "No" : num,
    #     "기관명" : apparatusName,
    #     "사건번호" : caseNum,
    #     "아동명" : victim,
    #     "학대행위자명" : assailant,
    #     "서비스제공일시" : s_Date,
    #     "서비스제공자" : counselor_in_s,
    #     "대상자" : facility_assailant,
    #     "서비스세부목표" : goal_s,
    #     "제공 서비스" : served_s,
    #     "제공구분" : s_category,
    #     "제공 방법" :  method_of_serving,
    #     "서비스제공횟수" : number_of_s,
    #     "사건담당자" : counselor_in_charge,
    #     "사례담당자" : case_counselor,
    # }
    return print(one_row)

def pagination(): # 반복횟수
    # inspect_num = parsing.select_one("div#gridPaging01 > div.paginate > ol > li.last")
    # 경로 : div[@id='gridPaging01']/div[@class='paginate']/ol/li[@class='last']
    number = browser.find_element(By.XPATH, "//li[@class='last']").text
    last_page = int(re.sub(r"[A-Za-z()]", "", number))
    return last_page    