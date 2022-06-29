# SELENIUM MODULE
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.ie.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as Condition
import re # trimming string
import json # detaching key value with code

ie_options = webdriver.IeOptions()
ie_options.attach_to_edge_chrome = True
ie_options.edge_executable_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
ser = Service("C:/Scrapper/browserDriver/IEDriverServer(x86).exe")
browser = webdriver.Ie(options=ie_options, service=ser)
wait = WebDriverWait(browser, 60)

with open("./key/string.json", "r") as string_json:
    string = json.load(string_json)
    string = json.dumps(string, ensure_ascii = False)
LOGIN_PAGE = string["pages"]["login"]
MAIN_PAGE = string["pages"]["main"]
SEARCH_PAGE = string["pages"]["search"]
STATEMENT = string["pages"]["reason_of_download"]

def load_page():
    browser.get(LOGIN_PAGE)

    browser.implicitly_wait(time_to_wait=60) # for MANUAL INPUT

    url_now = browser.current_url
    wait.until(Condition.element_to_be_clickable(browser.find_element(By.XPATH, "//img[@id='iconMnu_06']")))
    if url_now == MAIN_PAGE :
        browser.execute_script(f'window.open(\"{SEARCH_PAGE}\");')

    browser.implicitly_wait(time_to_wait=60) # for MANUAL INPUT
    # browser.quit()

def pagination(): # extract last page number
    # route : div[@id='gridPaging01']/div[@class='paginate']/ol/li[@class='last']
    p_num = browser.find_element(By.XPATH, "//li[@class='last']").text
    last_page = int(re.sub(r"[A-Za-z()]", "", p_num))
    return last_page    

def xlsx_extraction(): 
    download_btn = browser.find_element(By.XPATH, f"//a[@id='']") ###
    wait.until(Condition.element_to_be_clickable(download_btn))
    download_btn.click()
    browser.implicitly_wait(time_to_wait=10)
    browser.execute_script(f"document.querySelector('#').value=\'{STATEMENT}\'")
    # 다운로드 사유 입력
    confirm_btn = browser.find_element(By.XPATH, "button[@id='']") ###
    confirm_btn.click()

def switch_page(page):
    page_btn = browser.find_element(By.XPATH, f"//a[@id='{page}']") ###
    wait.until(Condition.element_to_be_clickable(page_btn))
    page_btn.click()

def auto_download(last_page):
    for page in range(last_page):
        xlsx_extraction()
        print(f"Downloading page {page} in excel format...")
        switch_page(page)
        browser.implicitly_wait(time_to_wait=45)