# SELENIUM MODULE
from turtle import goto
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.ie.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as Condition
import re # trimming string
import json # detaching key value with code
import pyautogui as gui

ie_options = webdriver.IeOptions()
ie_options.attach_to_edge_chrome = True
ie_options.edge_executable_path = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
ser = Service("C:/Scrapper/browserDriver/IEDriverServer(x86).exe")
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
    status_login = gui.alert(text='로그인 완료시 확인을 눌러주세요', title='Confirmation', button='Done')
    if status_login == "Done":
        browser.execute_script(f'window.open(\"{SEARCH_PAGE}\");')
    status_set = gui.alert(text='날짜 등 세팅완료시 확인을 눌러주세요 (조회 페이지로 넘어가지 않는다면 수동으로 넘어가서 입력을 완료하시고 확인을 눌러주세요)', title='Confirmation', button='Done')
    if status_set == "Done":
        print("DONE!")


def pagination(): # extract last page number
    # route : div[@id='gridPaging01']/div[@class='paginate']/ol/li[@class='last']
    p_num = WebDriverWait(browser, 20).until(Condition.element_to_be_clickable(By.XPATH, "//li[@class='last']/a")).get_attribute("href") #!#!#!#
    last_page = re.sub(r"[A-Za-z()/]", "", p_num)
    return last_page

def describe_DR() :
    reason_text = gui.prompt(text='페이지 다운로드마다 반복 입력될 다운로드 사유를 입력해 주세요', title='다운로드 사유 입력창')
    return reason_text

def xlsx_extraction(DR_reason): 
    download_btn = browser.find_element(By.XPATH, f"//a[@id='btnExcel']") ###
    wait.until(Condition.element_to_be_clickable(download_btn))
    download_btn.click()
    browser.implicitly_wait(time_to_wait=10)

    # 다운로드 사유 입력
    download_reason = browser.find_element(By.XPATH, f"//inputarea[@id='btnExcel']") ###
    download_reason.send_keys(DR_reason)
    confirm_btn = browser.find_element(By.XPATH, "a[@class='btn_save']") ###
    confirm_btn.click()

def switch_page(page):
    page_btn = browser.find_element(By.LINK_TEXT, f"{page}") ### ?
    wait.until(Condition.element_to_be_clickable(page_btn))
    page_btn.click()

def auto_download(last_page):
    reason = describe_DR()
    for page in range(last_page):
        xlsx_extraction(reason)
        print(f"Downloading page {page} in excel format...")
        switch_page(page)
        browser.implicitly_wait(time_to_wait=45)
        if page == last_page:
            browser.quit()