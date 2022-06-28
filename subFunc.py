import csv
import subprocess

def list_up(last_page): # 페이지 전환 후 extract_row_info 실행
    sets = []
    for page in range(last_page):
        echo(f"Scrapping Ongoing : page {page}") # subprocess
        # 페이지 전환
        result = requests.get()
        for result in results:
            unavailable = "unavailable"

def describe_in_excel(sets):
    file = open("DB_추출_리스트", mode="w")
    writer = csv.writer(file)
    writer.writerow("번호, 사례번호, 기관명, 대상자 성명, ")
    for set in sets:
        writer.writerow(list(set.values()))
    return