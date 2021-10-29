from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from RPA.Excel.Application import Application
import pandas as pd
from dotenv import dotenv_values
import json

config = dotenv_values(".env")

browser = Selenium()
file_system = FileSystem()
url = config.get('DOMAIN')
excel_name = config.get('EXCEL')
agencia = config.get('AGENCIA')
output = config.get('OUTPUT')
app = Application()
excel = Files()
fs = FileSystem()

div_agency_container = 'id:agency-tiles-container'

div_learn = 'id:block-nodeblock-learn-it-spending'

btn_navigate_to_last = 'css:#investments-table-object_last.disabled'

list_show_itens = 'name:investments-table-object_length'

table_investiment = 'id:investments-table-object'

span_gerando_arquivo = 'css:#business-case-pdf > span > img'

link_downaload_pdf = 'css:#business-case-pdf  a'

table_link_details = 'css:#investments-table-object td a'

link_details_ag = 'xpath://*[contains(text(),"' + agencia + '")]'

WORKSHEET_AG = "Agências"

link_home = "#home-dive-in"

SECONDS = '30 seconds'


def store_web_page_content():
    create_output(output + '/' + excel_name)
    browser.open_available_browser(url)
    browser.set_download_directory(fs.absolute_path(output))
    browser.set_browser_implicit_wait(SECONDS)
    browser.click_link(link_home)
    table = get_table()
    append_excel(output + '/' + excel_name, table, WORKSHEET_AG, True)
    navigate_agencia()
    individual_investments = extract_individual_investments()
    res = json.loads(individual_investments.to_json(orient='records'))
    append_excel(output + '/' + excel_name, res, agencia, True)
    download_links()


def navigate_agencia():
    browser.click_element(browser.find_elements(link_details_ag))


def download_links():
    links = browser.find_elements(table_link_details)
    for li in links:
        link = str(li.get_attribute('href'))
        browser.open_available_browser(link)
        browser.set_download_directory(fs.absolute_path(output))
        browser.set_browser_implicit_wait(SECONDS)
        url_arr = str(link).split("/")
        name = url_arr.reverse()
        name = url_arr[0]
        browser.click_element(link_downaload_pdf)
        browser.wait_until_element_is_visible(span_gerando_arquivo)
        browser.wait_until_element_is_not_visible(span_gerando_arquivo)
        fs.wait_until_created(output + '/' + str(name) + ".pdf", 50)
        browser.close_browser()


def extract_individual_investments():
    browser.wait_until_element_is_visible(table_investiment)
    browser.select_from_list_by_label(list_show_itens, 'All')
    browser.wait_until_element_is_visible(btn_navigate_to_last)
    df = pd.read_html(
        browser.find_element(table_investiment).get_attribute('outerHTML'))[0]
    return df


def get_table():
    browser.wait_until_element_is_enabled(div_learn)
    browser.wait_until_element_is_visible(div_agency_container)
    return_arr_ = " var ele = document.querySelector('#agency-tiles-widget').querySelectorAll('div.tuck-5');" \
                  "var arr = [];" \
                  "ele.forEach(e => " \
                  "arr.push({title : e.querySelector('span.h4').innerText, valor:e.querySelector('span.h1').innerText}));" \
                  " return arr;"
    table = browser.execute_javascript(
        return_arr_
    )
    return table


def append_excel(path, table, worksheet, header):
    try:
        excel.open_workbook(path)
        excel.create_worksheet(worksheet, exist_ok=True)
        excel.set_active_worksheet(worksheet)
        excel.append_rows_to_worksheet(table, header=header)
        excel.save_workbook(path)
    finally:
        excel.close_workbook()


def create_output(path):
    try:
        fs.create_directory(output)
        fs.empty_directory(output)
        excel.create_workbook(excel_name)
        excel.save_workbook(path)
        excel.close_workbook()
    finally:
        excel.close_workbook()


def main():
    try:
        store_web_page_content()
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
