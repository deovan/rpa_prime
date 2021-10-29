import pandas as pd
from RPA.Excel.Application import Application
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem

from page.selectors import *

app = Application()
excel = Files()
fs = FileSystem()

SECONDS = '30 seconds'


def setup_browser(browser, url, output):
    browser.open_available_browser(url)
    browser.set_download_directory(fs.absolute_path(output))
    browser.set_browser_implicit_wait(SECONDS)


def expand_agency(browser):
    browser.click_link(link_home)


def navigate_agencia(browser, agencia):
    browser.click_element(browser.find_elements(link_details_ag(agencia)))


def download_links(browser, output):
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


def extract_individual_investments(browser):
    browser.wait_until_element_is_visible(table_investiment)
    browser.select_from_list_by_label(list_show_itens, 'All')
    browser.wait_until_element_is_visible(btn_navigate_to_last)
    df = pd.read_html(
        browser.find_element(table_investiment).get_attribute('outerHTML'))[0]
    return df


def get_table(browser):
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


def create_output(output, excel_name):
    try:
        fs.create_directory(output)
        fs.empty_directory(output)
        excel.create_workbook(excel_name)
        excel.save_workbook(output + "/" + excel_name)
        excel.close_workbook()
    finally:
        excel.close_workbook()
