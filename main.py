import json

from RPA.Browser.Selenium import Selenium
from dotenv import dotenv_values

from page.actions import *

config = dotenv_values(".env")

file_system = FileSystem()
url = config.get('DOMAIN')
excel_name = config.get('EXCEL')
agencia = config.get('AGENCIA')
output = config.get('OUTPUT')
path_excel = output + '/' + excel_name

app = Application()
excel = Files()
fs = FileSystem()
browser = Selenium()


def store_web_page_content():
    create_output(output, excel_name)
    setup_browser(browser, url, output)
    expand_agency(browser)
    table = get_table(browser)
    append_excel(path_excel, table, worksheet_ag, True)
    navigate_agencia(browser, agencia)
    individual_investments = extract_individual_investments(browser)
    res = json.loads(individual_investments.to_json(orient='records'))
    append_excel(path_excel, res, agencia, True)
    download_links(browser, output)


def main():
    try:
        store_web_page_content()
    finally:
        browser.close_browser()


if __name__ == "__main__":
    main()
