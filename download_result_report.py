import os
import re
import sys
import shutil
import datetime
import gspread
import openpyxl
import codecs
import configparser
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from fake_useragent import UserAgent
from oauth2client.service_account import ServiceAccountCredentials
from webdriver_manager.chrome import ChromeDriverManager

# Logger setting
from logging import getLogger, FileHandler, DEBUG
logger = getLogger(__name__)
today = datetime.datetime.now()
os.makedirs('./log', exist_ok=True)
handler = FileHandler(f'log/{today.strftime("%Y-%m-%d")}_result.log', mode='a')
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False

### functions ###
def modify_excel(fileName, sheetName):
    wb = openpyxl.load_workbook(fileName)

    for ws in wb:
        if ws.title != sheetName:
            wb.remove(ws)

    wb.save(fileName)

### main_script ###
if __name__ == '__main__':

    if len(sys.argv) == 1:
        logger.debug('No parameter')
        exit(1)

    param = int(sys.argv[1])

    downloadsDirPath = './excel'
    os.makedirs(downloadsDirPath, exist_ok=True)
    shutil.rmtree(downloadsDirPath)
    os.makedirs(downloadsDirPath, exist_ok=True)

    thismonth = datetime.datetime(today.year, today.month, 1)

    if param == 1:
        lastmonth = thismonth + datetime.timedelta(days=-1)
    else:
        lastmonth = thismonth

    year = lastmonth.strftime("%Y")
    month = lastmonth.strftime("%m")

    requestReportPath = f"{os.environ['REQUEST_REPORT_PATH']}/{year}/{month}"
    os.makedirs(requestReportPath, exist_ok=True)

    defaultUrl = "https://google.co.jp/"
    ua = UserAgent()
    logger.debug(f'download_result_report: UserAgent: {ua.chrome}')

    options = Options()
    options.add_argument(f'user-agent={ua.chrome}')

    prefs = {
        "profile.default_content_settings.popups": 1,
        "download.default_directory":
                os.path.abspath(downloadsDirPath),
        "directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)
    
    try:
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        
        driver.get(defaultUrl)

        config = configparser.ConfigParser()
        config.read_file(codecs.open("clientInfo.ini", "r", "utf8"))
        clients = config.sections()

        for client in clients:
            SPREADSHEET_ID = config[client]['SSID']
            if SPREADSHEET_ID == 'None':
                continue
            url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=xlsx"
            driver.get(url)
            sleep(3)
            fileName = fr"{client}_成果表"
            filePath = fr"{downloadsDirPath}/{fileName}.xlsx"
            logger.debug(f"download: {filePath}")
            modify_excel(filePath, f"{year}{month}")
            clientPath = f"{requestReportPath}/{client}/{client}/"
            os.makedirs(clientPath, exist_ok=True)
            clientPath += f"{fileName}_{year}_{month}.xlsx"
            shutil.move(filePath, clientPath)

        driver.close()
        driver.quit()

        exit(0)
    except Exception as err:
        logger.debug(f'download_result_report: {err}')
        exit(1)