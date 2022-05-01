import os
import re
import sys
import json
import shutil
import datetime
import codecs
import configparser
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
def PrintSetUp(downloadsDirPath):
    #印刷としてPDF保存する設定
    chopt = Options()
    appState = {
        "recentDestinations": [
            {
                "id": "Save as PDF",
                "origin": "local",
                "account":""
            }
        ],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isLandscapeEnabled": True, #印刷の向きを指定 tureで横向き、falseで縦向き。
        "pageSize": 'A4', #用紙タイプ(A3、A4、A5、Legal、 Letter、Tabloidなど)
        #"mediaSize": {"height_microns": 355600, "width_microns": 215900}, #紙のサイズ　（10000マイクロメートル = １cm）
        "marginsType": 1, #余白タイプ #0:デフォルト 1:余白なし 2:最小
        #"scalingType": 3 , #0：デフォルト 1：ページに合わせる 2：用紙に合わせる 3：カスタム
        #"scaling": "141" ,#倍率
        #"profile.managed_default_content_settings.images": 2,  #画像を読み込ませない
        "isHeaderFooterEnabled": False, #ヘッダーとフッター
        "isCssBackgroundEnabled": True, #背景のグラフィック
        "isDuplexEnabled": False, #両面印刷 trueで両面印刷、falseで片面印刷
        "isColorEnabled": True, #カラー印刷 trueでカラー、falseで白黒
        #"isCollateEnabled": True #部単位で印刷
    }
    
    prefs = {'printing.print_preview_sticky_settings.appState':
            json.dumps(appState),
            "profile.default_content_settings.popups": 1,
            "download.default_directory":
                os.path.abspath(downloadsDirPath),
            "savefile.default_directory": 
                os.path.abspath(downloadsDirPath),
            "directory_upgrade": True
    } #appState --> pref
    chopt.add_experimental_option('prefs', prefs) #prefs --> chopt
    chopt.add_argument('--kiosk-printing') #印刷ダイアログが開くと、印刷ボタンを無条件に押す。
    return chopt

def execute_pdf_download(driver, url, filePath, clientPath, cnt):
    try:
        driver.get(f"file:///{url}")
        sleep(7)
        driver.execute_script('return window.print()')
        sleep(3)
        shutil.move(filePath, clientPath)
        logger.debug(f"Finish moving the pdf file: {clientPath}") 
    except Exception as err:
        if os.path.exists(filePath) or cnt > 3:
            return
        else:
            cnt += 1
            execute_pdf_download(driver, filePath, cnt)

### main_script ###
if __name__ == '__main__':

    if len(sys.argv) == 1:
        logger.debug('No parameter')
        exit(1)

    param = int(sys.argv[1])

    downloadsDirPath = './pdf'
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

    options = PrintSetUp(downloadsDirPath)
    
    try:
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        
        driver.get(defaultUrl)

        config = configparser.ConfigParser()
        config.read_file(codecs.open("domainInfo.ini", "r", "utf8"))
        domains = config.sections()
        filePath = f"{downloadsDirPath}/順位計測結果.pdf"

        for domain in domains:
            client = config[domain]['CLIENT']
            clientPath = f'{requestReportPath}/{client}/'
            os.makedirs(clientPath, exist_ok=True)
            clientPath += f"{name}_順位計測結果_{year}_{month}.pdf"
 
            if domain == "wakigacenter.com":
                name = "リオラビューティークリニック子供わきが"
                client = config[domain]['CLIENT']
                url = f"{os.environ['RANK_REPORT_PATH']}/{year}/{month}/{domain}_kodomo_wakiga.html"
                execute_pdf_download(driver, url, filePath, clientPath, 0)

            name = config[domain]['NAME']
            client = config[domain]['CLIENT']
            url = f"{os.environ['RANK_REPORT_PATH']}/{year}/{month}/{domain}.html"
            execute_pdf_download(driver, url, filePath, clientPath, 0)
            logger.debug(f"pdf download comlete: {name}")

        driver.close()
        driver.quit()

        exit(0)
    except Exception as err:
        logger.debug(f'create_rank_result: {err}')
        exit(1)