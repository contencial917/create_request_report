import os
import re
import sys
import shutil
import datetime
import codecs
import configparser

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

### main_script ###
if __name__ == '__main__':

    if len(sys.argv) == 1:
        logger.debug('No parameter')
        exit(1)

    param = int(sys.argv[1])

    thismonth = datetime.datetime(today.year, today.month, 1)
    if param == 1:
        lastmonth = thismonth + datetime.timedelta(days=-1)
    else:
        lastmonth = thismonth
    year = lastmonth.strftime("%Y")
    month = lastmonth.strftime("%m")

    requestReportPath = f"{os.environ['REQUEST_REPORT_PATH']}/{year}/{month}"
    os.makedirs(requestReportPath, exist_ok=True)

    try:
        config = configparser.ConfigParser()
        config.read_file(codecs.open("clientInfo.ini", "r", "utf8"))
        clients = config.sections()

        downloadsDirPath = os.environ["GDRIVE_REQUEST"]

        for client in clients:
            fileName = f"{client}_ご請求書_{year}_{month}.pdf"
            filePath = f"{downloadsDirPath}/{year}{month}/{fileName}"
            clientPath = f"{requestReportPath}/{client}/"
            os.makedirs(clientPath, exist_ok=True)
            clientPath += f"{fileName}"
            shutil.copy(filePath, clientPath)

        exit(0)
    except Exception as err:
        logger.debug(f'move_request: {err}')
        exit(1)