import urllib

from bs4 import BeautifulSoup
import requests
import regex as re
import pandas as pd
from tabulate import tabulate
import os.path
import time
import json
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import Mustomb.FetchClass
import Mustomb.UtilClass

def fnc(pg, bildiriNumber):
    try:
        currPage = requests.get(pg)
    except requests.ConnectionError:
        print("ConnectionError")
        return
    except requests.exceptions.Timeout:
        print("Timeout")
        return
    except requests.exceptions.TooManyRedirects:
        print("TooManyRedirects")
        return
    except requests.exceptions.RequestException as e:
        print("RequestException")
        return

    #currPage = requests.get(pg)
    soup = BeautifulSoup(currPage.text, 'html.parser')

    for part in soup.find_all('h1'):
        if re.search("Finansal Rapor.*", part.text):
            fd.fetchData(soup, bildiriNumber, "filterStocksOnly")

    f2 = open("//Users//myilmaz//Documents//bist//bilancolar_yeni//SonBildiriNo.txt", "r+")
    f2.write(bildiriNumber)
    f2.close()

def fncMultiple(s):
    content = s
    content = content.strip()
    contentList = content.split("-")
    contentList = ["https://www.kap.org.tr/tr/Bildirim/" + x for x in contentList]
    print(contentList)

    for link in contentList:
        fnc(link)

#fnc("https://www.kap.org.tr/tr/Bildirim/846388")
#fncMultiple("873010-873193-873196-873216")

dirname = os.path.dirname(__file__)
#filename = os.path.join(dirname, "kap_linkler.txt")


file = open("//Users//myilmaz//Documents//bist//bilancolar_yeni//SonBildiriNo.txt", "r+")
content_=file.readline()
#content_=content_.strip()
#contentList_ =content_.split("-")
sonBildiriNo= content_
print("sonBildiriNo= ",sonBildiriNo)

print()
file.close()
s = Mustomb.UtilClass.FinMethods()
fd=  Mustomb.FetchClass.Fetch()
for i in range(999999999):
    bildiriNo = str(int(sonBildiriNo) + 1 + i)
    link = "https://www.kap.org.tr/tr/Bildirim/" + bildiriNo
    print("link= ", link)
    print(" bildiri numarası: ", bildiriNo)
    fnc(link, bildiriNo)
    time.sleep(0.51)