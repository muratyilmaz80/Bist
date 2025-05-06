#bildiri no'ları almak için
import os.path
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import regex as re
from tabulate import tabulate
import os.path

dirname = os.path.dirname(__file__)
file = open("/Users/myilmaz/Documents/bist/bilancolar_yeni/SonBildiriNo.txt", "r+")
content_=file.readline()
sonBildiriNo= content_
print("sonBildiriNo= ",sonBildiriNo)


def load_page(url):
    chrome_driver_path = "//Users//myilmaz//PycharmProjects//chromedriver//chromedriver"  # <- set this correctly
    service = Service(executable_path=chrome_driver_path)
    options = Options()
    options.add_argument("--headless")  # Headless mode
    driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)

    # Load the page
    # url = "https://www.kap.org.tr/tr/Bildirim/1392724"
    driver.get(url)
    return driver


def is_hisse_fr(no):
    chrome_driver = None
    try:
        url = "https://www.kap.org.tr/tr/Bildirim/" + no
        chrome_driver = load_page(url)
        # Wait for the main financial content to load
        WebDriverWait(chrome_driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "disclosureScrollableArea"))
        )
        print("Page loaded successfully!")
        full_html = chrome_driver.page_source
        soup = BeautifulSoup(full_html, 'html.parser')

        if re.search("Finansal Rapor.*", soup.text):
            stockName = soup.find('div', {"class": "lg:flex lg:flex-row flex flex-col justify-between items-center py-5"})
            stockCode = soup.find('div', {"class": "font-semibold text-base lg:text-[23px]"})
            stockCode = stockCode.text.split(",")[0]
            # bildirimTipi = soup.find_all('div', {"class": "text-danger text-base font-semibold leading-4 lg:w-auto w-1/2"})

            print("stockName: ", stockName.text)
            print("stockCode: ", stockCode)
            # print("bildirimTipi: ", bildirimTipi.text)

            for fr in soup.find_all('div', {"class": "flex flex-row justify-between text-danger font-semibold text-xl pb-9"}):
                txt = fr.find_next('div').text
                if txt == "Finansal Rapor":
                    print("Bildirim Tipi: Finansal rapor", )

                    if len(stockCode) >= 4:
                        print("Stok kodu 3 haneden büyük, hisse senedi olabilir")
                        return no, stockCode
                    else:
                        if stockCode != "OMD":  # OSMEN degil ise
                            print("Stok kodu 4 haneden küçük, hisse senedi değil. Sıradaki bildiri no'ya geçiliyor.")
                            return None, None

    finally:
        if chrome_driver is not None:
            # Close the browser window
            chrome_driver.quit()
    return None, None

print()
file.close()
kap_bildiri_no_list = []
kap_scode_list = []

for i in range(999999999):
    bildiriNo = str(int(sonBildiriNo) + 1 + i)
    link = "https://www.kap.org.tr/tr/Bildirim/" + bildiriNo
    print("link= ", link)
    print(" bildiri numarası: ", bildiriNo)
    f2 = open("/Users/myilmaz/Documents/bist/bilancolar_yeni/SonBildiriNo.txt", "r+")
    f2.write(bildiriNo)
    f2.close()
    no, scode = is_hisse_fr(bildiriNo)
    if no:
        print("Hisseye ait bir finansal rapor bulundu: ", no)
        kap_bildiri_no_list.append(no)
        kap_scode_list.append(scode)
    #time.sleep(0.51)
    print()

    print("kap_bildiri_no_list: ", kap_bildiri_no_list)
    print("kap_scode_list: ", kap_scode_list)
    # Join numbers into a single string separated by dashes
    kap_bildiri_no_string = '-'.join(str(num) for num in kap_bildiri_no_list)
    kap_scode_string = '-'.join(str(code) for code in kap_scode_list)

    with open("/Users/myilmaz/Documents/bist/bilancolar_yeni/kap_linkler.txt", "w", encoding="utf-8") as f1:
        f1.write(kap_bildiri_no_string)
    f1.close()

    with open("/Users/myilmaz/Documents/bist/bilancolar_yeni/BilancosuYeniGelenHisseler.txt", "w", encoding="utf-8") as f3:
        f3.write(kap_scode_string)
    f1.close()