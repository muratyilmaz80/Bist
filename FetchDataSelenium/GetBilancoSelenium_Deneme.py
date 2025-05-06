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


# Setup Selenium
chrome_driver_path = "//Users//myilmaz//PycharmProjects//chromedriver//chromedriver"  # <- set this correctly
service = Service(executable_path=chrome_driver_path)
options = Options()
options.add_argument("--headless")  # Headless mode
# driver = webdriver.Chrome(service=service, options=options)
# driver = webdriver.Chrome(options=options)
driver = webdriver.Chrome(executable_path=chrome_driver_path, options=options)


# Load the page
url = "https://www.kap.org.tr/tr/Bildirim/1428868"
driver.get(url)


def fnc(pg):
    soup = BeautifulSoup(pg, 'html.parser')

    #print(soup.text)
    if re.search("Finansal Rapor.*", soup.text):

        reportType = "Finansal Rapor"

        stockName = soup.find('div', {"class": "lg:flex lg:flex-row flex flex-col justify-between items-center py-5"})
        stockCode = soup.find('div', {"class": "font-semibold text-base lg:text-[23px]"})
        year = ""
        period = ""

        print(stockName.text)
        print(stockCode.text)

        for y in soup.find_all('div', {"class": "text-danger text-base font-semibold leading-4 lg:w-auto w-1/2"}):
            if y.text == "Yıl":
                year = y.find_next('div').text
                print("year: ", year)
            if y.text == "Periyot":
                period = y.find_next('div').text
                if period == "Yıllık":
                    period = "12"
                elif period == "9 Aylık":
                    period = "09"
                elif period == "6 Aylık":
                    period = "06"
                elif period == "3 Aylık":
                    period = "03"
                print("period: ", period)

        colName = year + period
        colName = int(colName)
        cols = [colName]

        # str1 = '.*general_role_.*data-input-row.*presentation-enabled'
        # str2 = '.*holding_role_.*data-input-row.*presentation-enabled'
        str3 = '.*_role_.*data-input-row.*presentation-enabled'
        trTagClass = re.compile(str3)

        labelClass = "gwt-Label multi-language-content content-tr"
        currDataClass = re.compile("taxonomy-context-value.*")

        df = pd.DataFrame(columns=cols)

        i = 0
        hitTa = 0
        hitFy = 0
        hitDa = 0
        lst = set()

        for EachPart in soup.find_all('tr', {"class": trTagClass}):
            for ep in EachPart.find_all(True, {"class": labelClass}):
                label = ep.get_text()
                label = label.strip(' \n\t')

                if label == "Ticari Alacaklar":
                    hitTa = hitTa + 1
                    if hitTa == 2:
                        label = "Ticari Alacaklar1"

                if label == "Finansal Yatırımlar":
                    hitFy = hitFy + 1
                    if hitFy == 2:
                        label = "Finansal Yatırımlar1"

                if label == "Diğer Alacaklar":
                    hitDa = hitDa + 1
                    if hitDa == 2:
                        label = "Diğer Alacaklar1"

                df.rename(index={i: label}, inplace=True)

                res = EachPart.find('td', {"class": currDataClass})
                value = res.text
                value = value.strip(' \n\t')
                if not lst.__contains__(label):
                    lst.add(label)
                    if value:
                        df.loc[label, colName] = float(value.replace('.', '').replace(',', '.'))
                    else:
                        df.loc[label, colName] = value
                    i = i + 1

        print("Hisse Adı: ", stockName.get_text())
        stockCode = stockCode.text.split(",")[0]
        # stockCode = stockCode[0]
        print("Hisse Kodu: ", stockCode)
        print("Bildirim Tipi: ", reportType)
        print("Dönem: ", colName)

        print(tabulate(df, headers=[colName]))

        print("=========================")


        fileName = "//Users//myilmaz//Documents//bist//bilancolar_deneme//" + stockCode + ".xlsx"

        if os.path.isfile(fileName):
            archiveDf = pd.read_excel(fileName, index_col=0, engine='openpyxl')
            new_df = pd.concat([archiveDf, df], axis=1, sort=False)
            new_df = new_df.replace('', 0)
            new_df.to_excel(fileName)
        else:
            df = df.replace('', 0)
            df.to_excel(fileName)


try:
    # Wait for the main financial content to load
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CLASS_NAME, "disclosureScrollableArea"))
    )
    print("Page loaded successfully!")
    full_html = driver.page_source

    fnc(full_html)

    print("Full page source saved to F:/full_page_source.txt")


finally:
    if driver is not None:
        # Close the browser window
        driver.quit()