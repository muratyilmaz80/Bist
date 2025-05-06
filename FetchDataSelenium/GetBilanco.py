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


def fnc(pg):
    global sunum_pb_int
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

        for spb in soup.find('td', {"class": "financial-header-title"}):
            if spb == 'Sunum Para Birimi':
                sunum_pb = spb.find_next('td').text
                sunum_pb = sunum_pb.replace(".", "")  # Remove dots used as thousand separators
                sunum_pb = sunum_pb.replace("TL", "")  # Remove currency symbol
                sunum_pb = sunum_pb.strip()  # Remove any extra spaces

                # Convert to integer
                if sunum_pb == '':
                    sunum_pb_int = 1
                else:
                    sunum_pb_int = int(sunum_pb)

                print("sunum_pb_int: ", sunum_pb_int)

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
        lst = set()

        varliklarUstLabel = ""
        yukumluluklerUstLabel = ""

        for EachPart in soup.find_all('tr', {"class": trTagClass}):
            for ep in EachPart.find_all(True, {"class": labelClass}):
                label = ep.get_text()
                label = label.strip(' \n\t')

                if label == "TOPLAM DÖNEN VARLIKLAR":
                    varliklarUstLabel = "TOPLAM DÖNEN VARLIKLAR";

                elif label == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                    yukumluluklerUstLabel = "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER";

                # Varliklar
                elif label == "Ticari Alacaklar":
                    if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                        label = "Duran Ticari Alacaklar"
                    else:
                        label = "Dönen Ticari Alacaklar"

                elif label == "Finansal Yatırımlar":
                    if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                        label = "Duran Finansal Yatırımlar"
                    else:
                        label = "Dönen Finansal Yatırımlar"

                elif label == "Diğer Alacaklar":
                    if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                        label = "Duran Diğer Alacaklar"
                    else:
                        label = "Dönen Diğer Alacaklar"

                elif label == "Peşin Ödenmiş Giderler":
                    if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                        label = "Duran Peşin Ödenmiş Giderler"
                    else:
                        label = "Dönen Peşin Ödenmiş Giderler"

                elif label == "Türev Araçlar":
                    if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                        label = "Duran Türev Araçlar"
                    else:
                        label = "Dönen Türev Araçlar"

                elif label == "Finans Sektörü Faaliyetlerinden Alacaklar":
                    if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                        label = "Duran Finans Sektörü Faaliyetlerinden Alacaklar"
                    else:
                        label = "Dönen Finans Sektörü Faaliyetlerinden Alacaklar"

                # Yukumlulukler
                elif label == "Diğer Finansal Yükümlülükler":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Diğer Finansal Yükümlülükler1"
                    else:
                        label = "Kısa Diğer Finansal Yükümlülükler1"

                elif label == "Ertelenmiş Gelirler":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Ertelenmiş Gelirler"
                    else:
                        label = "Kısa Ertelenmiş Gelirler"

                elif label == "Diğer Borçlar":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Diğer Borçlar"
                    else:
                        label = "Kısa Diğer Borçlar"

                elif label == "Ticari Borçlar":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Ticari Borçlar"
                    else:
                        label = "Kısa Ticari Borçlar"

                elif label == "Türev Araçlar":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Türev Araçlar"
                    else:
                        label = "Kısa Türev Araçlar"

                elif label == "Finans Sektörü Faaliyetlerinden Borçlar":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Finans Sektörü Faaliyetlerinden Borçlar"
                    else:
                        label = "Kısa Finans Sektörü Faaliyetlerinden Borçlar"

                elif label == "Müşteri Sözleşmelerinden Doğan Yükümlülükler":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Müşteri Sözleşmelerinden Doğan Yükümlülükler"
                    else:
                        label = "Kısa Müşteri Sözleşmelerinden Doğan Yükümlülükler"

                elif label == "Ertelenmiş Gelirler (Müşteri Sözleşmelerinden Doğan Yükümlülüklerin Dışında Kalanlar)":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Ertelenmiş Gelirler (Müşteri Sözleşmelerinden Doğan Yükümlülüklerin Dışında Kalanlar)"
                    else:
                        label = "Kısa Ertelenmiş Gelirler (Müşteri Sözleşmelerinden Doğan Yükümlülüklerin Dışında Kalanlar)"
                elif label == "Çalışanlara Sağlanan Faydalar Kapsamında Borçlar":
                    if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        label = "Uzun Çalışanlara Sağlanan Faydalar Kapsamında Borçlar"
                    else:
                        label = "Kısa Çalışanlara Sağlanan Faydalar Kapsamında Borçlar"
                elif label == "Hasılat":
                    isHasılatLabelFound = True

                df.rename(index={i: label}, inplace=True)

                res = EachPart.find('td', {"class": currDataClass})
                value = res.text
                value = value.strip(' \n\t')
                if not lst.__contains__(label):
                    lst.add(label)
                    if value:
                        df.loc[label, colName] = float(value.replace('.', '').replace(',', '.')) * sunum_pb_int
                    else:
                        df.loc[label, colName] = value * sunum_pb_int
                    i = i + 1

        print("Hisse Adı: ", stockName.get_text())
        stockCode = stockCode.text.split(",")[0]
        # stockCode = stockCode[0]
        print("Hisse Kodu: ", stockCode)
        print("Bildirim Tipi: ", reportType)
        print("Dönem: ", colName)

        print(tabulate(df, headers=[colName]))

        print("=========================")

        fileName = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + stockCode + ".xlsx"

        if os.path.isfile(fileName):
            archiveDf = pd.read_excel(fileName, index_col=0, engine='openpyxl')
            new_df = pd.concat([archiveDf, df], axis=1, sort=False)
            new_df = new_df.replace('', 0)
            new_df.to_excel(fileName)
        else:
            df = df.replace('', 0)
            df.to_excel(fileName)


def get_content_list(s):
    content = s
    content = content.strip()
    contentList = content.split("-")
    contentList = ["https://www.kap.org.tr/tr/Bildirim/" + x for x in contentList]
    return contentList


chrome_driver = None
try:
    with open("/Users/myilmaz/Documents/bist/bilancolar_yeni/kap_linkler.txt", "r", encoding="utf-8") as f:
        content = f.read()
    print(get_content_list(content))
    for url in get_content_list(content):
        chrome_driver = load_page(url)
        # Wait for the main financial content to load
        WebDriverWait(chrome_driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "disclosureScrollableArea"))
        )
        print("Page loaded successfully!")
        full_html = chrome_driver.page_source

        fnc(full_html)


finally:
    if chrome_driver is not None:
        # Close the browser window
        chrome_driver.quit()
