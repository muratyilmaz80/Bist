from bs4 import BeautifulSoup
import requests
import regex as re
import pandas as pd
from tabulate import tabulate
import os.path


def fnc(pg):
    currPage = requests.get(pg)
    soup = BeautifulSoup(currPage.text, 'html.parser')

    for part in soup.find_all('h1'):
        if re.search("Finansal Rapor.*", part.text):
            reportType = "Finansal Rapor"

            stockName = soup.find('div', {"class": "type-medium type-bold bi-sky-black"})
            stockCode = soup.find('div', {"class": "type-medium bi-dim-gray"})
            year = ""
            period = ""

            for p in soup.find_all('div', {"class": "w-col w-col-3 modal-briefsumcol"}):
                for y in p.find_all('div', {"type-small bi-lightgray"}):
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
            #stockCode = stockCode[0]
            print("Hisse Kodu: ", stockCode)
            print("Bildirim Tipi: ", reportType)
            print("Dönem: ", colName)

            print(tabulate(df, headers=[colName]))

            print("=========================")

            fileName = "D:\\bist\\bilancolar\\" + stockCode + ".xlsx"

            if os.path.isfile(fileName):
                archiveDf = pd.read_excel(fileName, index_col=0)
                new_df = pd.concat([archiveDf, df], axis=1, sort=False)
                new_df = new_df.replace('', 0)
                new_df.to_excel(fileName)
            else:
                df = df.replace('', 0)
                df.to_excel(fileName)


def fncMultiple(s):
    content = s
    content = content.strip()
    contentList = content.split("-")
    contentList = ["https://www.kap.org.tr/tr/Bildirim/" + x for x in contentList]
    print(contentList)

    for link in contentList:
        fnc(link)

#fnc("https://www.kap.org.tr/tr/Bildirim/846388")

fncMultiple("689119-703362-720128-741273-768198-782307-798831-820286-851321")