import datetime
import locale
from datetime import date, time, timedelta

import gspread
import numpy as np
import pandas as pd
import regex as re
import requests
import yfinance as yf
from bs4 import BeautifulSoup
from google.oauth2.credentials import Credentials
from google.oauth2.service_account import Credentials


class FinMethods:
    @staticmethod
    def calculateSalesAndRev(dfIn, currPer, rowName):
        if (rowName in dfIn.index):
            prevPer = int(currPer) - 3
            if (prevPer % 100) == 0:
                prevPer = int(currPer) - 91
            prevYearCurrPer = int(currPer) - 100
            prevYearPrevPer = int(prevYearCurrPer) - 3
            if (prevYearPrevPer % 100) == 0:
                prevYearPrevPer = int(currPer) - 91

            if int(currPer) % 10 == 3:
                currYearSales = dfIn.at[rowName, int(currPer)]
            else:
                currYearSales = dfIn.at[rowName, int(currPer)] - dfIn.at[rowName, prevPer]

            if int(prevYearCurrPer) % 10 == 3:
                prevYearSales = dfIn.at[rowName, prevYearCurrPer]
            else:
                prevYearSales = dfIn.at[rowName, prevYearCurrPer] - dfIn.at[rowName, prevYearPrevPer]

            pctChange = FinMethods.get_change(currYearSales, prevYearSales)
            pctChange = np.round(pctChange, decimals=1)
            retDf = pd.DataFrame({'cari': currYearSales, 'yuzde': pctChange, 'oncYıl': prevYearSales}, index=[currPer])
        else:
            retDf = pd.DataFrame({'cari': 0, 'yuzde': 0, 'oncYıl': 0}, index=[currPer])
        return retDf

    @staticmethod
    def calculateData(dfIn, currPer, rowName):
        prevPer = currPer - 3
        if (prevPer % 100) == 0:
            prevPer = currPer - 91
        prevYearCurrPer = currPer - 100
        prevYearPrevPer = prevYearCurrPer - 3
        if (prevYearPrevPer % 100) == 0:
            prevYearPrevPer = currPer - 91

        if currPer % 10 == 3:
            currYearData = dfIn.at[rowName, currPer]
        else:
            currYearData = dfIn.at[rowName, currPer] - dfIn.at[rowName, prevPer]

        if prevYearCurrPer % 10 == 3:
            prevYearData = dfIn.at[rowName, prevYearCurrPer]
        else:
            prevYearData = dfIn.at[rowName, prevYearCurrPer] - dfIn.at[rowName, prevYearPrevPer]

        pctChange = FinMethods.get_change(currYearData, prevYearData)
        pctChange = np.round(pctChange, decimals=1)
        retDf = pd.DataFrame({'cari': currYearData, 'yuzde': pctChange, 'oncYıl': prevYearData}, index=[currPer])
        return retDf

    @staticmethod
    def getCariData(dfIn, currPer, rowName):
        prevPer = currPer - 3
        if (prevPer % 100) == 0:
            prevPer = currPer - 91

        if (rowName in dfIn.index):
            # if not pd.isnull(dfIn.at[ceyrekler[0]]):
            if currPer % 10 == 3:
                if not pd.isnull(dfIn.at[rowName, currPer]):
                    currYearData = dfIn.at[rowName, currPer]
                else:
                    currYearData = 0
            else:
                if not pd.isnull(dfIn.at[rowName, currPer]):  # and not pd.isnull(dfIn.at[rowName, prevPer])
                    if not pd.isnull(dfIn.at[rowName, prevPer]):
                        currYearData = dfIn.at[rowName, currPer] - dfIn.at[rowName, prevPer]
                    else:
                        currYearData = dfIn.at[rowName, currPer]
                else:
                    currYearData = 0
            return currYearData
        else:
            return 0

    @staticmethod
    def getYilliklandirilmisData(dfIn, per, rowName):
        data = 0
        for i in range(4):
            data = data + (FinMethods.getCariData(dfIn, per, rowName))
            prevPer_ = per - 3
            if (prevPer_ % 100) == 0:
                prevPer_ = per - 91
            per = prevPer_
        return data

    @staticmethod
    def getAnyCell(dfIn, currPer, rowName):
        try:
            retValue = dfIn.at[rowName, int(currPer)]
            if pd.isnull(retValue): #np.isnan(retValue):
                return 0.0
            else:
                return retValue
        except KeyError:
            return 0.0

    @staticmethod
    def divide(a, b):
        if b!=0:
            return a/b
        else:
            return 0.0

    @staticmethod
    def calculateCeyreklikFVOK(dfIn, per):

        brutKar = FinMethods.getCariData(dfIn, per, "BRÜT KAR (ZARAR)")
        # genelYonetimGiderleri = s.getCariData(dfIn, per, "Genel Yönetim Giderleri")
        # pazarlamaGiderleri = s.getCariData(dfIn, per, "Pazarlama Giderleri")

        esasNFK_ceyrek = brutKar
        genelYonetimGiderleri = FinMethods.getCariData(dfIn, per, "Genel Yönetim Giderleri")
        pazarlamaGiderleri = FinMethods.getCariData(dfIn, per, "Pazarlama Giderleri")
        argeGiderleri = FinMethods.getCariData(dfIn, per, "Araştırma ve Geliştirme Giderleri")

        istirakFaaliyetKarı = FinMethods.getCariData(dfIn, per,
                                            "Özkaynak Yöntemiyle Değerlenen Yatırımların Karlarından (Zararlarından) Paylar")

        esasNFK_ceyrek = esasNFK_ceyrek + genelYonetimGiderleri + pazarlamaGiderleri + argeGiderleri + istirakFaaliyetKarı
        # if (stockCode == "ALARK"):
        #     esasNFK_ceyrek = esasNFK_ceyrek + genelYonetimGiderleri + pazarlamaGiderleri + argeGiderleri + istirakFaaliyetKarı * 0.82  # iştiraklerden gelen karın % 82 si elektrikten gelmiş manuel düzeltilecek
        # else:
        #     esasNFK_ceyrek = esasNFK_ceyrek + genelYonetimGiderleri + pazarlamaGiderleri + argeGiderleri + istirakFaaliyetKarı

        return esasNFK_ceyrek

    @staticmethod
    def calculateYillikFVOK(dfIn, per):
        data = 0
        for i in range(4):
            data = data + (FinMethods.calculateCeyreklikFVOK(dfIn, per))
            prevPer_ = per - 3
            if (prevPer_ % 100) == 0:
                prevPer_ = per - 91
            per = prevPer_
        return data

    @staticmethod
    def get_change(current, previous):
        res = ((current - previous) / abs(previous)) * 100.0 if previous != 0 else 0
        return res

    @staticmethod
    def get_changeNAN(current, previous):
        if (previous != 0 and current != 0):
            res = ((current - previous) / abs(previous)) * 100.0
            return np.round(res, decimals=2)
        else:
            return None


    @staticmethod
    def nextPerEst(d):
        pct = max(10.0, d.iat[3, 1])
        ret = d.iat[0, 0] + abs(d.iat[0, 0] * (pct / 100.0))
        return ret

    @staticmethod
    def fPrint(p1, p2):
        print(p1, locale.format_string('%.2f', p2, True))

    @staticmethod
    def convertUsdDf(dfToConvert, usdDf, ceyrekler):
        length = len(usdDf.index)

        ortalamaOrKapanis = "Ortalama"  # Ortalama, Kapanis

        q1 = usdDf.at[ceyrekler[7], "Ortalama"]
        q2 = usdDf.at[ceyrekler[6], "Ortalama"]
        q3 = usdDf.at[ceyrekler[5], "Ortalama"]
        q4 = usdDf.at[ceyrekler[4], "Ortalama"]
        q5 = usdDf.at[ceyrekler[3], "Ortalama"]
        q6 = usdDf.at[ceyrekler[2], "Ortalama"]
        q7 = usdDf.at[ceyrekler[1], "Ortalama"]
        q8 = usdDf.at[ceyrekler[0], "Ortalama"]

        retSdf = dfToConvert

        retSdf.iat[3, 0] = FinMethods.divide(retSdf.iat[3, 0], q8)
        retSdf.iat[2, 0] = FinMethods.divide(retSdf.iat[2, 0], q7)
        retSdf.iat[1, 0] = FinMethods.divide(retSdf.iat[1, 0], q6)
        retSdf.iat[0, 0] = FinMethods.divide(retSdf.iat[0, 0], q5)
        retSdf.iat[3, 2] = FinMethods.divide(retSdf.iat[3, 2], q4)
        retSdf.iat[2, 2] = FinMethods.divide(retSdf.iat[2, 2], q3)
        retSdf.iat[1, 2] = FinMethods.divide(retSdf.iat[1, 2], q2)
        retSdf.iat[0, 2] = FinMethods.divide(retSdf.iat[0, 2], q1)
        retSdf.iat[4, 0] = retSdf.iat[0, 0] + retSdf.iat[1, 0] + retSdf.iat[2, 0] + retSdf.iat[3, 0]
        retSdf.iat[4, 2] = retSdf.iat[0, 2] + retSdf.iat[1, 2] + retSdf.iat[2, 2] + retSdf.iat[3, 2]

        retSdf.iat[0, 1] = FinMethods.get_change(retSdf.iat[0, 0], retSdf.iat[0, 2])
        retSdf.iat[1, 1] = FinMethods.get_change(retSdf.iat[1, 0], retSdf.iat[1, 2])
        retSdf.iat[2, 1] = FinMethods.get_change(retSdf.iat[2, 0], retSdf.iat[2, 2])
        retSdf.iat[3, 1] = FinMethods.get_change(retSdf.iat[3, 0], retSdf.iat[3, 2])
        retSdf.iat[4, 1] = FinMethods.get_change(retSdf.iat[4, 0], retSdf.iat[4, 2])

        return retSdf

    @staticmethod
    def passFail(d, ratio):
        if (d.iat[3, 1] >= d.iat[2, 1]) or (d.iat[2, 2] < 0 and d.iat[2, 0] > 0):
            print("Önceki dönem:", "GEÇTİ")
        else:
            print("Önceki dönem:", "KALDI")

        if (d.iat[3, 1] > ratio) or (d.iat[3, 2] < 0 and d.iat[3, 0] > 0):
            print("Cari dönem:", "GEÇTİ")
        else:
            print("Cari dönem:", "KALDI")

    @staticmethod
    def getStockPriceFromYahoo(stockCode):
        stockInfo = yf.Ticker(stockCode + ".IS").info
        print("current price: ", stockInfo['currentPrice'])
        price = stockInfo['currentPrice']

        return price

    @staticmethod
    def getStockPriceFromIsyatirim(stockCode):
        currPage = requests.get("https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/default.aspx")
        soup = BeautifulSoup(currPage.text, 'html.parser')
        for part in soup.find_all('div', {"class": "table-wrap-data"}):
            for p in part.findChildren('a'):
                stockCodeCurrent = p.text.strip(' \r\n\t')
                if stockCodeCurrent != stockCode:
                    continue
                parentTag = p.parent
                next_td_tag = parentTag.find_next('td')  # ,{"class=": "text-right"}
                stockPrice = next_td_tag.text
                currPriceString = str(stockPrice)
                currPriceString = re.sub(r'[.]{1,}', '', currPriceString)
                price = currPriceString.replace(',', '.')
                price = float(price)

        return price

    @staticmethod
    def getStockPricesFromIsyatirim():
        currPage = requests.get("https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/default.aspx")
        soup = BeautifulSoup(currPage.text, 'html.parser')
        stockPricesDict = {}
        for part in soup.find_all('div', {"class": "table-wrap-data"}):
            for p in part.findChildren('a'):
                stockCodeCurrent = p.text.strip(' \r\n\t')
                if len(stockCodeCurrent) < 4:
                    continue
                parentTag = p.parent
                next_td_tag = parentTag.find_next('td')  # ,{"class=": "text-right"}
                stockPrice = next_td_tag.text
                currPriceString = str(stockPrice)
                currPriceString = re.sub(r'[.]{1,}', '', currPriceString)
                price = currPriceString.replace(',', '.')
                # print("price:", price)
                price = float(price)
                stockPricesDict[stockCodeCurrent] = price

        return stockPricesDict

    @staticmethod
    def findPeriods(cPer):
        cPer = int(cPer)
        qRes = [str(cPer)]
        for i in range(3):
            pPer = cPer - 3
            if (pPer % 100) == 0:
                pPer = cPer - 91
            qRes.append(str(pPer))
            cPer = pPer
        return qRes

    @staticmethod
    def getYield():
        p = requests.get("https://www.bloomberght.com/tahvil/faiz")
        soup = BeautifulSoup(p.text, 'html.parser')
        tagCls = re.compile('(upGreen)|(downRed)|(round)')
        for part in soup.find_all('span', {"class": tagCls}):
            val = part.text
            val = val.strip(' \n\t')
            val = float(val.replace('.', '').replace(',', '.'))
            return val

    @staticmethod
    def divisionWOException(n, d):
        return n / d if d else 0

    @staticmethod
    def writeToGoogleSheets(stockCode, gsData):

        # If modifying these scopes, delete the file token.json.
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive']

        creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
        client = gspread.authorize(creds)
        sheet = client.open("Test").sheet1

        gsData["excel_hedefAlim_Ortalama"] = (gsData["excel_hedefAlimEFK"] + gsData[
            "excel_hedefAlimNFK"]) / 2  # gsData[]
        gsData["excel_hedefFiyatNFK"] = np.round(gsData["excel_hedefAlimNFK"] * 1.515151, decimals=2)
        gsData["excel_guncelHAU_NFK"] = ((gsData["excel_hedefAlimNFK"] / gsData["excel_guncelFiyat"]) - 1)
        gsData["excel_guncelHAU_EFK"] = ((gsData["excel_hedefAlimEFK"] / gsData["excel_guncelFiyat"]) - 1)
        gsData["excel_guncelHAU_Ortalama"] = (gsData["excel_guncelHAU_NFK"] + gsData["excel_guncelHAU_EFK"]) / 2
        stockInfoRow = [stockCode, gsData["excel_hedefAlim_Ortalama"], gsData["excel_hedefAlimNFK"],
                        gsData["excel_hedefAlimEFK"], gsData["excel_hedefFiyatNFK"], gsData["excel_guncelFiyat"],
                        gsData["excel_netPro"], gsData["excel_guncelHAU_Ortalama"], gsData["excel_guncelHAU_NFK"],
                        gsData["excel_guncelHAU_EFK"], gsData["excel_nakitAkis"]]
        #print("googleSheetsData: ",stockInfoRow)
        # pp(stockInfoRow)
        stockNames = sheet.col_values(1)
        # pp(stockNames)
        stockNamesLength = len(stockNames)
        stockIndex = -1
        for i in range(stockNamesLength):
            if (stockCode == stockNames[i]):
                stockIndex = i
        if (stockIndex >= 0):
            print(stockCode, " Google sheets listesinde zaten var, yeniden eklenmedi")
        else:
            sheet.insert_row(stockInfoRow, stockNamesLength + 1)
            print(stockCode, " Google sheets listesine eklendi")
        print()

    @staticmethod
    def getPeriods(cPer):
        cPer = cPer
        qRes = [cPer]
        for i in range(30):
            pPer = cPer - 3
            if (pPer % 100) == 0:
                pPer = cPer - 91
            qRes.append(pPer)
            cPer = pPer
        return qRes

    @staticmethod
    def nth_root(num, root):
        answer = num ** (1 / root)
        return answer

    @staticmethod
    def getOrtalamaArtis(df, ceyrekler, forYear, label):
        totalYillikDataArtisi = 0
        for i in range(forYear):
            currYearData = FinMethods.getYilliklandirilmisData(df, ceyrekler[i * 4], label)
            prevYearData = FinMethods.getYilliklandirilmisData(df, ceyrekler[(i + 1) * 4], label)
            yillikDataDegisim = FinMethods.get_change(currYearData, prevYearData)
            totalYillikDataArtisi = totalYillikDataArtisi + yillikDataDegisim

        ortalamaYillikDataArtisi = totalYillikDataArtisi / forYear
        return ortalamaYillikDataArtisi

    @staticmethod
    def getOrtalamaNFKArtis(df, ceyrekler, forYear):
        totalYillikNFKArtisi = 0
        for i in range(forYear):
            currYearData = FinMethods.calculateYillikFVOK(df, ceyrekler[i * 4])
            prevYearData = FinMethods.calculateYillikFVOK(df, ceyrekler[(i + 1) * 4])
            yillikNFKDegisim = FinMethods.get_change(currYearData, prevYearData)
            totalYillikNFKArtisi = totalYillikNFKArtisi + yillikNFKDegisim

        ortalamaYillikNFKArtisi = totalYillikNFKArtisi / forYear
        return ortalamaYillikNFKArtisi

    @staticmethod
    def getNetIsletmeSermayesi(df, per):
        global kısaDigerBorclar, netIsletmeSermayesi
        kısaDigerBorclar = FinMethods.getAnyCell(df, per, "Kiralama İşlemlerinden Borçlar") \
                           + FinMethods.getAnyCell(df, per, "Kısa Çalışanlara Sağlanan Faydalar Kapsamında Borçlar") \
                           + FinMethods.getAnyCell(df, per, "Kısa Diğer Borçlar") \
                           + FinMethods.getAnyCell(df, per,
                                          "Kısa Ertelenmiş Gelirler (Müşteri Sözleşmelerinden Doğan Yükümlülüklerin Dışında Kalanlar)") \
                           + FinMethods.getAnyCell(df, per, "Dönem Karı Vergi Yükümlülüğü") \
                           + FinMethods.getAnyCell(df, per, "Kısa Vadeli Karşılıklar")
        netIsletmeSermayesi = FinMethods.getAnyCell(df, per, "Dönen Ticari Alacaklar") \
                              + FinMethods.getAnyCell(df, per, "Dönen Diğer Alacaklar") \
                              + FinMethods.getAnyCell(df, per, "Dönen Peşin Ödenmiş Giderler") \
                              + FinMethods.getAnyCell(df, per, "Stoklar") \
                              - FinMethods.getAnyCell(df, per, "Kısa Ticari Borçlar") \
                              - kısaDigerBorclar
        # print("kısaDigerBorclar:", kısaDigerBorclar)
        return netIsletmeSermayesi

    @staticmethod
    def getROE(df, ceyrekler):
        yillikROE = FinMethods.divide(FinMethods.getYilliklandirilmisData(df, ceyrekler[0], "DÖNEM KARI (ZARARI)"),
                                      ((FinMethods.getAnyCell(df, ceyrekler[0],
                                                                        "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                               df,
                                               ceyrekler[4],
                                               "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        yillikROE_prev = FinMethods.divide(FinMethods.getYilliklandirilmisData(df, ceyrekler[1], "DÖNEM KARI (ZARARI)"),
                                           ((FinMethods.getAnyCell(df, ceyrekler[1],
                                                                             "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                    df,
                                                    ceyrekler[
                                                        5],
                                                    "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        yillikROE_prev2 = FinMethods.divide(FinMethods.getYilliklandirilmisData(df, ceyrekler[2], "DÖNEM KARI (ZARARI)"),
                                            ((FinMethods.getAnyCell(df, ceyrekler[2],
                                                                              "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                     df,
                                                     ceyrekler[6], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        yillikROE_prev3 = FinMethods.divide(FinMethods.getYilliklandirilmisData(df, ceyrekler[3], "DÖNEM KARI (ZARARI)"),
                                            ((FinMethods.getAnyCell(df, ceyrekler[3],
                                                                              "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                     df,
                                                     ceyrekler[7], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        yillikROE_prev4 = FinMethods.divide(FinMethods.getYilliklandirilmisData(df, ceyrekler[4], "DÖNEM KARI (ZARARI)"),
                                            ((FinMethods.getAnyCell(df, ceyrekler[4],
                                                                              "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                     df,
                                                     ceyrekler[8], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        roeDf = pd.DataFrame(
            {'Ö.S.K': [yillikROE_prev4, yillikROE_prev3, yillikROE_prev2, yillikROE_prev, yillikROE]},
            index=[ceyrekler[4], ceyrekler[3], ceyrekler[2], ceyrekler[1],
                   ceyrekler[0]])
        return roeDf

    @staticmethod
    def getROEFromNFK(df, ceyrekler):
        modified_yillikROE = FinMethods.divide(FinMethods.calculateYillikFVOK(df, ceyrekler[0]),
                                               ((FinMethods.getAnyCell(df, ceyrekler[0],
                                                                                 "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                        df,
                                                        ceyrekler[4], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        modified_yillikROE_prev = FinMethods.divide(FinMethods.calculateYillikFVOK(df, ceyrekler[1]),
                                                    ((FinMethods.getAnyCell(df, ceyrekler[1],
                                                                                      "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                             df,
                                                             ceyrekler[5], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        modified_yillikROE_prev2 = FinMethods.divide(FinMethods.calculateYillikFVOK(df, ceyrekler[2]),
                                                     ((FinMethods.getAnyCell(df, ceyrekler[2],
                                                                                       "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                              df,
                                                              ceyrekler[6], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        modified_yillikROE_prev3 = FinMethods.divide(FinMethods.calculateYillikFVOK(df, ceyrekler[3]),
                                                     ((FinMethods.getAnyCell(df, ceyrekler[3],
                                                                                       "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                              df,
                                                              ceyrekler[7], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        modified_yillikROE_prev4 = FinMethods.divide(FinMethods.calculateYillikFVOK(df, ceyrekler[4]),
                                                     ((FinMethods.getAnyCell(df, ceyrekler[4],
                                                                                       "TOPLAM ÖZKAYNAKLAR") + FinMethods.getAnyCell(
                                                              df,
                                                              ceyrekler[8], "TOPLAM ÖZKAYNAKLAR")) / 2)) * 100
        modifiedROEDf = pd.DataFrame(
            {'Ö.S.K(FROM NET FAALİYET KARI)': [modified_yillikROE_prev4, modified_yillikROE_prev3,
                                                             modified_yillikROE_prev2, modified_yillikROE_prev,
                                                             modified_yillikROE]},
            index=[ceyrekler[4], ceyrekler[3], ceyrekler[2], ceyrekler[1], ceyrekler[0]])
        return modifiedROEDf

    @staticmethod
    def dateToString(d):
        try:
            year = d.year
            month = d.month
            day = d.day
            if month < 10:
                month = "0" + str(month)
            if day < 10:
                day = "0" + str(day)
            res = str(day) + "." + str(month) + "." + str(year)
            return res
        except AttributeError:
            print("TARİH BİLGİSİ YOK")
            return "01.01.1970"

    @staticmethod
    def getYieldByDate(tDate, f):
        yieldDf = pd.read_csv(f, index_col=0)

        yCell = ''
        while yCell == '':
            try:
                yCell = yieldDf.at[tDate, "Şimdi"]
            except KeyError:
                nextDay=FinMethods.getNextDay(tDate)
                #print("nextDay: ",nextDay)
                tDate = FinMethods.dateToString(nextDay)
        yCell= yCell.replace(',', '.')
        return yCell

    @staticmethod
    def getclosestPrevQuarterDate(dt):
        currDateArray = dt.split('-')
        currDate = date(int(currDateArray[2]), int(currDateArray[1]), int(currDateArray[0]))

        currentYear = int(currDateArray[2])
        lastYear= currentYear-1

        q1End = date(currentYear, 3, 31)
        q2End = date(currentYear, 6, 30)
        q3End = date(currentYear, 9, 30)
        q4End = date(currentYear, 12, 31)

        if currDate <= q1End:
            return date(lastYear, 12, 31).strftime('%d-%m-%Y')
        elif q1End < currDate <= q2End:
            return q1End.strftime('%d-%m-%Y')
        elif q2End < currDate <= q3End:
            return q2End.strftime('%d-%m-%Y')
        elif q3End < currDate <= q4End:
            return q3End.strftime('%d-%m-%Y')
        else:
            print("closestPrevQuarterDate is None!!!")
            return None

    @staticmethod
    def getclosestPrevQuarterDate_(year, period):
        print("yearr: ",year)
        intYear= int(year.strip())
        q1End = date(intYear, 3, 31)
        q2End = date(intYear, 6, 30)
        q3End = date(intYear, 9, 30)
        q4End = date(intYear, 12, 31)

        if period == "03":
            return q1End.strftime('%d-%m-%Y')
        elif period == "06":
            return q2End.strftime('%d-%m-%Y')
        elif period == "09":
            return q3End.strftime('%d-%m-%Y')
        elif period == "12":
            return q4End.strftime('%d-%m-%Y')
        else:
            print("closestPrevQuarterDate is None!!!")
            return None

    @staticmethod
    def getNextDay(sDate):
        sDate = sDate.split('.')
        d = date(int(sDate[2]), int(sDate[1]), int(sDate[0]))
        d = d + timedelta(days=1)
        return d

    @staticmethod
    def getClosestPrevWeekday(date, hour):
        myDate = datetime.datetime.strptime(date, "%d-%m-%Y")
        if not FinMethods.isTimeBeforeMidnight(hour):
            print("bilanço 18-24 arası gonderilmemiş, düzeltiliyor")
            myDate= myDate - datetime.timedelta(days=1)

        # Get Day Number from weekday
        weekno = myDate.weekday()
        returnDate=""
        if weekno < 5:
            returnDate= myDate
        elif weekno==5:
            returnDate= myDate - datetime.timedelta(days=1)
        elif weekno==6:
            returnDate= myDate - datetime.timedelta(days=2)
        else:
            print("Weekday hatalı dönüyor")

        return returnDate.strftime('%d-%m-%Y')

    @staticmethod
    def isTimeBeforeMidnight(hour):
        hour = hour.split(':')
        time_ = time(int(hour[0]), int(hour[1]), int(hour[2]))
        begin_time = time(18, 00, 00)
        end_time = time(23, 59, 59)
        if begin_time < end_time:
            return time_ >= begin_time and time_ <= end_time
        else:  # crosses midnight
            return time_ >= begin_time or time_ <= end_time

    @staticmethod
    def getPrevDayString(date):
        myDate = datetime.datetime.strptime(date, "%d-%m-%Y")
        myDate = myDate - datetime.timedelta(days=1)
        return myDate.strftime('%d-%m-%Y')

