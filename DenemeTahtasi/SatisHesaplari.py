import xlrd
from prettytable import PrettyTable

from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri
import os
import sys


hisseAdi = "SISE"
bilancoDonemi = 202209
directory = "//Users//myilmaz//Documents//bist//bilancolar"

def hesapla(varHisseAdi, varBilancoDonemi):

    print ("Hisse Adı: ", varHisseAdi)

    bilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"
    wb = xlrd.open_workbook(bilancoDosyasi)
    sheet = wb.sheet_by_index(0)

    def donemColumnFind(col):
        for columni in range(sheet.ncols):
            cell = sheet.cell(0, columni)
            if cell.value == col:
                return columni
        return -1


    def birOncekiBilancoDoneminiHesapla(dnm):
        yil = int(dnm / 100)
        ceyrek = int(dnm % 100)

        if ceyrek == 3:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + (ceyrek - 3)


    def bilancoDoneminiBul (i):
        if (i > 0):
            print ("Hatalı Bilanço Dönemi!")
            return -999;
        elif (i == 0):
            return bilancoDonemi
        else:
            a = bilancoDonemi
            while (i<0):
                a = birOncekiBilancoDoneminiHesapla(a)
                i = i + 1
            return a


    def ceyrekDegeriHesapla(r, col):
        donem = bilancoDoneminiBul(col)
        c = donemColumnFind (donem)
        quarter = (sheet.cell_value(0, c)) % (100)
        if (quarter == 3):
            return sheet.cell_value(r, c)
        else:
            if (sheet.cell_value(0,c)-sheet.cell_value(0,(c-1)) == 3):
                return (sheet.cell_value(r, c) - sheet.cell_value(r, (c - 1)))
            else:
                return -1


    def getBilancoDegeri(label, col):

        column = donemColumnFind(bilancoDonemi) + col
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == label:
                if sheet.cell_value(rowi, column)=="":
                    # print (label + " :Bilanço alanı boş!")
                    return 0
                else:
                    return sheet.cell_value(rowi, column)
        return 0

    def getBilancoTitleRow(title):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == title:
                return rowi
        return -1

    def yilliklandirmisDegerHesapla (row, bd):
        toplam = ceyrekDegeriHesapla(row, bd) + ceyrekDegeriHesapla(row, bd-1) + ceyrekDegeriHesapla(row, bd-2) + ceyrekDegeriHesapla(row, bd-3)
        return toplam



    # hisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
    # print ("Güncel Hisse Fiyatı: ", hisseFiyati)


    # BİLANÇO KALEMLERİ TANIMLAMALARI

    hasilatRow = getBilancoTitleRow("Hasılat")
    brutKarRow = getBilancoTitleRow("BRÜT KAR (ZARAR)")
    efkRow = getBilancoTitleRow("ESAS FAALİYET KARI (ZARARI)")
    donemKariRow = getBilancoTitleRow("DÖNEM KARI (ZARARI)")
    netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı")
    esasFaaliyetlerdenDigerGelirlerRow = getBilancoTitleRow("Esas Faaliyetlerden Diğer Gelirler")
    esasFaaliyetlerdenDigerGiderlerRow = getBilancoTitleRow("Esas Faaliyetlerden Diğer Giderler")
    pazarlamaGiderleriRow = getBilancoTitleRow("Pazarlama Giderleri")
    genelYonetimGiderleriRow = getBilancoTitleRow("Genel Yönetim Giderleri")
    argeGiderleriRow = getBilancoTitleRow("Araştırma ve Geliştirme Giderleri")
    amortismanlarRow = getBilancoTitleRow("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler")

    hasilat = ceyrekDegeriHesapla (hasilatRow, 0)
    print("Hasılat: ", "{:,.0f}".format(hasilat).replace(",", "."))

    brutKar = ceyrekDegeriHesapla (brutKarRow, 0)
    print("Brüt Kar: ", "{:,.0f}".format(brutKar).replace(",", "."))

    efk = ceyrekDegeriHesapla (efkRow, 0)
    print("Esas Faaliyet Karı: ", "{:,.0f}".format(efk).replace(",", "."))

    donemKari = ceyrekDegeriHesapla (donemKariRow, 0)
    print("Dönem Karı: ", "{:,.0f}".format(donemKari).replace(",", "."))

    netFaaliyetKari = efk - ceyrekDegeriHesapla (esasFaaliyetlerdenDigerGelirlerRow, 0) - ceyrekDegeriHesapla (esasFaaliyetlerdenDigerGiderlerRow, 0)
    print("Net Faaliyet Karı: ", "{:,.0f}".format(netFaaliyetKari).replace(",", "."))

    #Yıllık FAVÖK Hesabı:

    yillikBrutKar = yilliklandirmisDegerHesapla(brutKarRow, 0)

    try:
        yillikGenelYonetimGiderleri = yilliklandirmisDegerHesapla(genelYonetimGiderleriRow, 0)
    except Exception as e:
        # print("Bilançoda Yıllık Genel Yönetim Giderleri Bulunmamaktadır!")
        yillikGenelYonetimGiderleri = 0

    try:
        yillikPazarlamaGiderleri = yilliklandirmisDegerHesapla(pazarlamaGiderleriRow, 0)
    except Exception as e:
        # print("Bilançoda Pazarlama Giderleri Bulunmamaktadır!")
        yillikPazarlamaGiderleri = 0

    try:
        yillikArgeGiderleri = yilliklandirmisDegerHesapla(argeGiderleriRow, 0)
    except Exception as e:
        # print("Bilançoda AR-GE Giderleri Bulunmamaktadır!")
        yillikArgeGiderleri = 0

    try:
        yillikAmortisman = yilliklandirmisDegerHesapla(amortismanlarRow, 0)
    except Exception as e:
            # print("Bilançoda YILLIK AMORTİSMAN Bulunmamaktadır!")
            yillikAmortisman = 0

    favokYillik = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
    print ("Yıllık FAVÖK: ", "{:,.0f}".format(favokYillik).replace(",","."))



    # Çeyrek FAVÖK Hesabı:

    ceyrekBrutKar = ceyrekDegeriHesapla(brutKarRow, 0)

    try:
        genelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, 0)
    except Exception as e:
        # print("Bilançoda Genel Yönetim Giderleri Bulunmamaktadır!")
        genelYonetimGiderleri = 0

    try:
        pazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, 0)
    except Exception as e:
        # print("Bilançoda Pazarlama Giderleri Bulunmamaktadır!")
        pazarlamaGiderleri = 0

    try:
        argeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, 0)
    except Exception as e:
        # print("Bilançoda AR-GE Giderleri Bulunmamaktadır!")
        argeGiderleri = 0

    try:
        amortisman = ceyrekDegeriHesapla(amortismanlarRow, 0)
    except Exception as e:
        # print("Bilançoda Amortisman Bulunmamaktadır!")
        amortisman = 0

    favokCeyrek = ceyrekBrutKar + pazarlamaGiderleri + genelYonetimGiderleri + argeGiderleri + amortisman
    print("Çeyreklik FAVÖK: ", "{:,.0f}".format(favokCeyrek).replace(",", "."))




# ÇEYREKLİK NET KAR TABLOSU YAZDIR

    netKar0 = ceyrekDegeriHesapla (netKarRow, 0)
    netKar1 = ceyrekDegeriHesapla (netKarRow, -1)
    netKar2 = ceyrekDegeriHesapla (netKarRow, -2)
    netKar3 = ceyrekDegeriHesapla (netKarRow, -3)
    netKar4 = ceyrekDegeriHesapla (netKarRow, -4)
    netKar5 = ceyrekDegeriHesapla (netKarRow, -5)
    netKar6 = ceyrekDegeriHesapla (netKarRow, -6)
    netKar7 = ceyrekDegeriHesapla (netKarRow, -7)
    netKar8 = ceyrekDegeriHesapla(netKarRow, -8)

    netKarTablosuCeyrek = PrettyTable()
    netKarTablosuCeyrek.field_names = ["ÇEYREK", "NET KAR", "% DEĞİŞİM"]
    netKarTablosuCeyrek.align["NET KAR"] = "r"
    netKarTablosuCeyrek.align["YÜZDE DEĞİŞİM"] = "r"
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(0), "{:,.0f}".format(netKar0).replace(",", "."), "{:.2%}".format(netKar0/netKar1-1)])
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(-1), "{:,.0f}".format(netKar1).replace(",", "."), "{:.2%}".format(netKar1/netKar2-1)])
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(-2), "{:,.0f}".format(netKar2).replace(",", "."), "{:.2%}".format(netKar2/netKar3-1)])
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(-3), "{:,.0f}".format(netKar3).replace(",", "."), "{:.2%}".format(netKar3/netKar4-1)])
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(-4), "{:,.0f}".format(netKar4).replace(",", "."), "{:.2%}".format(netKar4/netKar5-1)])
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(-5), "{:,.0f}".format(netKar5).replace(",", "."), "{:.2%}".format(netKar5/netKar6-1)])
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(-6), "{:,.0f}".format(netKar6).replace(",", "."), "{:.2%}".format(netKar6/netKar7-1)])
    netKarTablosuCeyrek.add_row([bilancoDoneminiBul(-7), "{:,.0f}".format(netKar7).replace(",", "."), "{:.2%}".format(netKar7/netKar8-1)])

    print(netKarTablosuCeyrek)






# YILLIKLANDIRILMIS NET KAR TABLOSU YAZDIR

    netKarYillik0 = yilliklandirmisDegerHesapla (netKarRow, 0)
    netKarYillik1 = yilliklandirmisDegerHesapla (netKarRow, -1)
    netKarYillik2 = yilliklandirmisDegerHesapla (netKarRow, -2)
    netKarYillik3 = yilliklandirmisDegerHesapla (netKarRow, -3)
    netKarYillik4 = yilliklandirmisDegerHesapla (netKarRow, -4)
    netKarYillik5 = yilliklandirmisDegerHesapla (netKarRow, -5)
    netKarYillik6 = yilliklandirmisDegerHesapla (netKarRow, -6)
    netKarYillik7 = yilliklandirmisDegerHesapla (netKarRow, -7)
    netKarYillik8 = yilliklandirmisDegerHesapla(netKarRow, -8)

    netKarTablosuYillik = PrettyTable()
    netKarTablosuYillik.field_names = ["ÇEYREK", "YILLIK NET KAR", "% DEĞİŞİM"]
    netKarTablosuYillik.align["NET KAR"] = "r"
    netKarTablosuYillik.align["YÜZDE DEĞİŞİM"] = "r"
    netKarTablosuYillik.add_row([bilancoDoneminiBul(0), "{:,.0f}".format(netKarYillik0).replace(",", "."), "{:.2%}".format(netKarYillik0/netKarYillik1-1)])
    netKarTablosuYillik.add_row([bilancoDoneminiBul(-1), "{:,.0f}".format(netKarYillik1).replace(",", "."), "{:.2%}".format(netKarYillik1/netKarYillik2-1)])
    netKarTablosuYillik.add_row([bilancoDoneminiBul(-2), "{:,.0f}".format(netKarYillik2).replace(",", "."), "{:.2%}".format(netKarYillik2/netKarYillik3-1)])
    netKarTablosuYillik.add_row([bilancoDoneminiBul(-3), "{:,.0f}".format(netKarYillik3).replace(",", "."), "{:.2%}".format(netKarYillik3/netKarYillik4-1)])
    netKarTablosuYillik.add_row([bilancoDoneminiBul(-4), "{:,.0f}".format(netKarYillik4).replace(",", "."), "{:.2%}".format(netKarYillik4/netKarYillik5-1)])
    netKarTablosuYillik.add_row([bilancoDoneminiBul(-5), "{:,.0f}".format(netKarYillik5).replace(",", "."), "{:.2%}".format(netKarYillik5/netKarYillik6-1)])
    netKarTablosuYillik.add_row([bilancoDoneminiBul(-6), "{:,.0f}".format(netKarYillik6).replace(",", "."), "{:.2%}".format(netKarYillik6/netKarYillik7-1)])
    netKarTablosuYillik.add_row([bilancoDoneminiBul(-7), "{:,.0f}".format(netKarYillik7).replace(",", "."), "{:.2%}".format(netKarYillik7/netKarYillik8-1)])

    print(netKarTablosuYillik)










hesapla(hisseAdi, bilancoDonemi)
