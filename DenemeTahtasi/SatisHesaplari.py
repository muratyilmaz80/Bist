import xlrd
from prettytable import PrettyTable

from GetGuncelHisseDegeri import returnGuncelHisseDegeri
import os
import sys


hisseAdi = "ULUUN"
bilancoDonemi = 202209
directory = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar"

def hesapla(varHisseAdi, varBilancoDonemi):

    print ("Hisse Adı: ", varHisseAdi)

    bilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + varHisseAdi + ".xlsx"
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



    #hisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
    #print ("Güncel Hisse Fiyatı: ", hisseFiyati)


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

    def yillikFavokHesapla (sonCeyrek):

        yillikBrutKar = yilliklandirmisDegerHesapla(brutKarRow, sonCeyrek)

        if (genelYonetimGiderleriRow != -1):
            yillikGenelYonetimGiderleri = yilliklandirmisDegerHesapla(genelYonetimGiderleriRow, sonCeyrek)
        else:
            yillikGenelYonetimGiderleri = 0

        if (pazarlamaGiderleriRow != -1):
            yillikPazarlamaGiderleri = yilliklandirmisDegerHesapla(pazarlamaGiderleriRow, sonCeyrek)
        else:
            yillikPazarlamaGiderleri = 0

        if (argeGiderleriRow != -1):
            try:
                yillikArgeGiderleri = yilliklandirmisDegerHesapla(argeGiderleriRow, sonCeyrek)
            except Exception as e:
                yillikArgeGiderleri = 0
        else:
            yillikArgeGiderleri = 0



        if (amortismanlarRow != -1):
            yillikAmortisman = yilliklandirmisDegerHesapla(amortismanlarRow, sonCeyrek)
        else:
            yillikAmortisman = 0

        favokYillik = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
        return favokYillik


    print("Yıllık FAVÖK: ", "{:,.0f}".format(yillikFavokHesapla(0)).replace(",", "."))



    def ceyrekFavokHesapla(ceyrek):

        ceyrekBrutKar = ceyrekDegeriHesapla(brutKarRow, ceyrek)

        if (genelYonetimGiderleriRow != -1):
            genelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, ceyrek)
        else:
            genelYonetimGiderleri = 0

        if (pazarlamaGiderleriRow != -1):
            pazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, ceyrek)
        else:
            pazarlamaGiderleri = 0

        if (argeGiderleriRow != -1):
            try:
                argeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, ceyrek)
            except Exception as e:
                argeGiderleri = 0
        else:
            argeGiderleri = 0


        if (amortismanlarRow != -1):
            amortisman = ceyrekDegeriHesapla(amortismanlarRow, ceyrek)
        else:
            amortisman = 0

        favokCeyrek = ceyrekBrutKar + pazarlamaGiderleri + genelYonetimGiderleri + argeGiderleri + amortisman
        return favokCeyrek

    print("Çeyreklik FAVÖK: ", "{:,.0f}".format(ceyrekFavokHesapla(0)).replace(",", "."))



# Istiraklerden Gelen Net Kar Kontrolu
    try:
        istiraklerdenGelenKarRow = getBilancoTitleRow("Özkaynak Yöntemiyle Değerlenen Yatırımların Karlarından (Zararlarından) Paylar")
        istiraklerdenGelenNetKarSonCeyrek = ceyrekDegeriHesapla(istiraklerdenGelenKarRow, 0)
        print("Özkaynak Yöntemiyle Değerlenen Yatırım Karları: ","{:,.0f}".format(istiraklerdenGelenNetKarSonCeyrek).replace(",", "."))
        print()
    except Exception as e:
        print("Bilançoda Özkaynak Yöntemiyle Değerlenen Yatırımların Karlarından (Zararlarından) Paylar Bulunmamaktadır!")



# HASILAT NETKAR NETKARMARJI HESAPLARI

    hasilat0 = ceyrekDegeriHesapla (hasilatRow, 0)
    hasilat1 = ceyrekDegeriHesapla (hasilatRow, -1)
    hasilat2 = ceyrekDegeriHesapla (hasilatRow, -2)
    hasilat3 = ceyrekDegeriHesapla (hasilatRow, -3)
    hasilat4 = ceyrekDegeriHesapla (hasilatRow, -4)
    hasilat5 = ceyrekDegeriHesapla (hasilatRow, -5)
    hasilat6 = ceyrekDegeriHesapla (hasilatRow, -6)
    hasilat7 = ceyrekDegeriHesapla (hasilatRow, -7)
    hasilat8 = ceyrekDegeriHesapla (hasilatRow, -8)

    netKar0 = ceyrekDegeriHesapla (netKarRow, 0)
    netKar1 = ceyrekDegeriHesapla (netKarRow, -1)
    netKar2 = ceyrekDegeriHesapla (netKarRow, -2)
    netKar3 = ceyrekDegeriHesapla (netKarRow, -3)
    netKar4 = ceyrekDegeriHesapla (netKarRow, -4)
    netKar5 = ceyrekDegeriHesapla (netKarRow, -5)
    netKar6 = ceyrekDegeriHesapla (netKarRow, -6)
    netKar7 = ceyrekDegeriHesapla (netKarRow, -7)
    netKar8 = ceyrekDegeriHesapla (netKarRow, -8)

    netKarMarji0 = netKar0 / hasilat0
    netKarMarji1 = netKar1 / hasilat1
    netKarMarji2 = netKar2 / hasilat2
    netKarMarji3 = netKar3 / hasilat3
    netKarMarji4 = netKar4 / hasilat4
    netKarMarji5 = netKar5 / hasilat5
    netKarMarji6 = netKar6 / hasilat6
    netKarMarji7 = netKar7 / hasilat7




    # ÇEYREKLİK HASILAT TABLOSU YAZDIR

    hasilatTablosuCeyrek = PrettyTable()
    hasilatTablosuCeyrek.field_names = ["ÇEYREK", "HASILAT", "% DEĞİŞİM"]
    hasilatTablosuCeyrek.align["HASILAT"] = "r"
    hasilatTablosuCeyrek.align["YÜZDE DEĞİŞİM"] = "r"
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(0), "{:,.0f}".format(hasilat0).replace(",", "."), "{:.2%}".format(hasilat0/hasilat1-1)])
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(-1), "{:,.0f}".format(hasilat1).replace(",", "."), "{:.2%}".format(hasilat1/hasilat2-1)])
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(-2), "{:,.0f}".format(hasilat2).replace(",", "."), "{:.2%}".format(hasilat2/hasilat3-1)])
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(-3), "{:,.0f}".format(hasilat3).replace(",", "."), "{:.2%}".format(hasilat3/hasilat4-1)])
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(-4), "{:,.0f}".format(hasilat4).replace(",", "."), "{:.2%}".format(hasilat4/hasilat5-1)])
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(-5), "{:,.0f}".format(hasilat5).replace(",", "."), "{:.2%}".format(hasilat5/hasilat6-1)])
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(-6), "{:,.0f}".format(hasilat6).replace(",", "."), "{:.2%}".format(hasilat6/hasilat7-1)])
    hasilatTablosuCeyrek.add_row([bilancoDoneminiBul(-7), "{:,.0f}".format(hasilat7).replace(",", "."), "{:.2%}".format(hasilat7/hasilat8-1)])

    print(hasilatTablosuCeyrek)


    # YILLIKLANDIRILMIS HASILAT TABLOSU YAZDIR

    hasilatYillik0 = yilliklandirmisDegerHesapla(hasilatRow, 0)
    hasilatYillik1 = yilliklandirmisDegerHesapla(hasilatRow, -1)
    hasilatYillik2 = yilliklandirmisDegerHesapla(hasilatRow, -2)
    hasilatYillik3 = yilliklandirmisDegerHesapla(hasilatRow, -3)
    hasilatYillik4 = yilliklandirmisDegerHesapla(hasilatRow, -4)
    hasilatYillik5 = yilliklandirmisDegerHesapla(hasilatRow, -5)
    hasilatYillik6 = yilliklandirmisDegerHesapla(hasilatRow, -6)
    hasilatYillik7 = yilliklandirmisDegerHesapla(hasilatRow, -7)
    hasilatYillik8 = yilliklandirmisDegerHesapla(hasilatRow, -8)


    hasilatTablosuYillik = PrettyTable()
    hasilatTablosuYillik.field_names = ["ÇEYREK", "YILLIK HASILAT", "% DEĞİŞİM"]
    hasilatTablosuYillik.align["YILLIK HASILAT"] = "r"
    hasilatTablosuYillik.align["YÜZDE DEĞİŞİM"] = "r"
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(0), "{:,.0f}".format(hasilatYillik0).replace(",", "."), "{:.2%}".format(hasilatYillik0/hasilatYillik1-1)])
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(-1), "{:,.0f}".format(hasilatYillik1).replace(",", "."), "{:.2%}".format(hasilatYillik1/hasilatYillik2-1)])
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(-2), "{:,.0f}".format(hasilatYillik2).replace(",", "."), "{:.2%}".format(hasilatYillik2/hasilatYillik3-1)])
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(-3), "{:,.0f}".format(hasilatYillik3).replace(",", "."), "{:.2%}".format(hasilatYillik3/hasilatYillik4-1)])
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(-4), "{:,.0f}".format(hasilatYillik4).replace(",", "."), "{:.2%}".format(hasilatYillik4/hasilatYillik5-1)])
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(-5), "{:,.0f}".format(hasilatYillik5).replace(",", "."), "{:.2%}".format(hasilatYillik5/hasilatYillik6-1)])
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(-6), "{:,.0f}".format(hasilatYillik6).replace(",", "."), "{:.2%}".format(hasilatYillik6/hasilatYillik7-1)])
    hasilatTablosuYillik.add_row([bilancoDoneminiBul(-7), "{:,.0f}".format(hasilatYillik7).replace(",", "."), "{:.2%}".format(hasilatYillik7/hasilatYillik8-1)])

    print(hasilatTablosuYillik)



















# ---------------------------------


# ÇEYREKLİK EFK TABLOSU YAZDIR

    efk0 = ceyrekDegeriHesapla(efkRow, 0)
    efk1 = ceyrekDegeriHesapla(efkRow, -1)
    efk2 = ceyrekDegeriHesapla(efkRow, -2)
    efk3 = ceyrekDegeriHesapla(efkRow, -3)
    efk4 = ceyrekDegeriHesapla(efkRow, -4)
    efk5 = ceyrekDegeriHesapla(efkRow, -5)
    efk6 = ceyrekDegeriHesapla(efkRow, -6)
    efk7 = ceyrekDegeriHesapla(efkRow, -7)
    efk8 = ceyrekDegeriHesapla(efkRow, -8)

    efkTablosuCeyrek = PrettyTable()
    efkTablosuCeyrek.field_names = ["ÇEYREK", "EFK", "% DEĞİŞİM"]
    efkTablosuCeyrek.align["EFK"] = "r"
    efkTablosuCeyrek.align["YÜZDE DEĞİŞİM"] = "r"
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(0), "{:,.0f}".format(efk0).replace(",", "."), "{:.2%}".format(efk0/efk1-1)])
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(-1), "{:,.0f}".format(efk1).replace(",", "."), "{:.2%}".format(efk1/efk2-1)])
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(-2), "{:,.0f}".format(efk2).replace(",", "."), "{:.2%}".format(efk2/efk3-1)])
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(-3), "{:,.0f}".format(efk3).replace(",", "."), "{:.2%}".format(efk3/efk4-1)])
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(-4), "{:,.0f}".format(efk4).replace(",", "."), "{:.2%}".format(efk4/efk5-1)])
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(-5), "{:,.0f}".format(efk5).replace(",", "."), "{:.2%}".format(efk5/efk6-1)])
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(-6), "{:,.0f}".format(efk6).replace(",", "."), "{:.2%}".format(efk6/efk7-1)])
    efkTablosuCeyrek.add_row([bilancoDoneminiBul(-7), "{:,.0f}".format(efk7).replace(",", "."), "{:.2%}".format(efk7/efk8-1)])

    print(efkTablosuCeyrek)


    # YILLIKLANDIRILMIS EFK TABLOSU YAZDIR

    efkYillik0 = yilliklandirmisDegerHesapla(efkRow, 0)
    efkYillik1 = yilliklandirmisDegerHesapla(efkRow, -1)
    efkYillik2 = yilliklandirmisDegerHesapla(efkRow, -2)
    efkYillik3 = yilliklandirmisDegerHesapla(efkRow, -3)
    efkYillik4 = yilliklandirmisDegerHesapla(efkRow, -4)
    efkYillik5 = yilliklandirmisDegerHesapla(efkRow, -5)
    efkYillik6 = yilliklandirmisDegerHesapla(efkRow, -6)
    efkYillik7 = yilliklandirmisDegerHesapla(efkRow, -7)
    efkYillik8 = yilliklandirmisDegerHesapla(efkRow, -8)


    efkTablosuYillik = PrettyTable()
    efkTablosuYillik.field_names = ["ÇEYREK", "YILLIK EFK", "% DEĞİŞİM"]
    efkTablosuYillik.align["YILLIK EFK"] = "r"
    efkTablosuYillik.align["YÜZDE DEĞİŞİM"] = "r"
    efkTablosuYillik.add_row([bilancoDoneminiBul(0), "{:,.0f}".format(efkYillik0).replace(",", "."), "{:.2%}".format(efkYillik0/efkYillik1-1)])
    efkTablosuYillik.add_row([bilancoDoneminiBul(-1), "{:,.0f}".format(efkYillik1).replace(",", "."), "{:.2%}".format(efkYillik1/efkYillik2-1)])
    efkTablosuYillik.add_row([bilancoDoneminiBul(-2), "{:,.0f}".format(efkYillik2).replace(",", "."), "{:.2%}".format(efkYillik2/efkYillik3-1)])
    efkTablosuYillik.add_row([bilancoDoneminiBul(-3), "{:,.0f}".format(efkYillik3).replace(",", "."), "{:.2%}".format(efkYillik3/efkYillik4-1)])
    efkTablosuYillik.add_row([bilancoDoneminiBul(-4), "{:,.0f}".format(efkYillik4).replace(",", "."), "{:.2%}".format(efkYillik4/efkYillik5-1)])
    efkTablosuYillik.add_row([bilancoDoneminiBul(-5), "{:,.0f}".format(efkYillik5).replace(",", "."), "{:.2%}".format(efkYillik5/efkYillik6-1)])
    efkTablosuYillik.add_row([bilancoDoneminiBul(-6), "{:,.0f}".format(efkYillik6).replace(",", "."), "{:.2%}".format(efkYillik6/efkYillik7-1)])
    efkTablosuYillik.add_row([bilancoDoneminiBul(-7), "{:,.0f}".format(efkYillik7).replace(",", "."), "{:.2%}".format(efkYillik7/efkYillik8-1)])

    print(efkTablosuYillik)

# ------------------------------------

















# ÇEYREKLİK NET KAR TABLOSU YAZDIR

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
    netKarYillik8 = yilliklandirmisDegerHesapla (netKarRow, -8)

    netKarTablosuYillik = PrettyTable()
    netKarTablosuYillik.field_names = ["ÇEYREK", "YILLIK NET KAR", "% DEĞİŞİM"]
    netKarTablosuYillik.align["YILLIK NET KAR"] = "r"
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
