import xlrd
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri
import os
import sys

bilancoDonemi = 202209
directory = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar"


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


def toplamKaynaklar(varHisseAdi, varBilancoDonemi):
    try:
        sermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)
        hisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
        pd = sermaye * hisseFiyati;
        toplamKaynaklar =  getBilancoDegeri("TOPLAM KAYNAKLAR", 0)
        print(varHisseAdi, "\t", "{:,.0f}".format(toplamKaynaklar).replace(",", "."), "\t", "{:,.0f}".format(pd).replace(",", "."))
    except Exception as e:
        print(varHisseAdi, "\t", "HATA")

def sermayeArtirimPot (varHisseAdi, varBilancoDonemi):
    odenmisSermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)
    ozkaynaklar = getBilancoDegeri("TOPLAM ÖZKAYNAKLAR", 0)

    try:
        sermayeArtirimPotansiyeli = (ozkaynaklar - odenmisSermaye) / odenmisSermaye
        print(varHisseAdi, "\t", "{:.0%}".format(sermayeArtirimPotansiyeli))
    except Exception as e:
        print(varHisseAdi, "\t", "HATA")



l=os.listdir(directory)
list=[x.split('.')[0] for x in l]
list.sort()
print (list)

for x in list:
    hisseAdi = x
    bilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + hisseAdi + ".xlsx"
    wb = xlrd.open_workbook(bilancoDosyasi)
    sheet = wb.sheet_by_index(0)
    # sermayeArtirimPot(x, 0)
    toplamKaynaklar(x,0)
