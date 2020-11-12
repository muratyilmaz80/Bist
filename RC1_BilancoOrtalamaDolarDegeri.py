import xlrd
import xlwt
from xlutils.copy import copy
import os.path
from datetime import datetime, timedelta
from RC1_GetDolarDegeri import DovizKurlari
from datetime import datetime

#dolarKurlariFile = "//Users//myilmaz//Documents//bist//Dolar_Kurlari.xlsx"

dovizKurlari = DovizKurlari()

# def tarihtekiDolarDegeriniBul(tarih):
#     wb = xlrd.open_workbook(dolarKurlariFile)
#     sheet = wb.sheet_by_index(0)
#
#     for rowi in range(sheet.nrows):
#         cell = sheet.cell(rowi, 0)
#         if cell.value == tarih:
#             while sheet.cell_value(rowi, 1) == "":
#                 #print(sheet.cell_value(rowi,0) , "tatil gününe denk geliyor, bir sonraki tarihe bakılıyor...")
#                 rowi = rowi + 1
#             #print (sheet.cell_value(rowi,0), "tarihindeki dolar değeri:")
#             return sheet.cell_value(rowi,1)
#     print("Verilen Tarihteki Dolar Değeri Bulunamadı!", tarih)
#     return 0



def tarihtekiDolarDegeriniBulOnline(tarih):

    dolarDegeri = dovizKurlari.Arsiv_tarih(tarih, "USD", "ForexBuying")
    date = datetime.strptime(tarih, "%d.%m.%Y").date()

    while (dolarDegeri == "Tatil Gunu"):
        print(tarih , "tatil gününe denk geliyor, bir sonraki tarihe bakılıyor...")
        date += timedelta(days=1)
        tarih = date.strftime("%d.%m.%Y")
        dolarDegeri = dovizKurlari.Arsiv_tarih(tarih, "USD", "ForexBuying")

    return dolarDegeri


def ucAylikBilancoDonemiOrtalamaDolarDegeriBul(bilancoDonemi):
    bitisYil = int(bilancoDonemi / 100)
    bitisAy = int(bilancoDonemi % 100)
    baslangicYil = bitisYil
    baslangicAy = bitisAy - 2

    baslangicAyString = str (baslangicAy)
    if (baslangicAy <10 ):
        baslangicAyString = "0" + str(baslangicAy)

    bitisAyString = str (bitisAy)
    if (bitisAy <10 ):
        bitisAyString = "0" + str(bitisAy)

    baslangicTarihi = "01." + baslangicAyString + "." + str(baslangicYil)
    bitisTarihi = "30." + bitisAyString + "." + str(bitisYil)
    print ("Başlangıç Tarihi:", baslangicTarihi)
    print("Bitiş Tarihi:", bitisTarihi)
    baslangicTarihiDolarDegeri = float (tarihtekiDolarDegeriniBulOnline(baslangicTarihi))
    bitisTarihiDolarDegeri = float (tarihtekiDolarDegeriniBulOnline(bitisTarihi))
    print("Başlangıç Tarihi Dolar Değeri:", baslangicTarihiDolarDegeri)
    print("Bitiş Tarihi Dolar Değeri:", bitisTarihiDolarDegeri)
    bilancoDonemiOrtalamaDolarDegeri = (baslangicTarihiDolarDegeri + bitisTarihiDolarDegeri) / 2
    print(bilancoDonemi, "Bilanço Dönemi Ortalama Dolar Değeri:", "{:.3}".format(bilancoDonemiOrtalamaDolarDegeri), "TL")
    return bilancoDonemiOrtalamaDolarDegeri

ucAylikBilancoDonemiOrtalamaDolarDegeriBul(202006)
