import xlrd
import xlwt
from xlutils.copy import copy
import os.path
from datetime import datetime, timedelta
from GetDolarDegeriOnline import DovizKurlari
from datetime import datetime

veriTabaniFile = "//Users//myilmaz//Documents//bist//VeriTabani.xls"

dovizKurlari = DovizKurlari()

def tarihtekiDolarDegeriniBulOnline(tarih):

    dolarDegeri = dovizKurlari.Arsiv_tarih(tarih, "USD", "ForexBuying")
    date = datetime.strptime(tarih, "%d.%m.%Y").date()

    while (dolarDegeri == "Tatil Gunu"):
        print(tarih , "tatil gününe denk geliyor, bir sonraki tarihe bakılıyor...")
        date += timedelta(days=1)
        tarih = date.strftime("%d.%m.%Y")
        dolarDegeri = dovizKurlari.Arsiv_tarih(tarih, "USD", "ForexBuying")

    return dolarDegeri


def tarihtekiDolarDegeriniBulVeriTabani(tarih):
    wb = xlrd.open_workbook(veriTabaniFile)
    sheet = wb.sheet_by_name("DolarKuru")


    for rowi in range(sheet.nrows):
        cell = sheet.cell_value(rowi, 0)
        okunanTarih = datetime(*xlrd.xldate_as_tuple(cell, 0))
        okunanTarihStr = okunanTarih.date().strftime("%d.%m.%Y")
        if okunanTarihStr == tarih:
            # while sheet.cell_value(rowi, 1) == "Tatil Gunu":
            #     # print(datetime(*xlrd.xldate_as_tuple(sheet.cell_value(rowi,0), 0)).date().strftime("%d.%m.%Y")
            #       #   , "tatil gününe denk geliyor, bir sonraki tarihe bakılıyor...")
            #     rowi = rowi + 1
            # #print (datetime(*xlrd.xldate_as_tuple(sheet.cell_value(rowi,0), 0)).date().strftime("%d.%m.%Y"),
            #  #      "tarihindeki dolar değeri:", sheet.cell_value(rowi,1))
            return sheet.cell_value(rowi,1)
    print("Verilen Tarihteki Dolar Değeri Bulunamadı!", okunanTarihStr)
    return 0



def ucAylikBilancoDonemiOrtalamaDolarDegeriHesapla(bilancoDonemi, yontem):
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

    delta = timedelta(days=1)
    baslangicTarihi = "01." + baslangicAyString + "." + str(baslangicYil)
    bitisTarihi = "30." + bitisAyString + "." + str(bitisYil)
    print ("Başlangıç Tarihi:", baslangicTarihi)
    print("Bitiş Tarihi:", bitisTarihi)
    #baslangicTarihiDolarDegeri = float (tarihtekiDolarDegeriniBulOnline(baslangicTarihi))
    #bitisTarihiDolarDegeri = float (tarihtekiDolarDegeriniBulOnline(bitisTarihi))
    #print("Başlangıç Tarihi Dolar Değeri:", baslangicTarihiDolarDegeri)
    #print("Bitiş Tarihi Dolar Değeri:", bitisTarihiDolarDegeri)

    toplamDeger = 0
    elemanSayisi = 0

    start_date = datetime.strptime(baslangicTarihi, "%d.%m.%Y").date()
    end_date = datetime.strptime(bitisTarihi, "%d.%m.%Y").date() + delta

    if (yontem=="online"):
        print ("Online veri alınarak hesaplama yapılacak.")
        for i in range((end_date - start_date).days):
            tempDate = start_date + i * delta

            if (dovizKurlari.Arsiv_tarih(tempDate.strftime("%d.%m.%Y"), "USD", "ForexBuying"))!="Tatil Gunu":
                tempDeger = dovizKurlari.Arsiv_tarih(tempDate.strftime("%d.%m.%Y"), "USD", "ForexBuying")
                print(tempDate, tempDeger)
                toplamDeger = toplamDeger + float(tempDeger)
                elemanSayisi = elemanSayisi + 1
        return toplamDeger/elemanSayisi

    if (yontem=="veritabani"):
        print ("Veritabanindan hesaplanacak")

        for i in range((end_date - start_date).days):
            tempDate = start_date + i * delta
            if (tarihtekiDolarDegeriniBulVeriTabani(tempDate.strftime("%d.%m.%Y"))) != "Tatil Gunu":
                tempDeger = dovizKurlari.Arsiv_tarih(tempDate.strftime("%d.%m.%Y"), "USD", "ForexBuying")
                print(tempDate, tempDeger)
                toplamDeger = toplamDeger + float(tempDeger)
                elemanSayisi = elemanSayisi + 1
        return toplamDeger / elemanSayisi


def ucAylikBilancoDonemiOrtalamaDolarDegeriBul(bilancoDonemi):
    ortalamaDolarDegeri = 0

    if os.path.isfile(veriTabaniFile):
        # logging.debug("Veri tabanı dosyası var: %s", veriTabaniFile)
        bookRead = xlrd.open_workbook(veriTabaniFile, formatting_info=True)
        sheetRead = bookRead.sheet_by_name("OrtDolarDegeri")
        rowNumber = sheetRead.nrows

        for rowi in range(sheetRead.nrows):
            cell = sheetRead.cell(rowi, 0)
            if cell.value == bilancoDonemi:
                # logging.debug ("Veritabanında bilanço dönemi ortalama dolar bilgisi mevcut.")
                ortalamaDolarDegeri = sheetRead.cell_value(rowi, 1)
                return ortalamaDolarDegeri

        if (ortalamaDolarDegeri == 0):
            print ("Bilanço dönemi için dolar bilgisi hesaplanacak.")
            ortalamaDolarDegeri = ucAylikBilancoDonemiOrtalamaDolarDegeriHesapla(bilancoDonemi,"veritabani")
            bookWrite = copy(bookRead)
            bookSheetWrite = bookWrite.get_sheet("OrtDolarDegeri")
            bookSheetWrite.write(rowNumber, 0, bilancoDonemi)
            bookSheetWrite.write(rowNumber, 1, ortalamaDolarDegeri)
            bookWrite.save(veriTabaniFile)
            return ortalamaDolarDegeri

    else:
        print("Veritabanı dosyası yeni oluşturulacak: ", veriTabaniFile)
        print("Bilanço dönemi için dolar bilgisi hesaplanacak.")
        ortalamaDolarDegeri = ucAylikBilancoDonemiOrtalamaDolarDegeriHesapla(bilancoDonemi,"veritabani")
        bookWrite = xlwt.Workbook()
        bookSheetWrite = bookWrite.add_sheet("OrtDolarDegeri")
        bookSheetWrite.write(0, 0, bilancoDonemi)
        bookSheetWrite.write(0, 1, ortalamaDolarDegeri)
        bookWrite.save(veriTabaniFile)
        return ortalamaDolarDegeri


def birOncekiBilancoDoneminiHesapla(dnm):
    yil = int(dnm / 100)
    ceyrek = int(dnm % 100)

    if ceyrek == 3:
        return (yil - 1) * 100 + 12
    else:
        return yil * 100 + (ceyrek - 3)

def altiAylikBilancoDonemiOrtalamaDolarDegeriBul(bilancoDonemi):
    temp1 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(bilancoDonemi)
    temp2 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(birOncekiBilancoDoneminiHesapla(bilancoDonemi))
    return (temp1+temp2)/2


# print ("Bilanço Dönemi ortalama dolar kuru:", "{:,.3f}".format(ucAylikBilancoDonemiOrtalamaDolarDegeriBul(202403)))

# print ("Bilanço Dönemi ortalama dolar kuru:", "{:,.3f}".format(altiAylikBilancoDonemiOrtalamaDolarDegeriBul(202012)))

# print (tarihtekiDolarDegeriniBulVeriTabani("25.10.2010"))

# print (ucAylikBilancoDonemiOrtalamaDolarDegeriBul(202312))
