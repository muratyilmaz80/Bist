import xlrd
import xlwt
from xlrd import xldate_as_tuple
from xlutils.copy import copy
import os.path
from datetime import datetime, timedelta
from RC1_GetDolarDegeri import DovizKurlari
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

def buguneKadarDolarVerisiGuncelle():
    today = datetime.today()
    todayStr = today.date().strftime("%d.%m.%Y")
    delta = timedelta(days=1)
    print ("Bugünün tarihi: ", todayStr)

    bookRead = xlrd.open_workbook(veriTabaniFile, formatting_info=True)
    sheetRead = bookRead.sheet_by_name("DolarKuru")
    rowNumber = sheetRead.nrows
    sonTarihExcel = sheetRead.cell_value((rowNumber-1), 0)

    sonTarih = datetime(*xlrd.xldate_as_tuple(sonTarihExcel, 0))
    sonTarihStr = sonTarih.date().strftime("%d.%m.%Y")
    print("Son Tarih:", sonTarihStr)

    start_date = datetime.strptime(sonTarihStr, "%d.%m.%Y").date()
    end_date = datetime.strptime(todayStr, "%d.%m.%Y").date() + delta

    bookWrite = copy(bookRead)
    bookSheetWrite = bookWrite.get_sheet("DolarKuru")

    date_format = xlwt.XFStyle()
    date_format.num_format_str = "dd/mm/yyyy"

    for i in range((end_date - start_date).days-1):

        tempDate = start_date + (i+1) * delta
        dateToPrint = tempDate.strftime("%d.%m.%Y")
        print(dateToPrint)
        value = dovizKurlari.Arsiv_tarih(dateToPrint, "USD", "ForexBuying")

        if (value!= "Tatil Gunu"):
            valueToPrint = float (value)
        else:
            valueToPrint = "Tatil Gunu"

        bookSheetWrite.write((rowNumber+i), 0, tempDate, date_format)
        bookSheetWrite.write((rowNumber + i), 1, valueToPrint)

    bookWrite.save(veriTabaniFile)



    # if os.path.isfile(veriTabaniFile):
    #     print("Veri tabanı dosyası var:", veriTabaniFile)
    #     bookRead = xlrd.open_workbook(veriTabaniFile)
    #     sheetRead = bookRead.sheet_by_name("DolarKuru")
    #     rowNumber = sheetRead.nrows
    #
    #     for rowi in range(sheetRead.nrows):
    #         cell = sheetRead.cell(rowi, 0)
    #         if cell.value == bilancoDonemi:
    #             print ("Veritabanında bilanço dönemi ortalama dolar bilgisi mevcut.")
    #             ortalamaDolarDegeri = sheetRead.cell_value(rowi, 1)
    #             return ortalamaDolarDegeri
    #
    #     if (ortalamaDolarDegeri == 0):
    #         print ("Bilanço dönemi için dolar bilgisi hesaplanacak.")
    #         ortalamaDolarDegeri = ucAylikBilancoDonemiOrtalamaDolarDegeriHesapla(bilancoDonemi)
    #         bookWrite = copy(bookRead)
    #         bookSheetWrite = bookWrite.get_sheet("DolarKuru")
    #         bookSheetWrite.write(rowNumber, 0, bilancoDonemi)
    #         bookSheetWrite.write(rowNumber, 1, ortalamaDolarDegeri)
    #         bookWrite.save(veriTabaniFile)
    #         return ortalamaDolarDegeri
    #
    # else:
    #     print("Veritabanı dosyası yeni oluşturulacak: ", veriTabaniFile)
    #     print("Bilanço dönemi için dolar bilgisi hesaplanacak.")
    #     ortalamaDolarDegeri = ucAylikBilancoDonemiOrtalamaDolarDegeriHesapla(bilancoDonemi)
    #     bookWrite = xlwt.Workbook()
    #     bookSheetWrite = bookWrite.add_sheet("DolarKuru")
    #     bookSheetWrite.write(0, 0, bilancoDonemi)
    #     bookSheetWrite.write(0, 1, ortalamaDolarDegeri)
    #     bookWrite.save(veriTabaniFile)
    #     return ortalamaDolarDegeri


buguneKadarDolarVerisiGuncelle()


    # wb = xlrd.open_workbook(veriTabaniFile)
    # sheet = wb.sheet_by_name()sheet_by_index(0)
    #
    #
    #
    # for rowi in range(sheet.nrows):
    #     cell = sheet.cell(rowi, 0)
    #     if cell.value == tarih:
    #         while sheet.cell_value(rowi, 1) == "":
    #             #print(sheet.cell_value(rowi,0) , "tatil gününe denk geliyor, bir sonraki tarihe bakılıyor...")
    #             rowi = rowi + 1
    #         #print (sheet.cell_value(rowi,0), "tarihindeki dolar değeri:")
    #         return sheet.cell_value(rowi,1)
    # print("Verilen Tarihteki Dolar Değeri Bulunamadı!", tarih)
    # return 0