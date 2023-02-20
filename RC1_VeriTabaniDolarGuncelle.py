import xlrd
import xlwt
from xlrd import xldate_as_tuple
from xlutils.copy import copy
import os.path
from datetime import datetime, timedelta
from RC1_GetDolarDegeriOnline import DovizKurlari
from datetime import datetime

veriTabaniFile = "//Users//myilmaz//Documents//bist//VeriTabani.xls"

dovizKurlari = DovizKurlari()

def buguneKadarDolarVerisiGuncelle():
    today = datetime.today()
    todayStr = today.date().strftime("%d.%m.%Y")
    # todayStr = "01.01.2023"
    delta = timedelta(days=1)
    print ("Bugünün tarihi: ", todayStr)
    sheetName = "DolarKuru"

    if os.path.isfile(veriTabaniFile):
        print ("Veri tabanı dosyası var, güncellenecek:", veriTabaniFile)
        bookRead = xlrd.open_workbook(veriTabaniFile, formatting_info=True)

        if sheetName in bookRead.sheet_names():
            print ("Worksheet var, güncellenecek.")
            sheetRead = bookRead.sheet_by_name(sheetName)
            rowNumber = sheetRead.nrows
            sonTarihExcel = sheetRead.cell_value((rowNumber - 1), 0)
            sonTarih = datetime(*xlrd.xldate_as_tuple(sonTarihExcel, 0))
            sonTarihStr = sonTarih.date().strftime("%d.%m.%Y")
            print("Son Tarih:", sonTarihStr)
            bookWrite = copy(bookRead)
            bookSheetWrite = bookWrite.get_sheet(sheetName)

        else:
            print("Worksheet yok, eklenecek.")
            rowNumber = 0
            sonTarih = datetime.strptime("31.12.1999", "%d.%m.%Y")
            sonTarihStr = sonTarih.date().strftime("%d.%m.%Y")
            print("Son Tarih:", sonTarihStr)
            bookWrite = copy(bookRead)
            bookSheetWrite = bookWrite.add_sheet(sheetName)

        start_date = datetime.strptime(sonTarihStr, "%d.%m.%Y").date()
        end_date = datetime.strptime(todayStr, "%d.%m.%Y").date() + delta
        # end_date = datetime.strptime("21.04.2022", "%d.%m.%Y").date() + delta

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


    else:
        print ("Veri tabanı dosyası yok, oluşturulacak:", veriTabaniFile)
        sonTarih = datetime.strptime("31.12.1999", "%d.%m.%Y")
        sonTarihStr = sonTarih.date().strftime("%d.%m.%Y")
        print("Son Tarih:", sonTarihStr)

        start_date = datetime.strptime(sonTarihStr, "%d.%m.%Y").date()
        end_date = datetime.strptime(todayStr, "%d.%m.%Y").date() + delta

        bookWrite = xlwt.Workbook()
        bookSheetWrite = bookWrite.add_sheet(sheetName)

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

            bookSheetWrite.write((0+i), 0, tempDate, date_format)
            bookSheetWrite.write((0+i), 1, valueToPrint)

    bookWrite.save(veriTabaniFile)

buguneKadarDolarVerisiGuncelle()