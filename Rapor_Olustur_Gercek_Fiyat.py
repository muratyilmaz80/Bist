import xlrd
import xlwt
from xlutils.copy import copy
import os.path
import ExcelRowClass


def exportReportExcelGercekFiyat(hisse,file,bilancoDonemi,guncelDeger,gercekDeger):

    def createTopRow():
        bookSheetWrite.write(0, 0, "Hisse Adı")
        bookSheetWrite.write(0, 1, "Güncel Değer")
        bookSheetWrite.write(0, 2, "Gerçek Değer")
        bookSheetWrite.write(0, 3, "Gerçek Değere Uzaklık")
        bookSheetWrite.write(0, 4, "temp1")
        bookSheetWrite.write(0, 5, "temp2")

    def reportResults(rowNumber):

        bookSheetWrite.write(rowNumber, 0, hisse)
        bookSheetWrite.write(rowNumber, 1, guncelDeger)
        bookSheetWrite.write(rowNumber, 2, gercekDeger)
        bookSheetWrite.write(rowNumber, 3, guncelDeger/gercekDeger)
        bookSheetWrite.write(rowNumber, 4, 0)
        bookSheetWrite.write(rowNumber, 5, 1)

    if os.path.isfile(file):
        print("Rapor dosyası var, güncellenecek:", file)
        bookRead = xlrd.open_workbook(file, formatting_info=True)
        sheetRead = bookRead.sheet_by_index(0)
        rowNumber = sheetRead.nrows
        bookWrite = copy(bookRead)
        bookSheetWrite = bookWrite.get_sheet(0)
        reportResults(rowNumber)
        bookWrite.save(file)

    else:
        print("Rapor dosyası yeni oluşturulacak: ", file)
        bookWrite = xlwt.Workbook()
        bookSheetWrite = bookWrite.add_sheet(str(bilancoDonemi))
        createTopRow()
        reportResults(1)
        bookWrite.save(file)