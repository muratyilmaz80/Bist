import xlrd
import xlwt
from xlutils.copy import copy
import os.path
import ExcelRowClass


def exportReportExcel(hisse,file,bilancoDonemi,ExcelRowClass):

    def createTopRow():
        bookSheetWrite.write(0, 0, "Hisse Adı")
        bookSheetWrite.write(0, 1, "Son Çeyrek Hasılat")
        bookSheetWrite.write(0, 2, "Önceki Yıl Aynı Çeyrek Hasılat")
        bookSheetWrite.write(0, 3, "Hasılat Artışı")
        bookSheetWrite.write(0, 4, "Bir Önceki Çeyrek Hasılat Artışı")
        bookSheetWrite.write(0, 5, "Kriter1")
        bookSheetWrite.write(0, 6, "Kriter3")
        bookSheetWrite.write(0, 7, "Son Çeyrek Faaliyet Karı")
        bookSheetWrite.write(0, 8, "Önceki Yıl Aynı Çeyrek Faaliyet Karı")
        bookSheetWrite.write(0, 9, "Faaliyet Karı Artışı")
        bookSheetWrite.write(0, 10, "Bir Önceki Çeyrek Faaliyet Karı Artışı")
        bookSheetWrite.write(0, 11, "Kriter2")
        bookSheetWrite.write(0, 12, "Kriter4")
        bookSheetWrite.write(0, 13, "Sermaye")
        bookSheetWrite.write(0, 14, "Ana Ortaklık Payı")
        bookSheetWrite.write(0, 15, "Son 4 Çeyrek Satış Toplamı")
        bookSheetWrite.write(0, 16, "Önümüzdeki 4 Çeyrek Satış Tahmini")
        bookSheetWrite.write(0, 17, "Önümüzdeki 4 Çeyrek Faaliyet Kar Marjı Tahmini")
        bookSheetWrite.write(0, 18, "Faaliyet Kar Tahmini 1")
        bookSheetWrite.write(0, 19, "Faaliyet Kar Tahmini 2")
        bookSheetWrite.write(0, 20, "Ortalama Faaliyet Kar Tahmini")
        bookSheetWrite.write(0, 21, "Hisse Başına Kar Tahmini")
        bookSheetWrite.write(0, 22, "Bilanço Etkisi")
        bookSheetWrite.write(0, 23, "Bilanço Tarihi Hisse Fiyatı")
        bookSheetWrite.write(0, 24, "Gerçek Hisse Değeri")
        bookSheetWrite.write(0, 25, "Target Buy")
        bookSheetWrite.write(0, 26, "Gerçek Fiyata Uzaklık")
        bookSheetWrite.write(0, 27, "NET Pro")
        bookSheetWrite.write(0, 28, "Forward PE")

    def reportResults(rowNumber):

        bookSheetWrite.write(rowNumber, 0, hisse)
        bookSheetWrite.write(rowNumber, 1, ExcelRowClass.sonCeyrekHasilat)
        bookSheetWrite.write(rowNumber, 2, ExcelRowClass.oncekiYilAyniCeyrekHasilat)
        bookSheetWrite.write(rowNumber, 3, ExcelRowClass.hasilatArtisi)
        bookSheetWrite.write(rowNumber, 4, ExcelRowClass.birOncekiCeyrekHasilatArtisi)
        bookSheetWrite.write(rowNumber, 5, ExcelRowClass.kriter1)
        bookSheetWrite.write(rowNumber, 6, ExcelRowClass.kriter3)
        bookSheetWrite.write(rowNumber, 7, ExcelRowClass.sonCeyrekFaaliyetKari)
        bookSheetWrite.write(rowNumber, 8, ExcelRowClass.oncekiYilAyniCeyrekFaaliyetKari)
        bookSheetWrite.write(rowNumber, 9, ExcelRowClass.faaliyetKarArtisi)
        bookSheetWrite.write(rowNumber, 10, ExcelRowClass.birOncekiCeyrekFaaliyetKarArtisi)
        bookSheetWrite.write(rowNumber, 11, ExcelRowClass.kriter2)
        bookSheetWrite.write(rowNumber, 12, ExcelRowClass.kriter4)
        bookSheetWrite.write(rowNumber, 13, ExcelRowClass.sermaye)
        bookSheetWrite.write(rowNumber, 14, ExcelRowClass.anaOrtaklikPayi)
        bookSheetWrite.write(rowNumber, 15, ExcelRowClass.son4CeyrekSatisToplami)
        bookSheetWrite.write(rowNumber, 16, ExcelRowClass.onumuzdeki4CeyrekSatisTahmini)
        bookSheetWrite.write(rowNumber, 17, ExcelRowClass.onumuzdeki4CeyrekFaaliyetKarMarjiTahmini)
        bookSheetWrite.write(rowNumber, 18, ExcelRowClass.faaliyetKarTahmini1)
        bookSheetWrite.write(rowNumber, 19, ExcelRowClass.faaliyetKarTahmini2)
        bookSheetWrite.write(rowNumber, 20, ExcelRowClass.ortalamaFaaliyetKarTahmini)
        bookSheetWrite.write(rowNumber, 21, ExcelRowClass.hisseBasinaKarTahmini)
        bookSheetWrite.write(rowNumber, 22, ExcelRowClass.bilancoEtkisi)
        bookSheetWrite.write(rowNumber, 23, ExcelRowClass.bilancoTarihiHisseFiyati)
        bookSheetWrite.write(rowNumber, 24, ExcelRowClass.gercekHisseDegeri)
        bookSheetWrite.write(rowNumber, 25, ExcelRowClass.targetBuy)
        bookSheetWrite.write(rowNumber, 26, ExcelRowClass.gercekFiyataUzaklik)
        bookSheetWrite.write(rowNumber, 27, ExcelRowClass.netPro)
        bookSheetWrite.write(rowNumber, 28, ExcelRowClass.forwardPe)

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