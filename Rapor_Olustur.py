import xlrd
import xlwt
from xlutils.copy import copy
import os.path


def exportReportExcel(hisse,file,bilancoDonemi,ExcelRowClass):

    def createTopRow():
        bookSheetWrite.write(0, 0, "Hisse Adı")
        bookSheetWrite.write(0, 1, "Bilanço Dönemi Hasılat")
        bookSheetWrite.write(0, 2, "Önceki Yıl Aynı Çeyrek Hasılat")
        bookSheetWrite.write(0, 3, "Bilanço Dönemi Hasılat Değişimi")
        bookSheetWrite.write(0, 4, "Bir Önceki Bilanço Dönemi Hasılat")
        bookSheetWrite.write(0, 5, "Beş Önceki Bilanço Dönemi Hasılat")
        bookSheetWrite.write(0, 6, "Bir Önceki Bilanço Dönemi Hasılat Değişimi")
        bookSheetWrite.write(0, 7, "Bilanço Dönemi Hasılat Gelir Artışı Geçme Durumu")
        bookSheetWrite.write(0, 8, "Önceki Bilanço Dönemi Hasılat Gelir Artışı Geçme Durumu")
        bookSheetWrite.write(0, 9, "Bilanço Dönemi Faaliyet Karı")
        bookSheetWrite.write(0, 10, "Önceki Yıl Aynı Çeyrek Faaliyet Karı")
        bookSheetWrite.write(0, 11, "Bilanço Dönemi Faaliyet Karı Değişimi")
        bookSheetWrite.write(0, 12, "Bir Önceki Bilanço Dönemi Faaliyet Karı")
        bookSheetWrite.write(0, 13, "Beş Önceki Bilanço Dönemi Faaliyet Karı")
        bookSheetWrite.write(0, 14, "Önceki Bilanço Dönemi Faaliyet Kari Değişimi")
        bookSheetWrite.write(0, 15, "Bilanço Dönemi Faaliyet Karı Artışı Geçme Durumu")
        bookSheetWrite.write(0, 16, "Önceki Bilanço Dönemi Faaliyet Karı Artışı Geçme Durumu")
        bookSheetWrite.write(0, 17, "Bilanço Dönemi Ortalama Dolar Kuru")
        bookSheetWrite.write(0, 18, "Bilanço Dönemi Dolar Hasılat")
        bookSheetWrite.write(0, 19, "Önceki Yıl Aynı Çeyrek Dolar Hasılat")
        bookSheetWrite.write(0, 20, "Bilanço Dönemi Dolar Hasılat Değişimi")
        bookSheetWrite.write(0, 21, "Bir Önceki Bilanço Dönemi Dolar Hasılat Değişimi")
        bookSheetWrite.write(0, 22, "Bilanço Dönemi Dolar Hasılat Gelir Artışı Geçme Durumu")
        bookSheetWrite.write(0, 23, "Önceki Bilanço Dönemi Dolar Hasılat Gelir Artışı Geçme Durumu")
        bookSheetWrite.write(0, 24, "Bilanço Dönemi Dolar Faaliyet Kari")
        bookSheetWrite.write(0, 25, "Önceki Yıl Aynı Çeyrek Dolar Faaliyet Karı")
        bookSheetWrite.write(0, 26, "Bilanço Dönemi Dolar Faaliyet Karı Değişimi")
        bookSheetWrite.write(0, 27, "Önceki Bilanço Dönemi Dolar Faaliyet Karı Değişimi")
        bookSheetWrite.write(0, 28, "Bilanço Dönemi Dolar Faaliyet Kari Artışı Geçme Durumu")
        bookSheetWrite.write(0, 29, "Önceki Bilanço Dönemi Dolar Faaliyet Kari Artışı Geçme Durumu")
        bookSheetWrite.write(0, 30, "Sermaye")
        bookSheetWrite.write(0, 31, "Ana Ortaklık Payı")
        bookSheetWrite.write(0, 32, "Son 4 Bilanço Dönemi Hasılat Toplamı")
        bookSheetWrite.write(0, 33, "Önümüzdeki 4 Bilanço Dönemi Hasılat Tahmini")
        bookSheetWrite.write(0, 34, "Önümüzdeki 4 Bilanço Dönemi Faaliyet Kar Marjı Tahmini")
        bookSheetWrite.write(0, 35, "Faaliyet Kar Tahmini 1")
        bookSheetWrite.write(0, 36, "Faaliyet Kar Tahmini 2")
        bookSheetWrite.write(0, 37, "Ortalama Faaliyet Kar Tahmini")
        bookSheetWrite.write(0, 38, "Hisse Başına Kar Tahmini")
        bookSheetWrite.write(0, 39, "Bilanço Etkisi")
        bookSheetWrite.write(0, 40, "Bilanço Tarihi Hisse Fiyatı")
        bookSheetWrite.write(0, 41, "Gerçek Hisse Değeri")
        bookSheetWrite.write(0, 42, "Target Buy")
        bookSheetWrite.write(0, 43, "Gerçek Fiyata Uzaklık")
        bookSheetWrite.write(0, 44, "NetPro")
        bookSheetWrite.write(0, 45, "ForwardPE")

        bookSheetWrite.write(0, 46, "Gerçek Hisse Değeri NFK")
        bookSheetWrite.write(0, 47, "Target Buy NFK")
        bookSheetWrite.write(0, 48, "Gerçek Fiyata Uzaklık NFK")
        bookSheetWrite.write(0, 49, "NetPro NFK")
        bookSheetWrite.write(0, 50, "ForwardPE NFK")

        bookSheetWrite.write(0, 51, "Tarih")
        bookSheetWrite.write(0, 52, "Net Kar Büyüme Yıllık")
        bookSheetWrite.write(0, 53, "Net Kar Büyüme 4 Önceki Çeyreğe Göre")
        bookSheetWrite.write(0, 54, "Esas Faaliyet Karı Büyüme Yıllık")
        bookSheetWrite.write(0, 55, "Hasılat Büyüme Yıllık")
        bookSheetWrite.write(0, 56, "FAVÖK Büyüme Yıllık")
        bookSheetWrite.write(0, 57, "F/K")
        bookSheetWrite.write(0, 58, "Nakit/PD")
        bookSheetWrite.write(0, 59, "Nakit/FD")
        bookSheetWrite.write(0, 60, "PD/DD")
        bookSheetWrite.write(0, 61, "PEG")
        bookSheetWrite.write(0, 62, "FD/Satışlar")
        bookSheetWrite.write(0, 63, "FD/FAVÖK")
        bookSheetWrite.write(0, 64, "PD/EFK")
        bookSheetWrite.write(0, 65, "Cari Oran")
        bookSheetWrite.write(0, 66, "Likit Oranı")
        bookSheetWrite.write(0, 67, "Nakit Oranı")
        bookSheetWrite.write(0, 68, "Asit Test Oranı")
        bookSheetWrite.write(0, 69, "ROE (Özsermaye Karlılığı)")
        bookSheetWrite.write(0, 70, "ROA (Aktif Karlılık)")
        bookSheetWrite.write(0, 71, "Yıllık Net Kar Marjı")
        bookSheetWrite.write(0, 72, "Son Çeyrek Net Kar Marjı")
        bookSheetWrite.write(0, 73, "Aktif Devir Hızı")
        bookSheetWrite.write(0, 74, "Borç/Kaynak")
        bookSheetWrite.write(0, 75, "Özsermaye Büyümesi")
        bookSheetWrite.write(0, 76, "Halka Açıklık Oranı")
        bookSheetWrite.write(0, 77, "Piyasa Değeri Milyon TL")
        bookSheetWrite.write(0, 78, "Sermaye Milyon TL")
        bookSheetWrite.write(0, 79, "Sermaye Artırım Potansiyeli")


    def reportResults(rowNumber):
        bookSheetWrite.write(rowNumber, 0, hisse)
        bookSheetWrite.write(rowNumber, 1, ExcelRowClass.bilancoDonemiHasilat)
        bookSheetWrite.write(rowNumber, 2, ExcelRowClass.oncekiYilAyniCeyrekHasilat)
        bookSheetWrite.write(rowNumber, 3, ExcelRowClass.bilancoDonemiHasilatDegisimi)
        bookSheetWrite.write(rowNumber, 4, ExcelRowClass.oncekiBilancoDonemiHasilat)
        bookSheetWrite.write(rowNumber, 5, ExcelRowClass.besOncekiBilancoDonemiHasilat)
        bookSheetWrite.write(rowNumber, 6, ExcelRowClass.oncekiBilancoDonemiHasilatDegisimi)
        bookSheetWrite.write(rowNumber, 7, ExcelRowClass.bilancoDonemiHasilatDegisimiGecmeDurumu)
        bookSheetWrite.write(rowNumber, 8, bool(ExcelRowClass.oncekiBilancoDonemiHasilatDegisimiGecmeDurumu))
        bookSheetWrite.write(rowNumber, 9, ExcelRowClass.bilancoDonemiFaaliyetKari)
        bookSheetWrite.write(rowNumber, 10, ExcelRowClass.oncekiYilAyniCeyrekFaaliyetKari)
        bookSheetWrite.write(rowNumber, 11, ExcelRowClass.bilancoDonemiFaaliyetKariDegisimi)
        bookSheetWrite.write(rowNumber, 12, ExcelRowClass.oncekiBilancoDonemiFaaliyetKari)
        bookSheetWrite.write(rowNumber, 13, ExcelRowClass.besOncekiBilancoDonemiFaaliyetKari)
        bookSheetWrite.write(rowNumber, 14, ExcelRowClass.oncekiBilancoDonemiFaaliyetKariDegisimi)
        bookSheetWrite.write(rowNumber, 15, ExcelRowClass.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu)
        bookSheetWrite.write(rowNumber, 16, bool(ExcelRowClass.oncekiBilancoDonemiFaaliyetKarDegisimiGecmeDurumu))
        bookSheetWrite.write(rowNumber, 17, ExcelRowClass.bilancoDonemiOrtalamaDolarKuru)
        bookSheetWrite.write(rowNumber, 18, ExcelRowClass.bilancoDonemiDolarHasilat)
        bookSheetWrite.write(rowNumber, 19, ExcelRowClass.oncekiYilAyniCeyrekDolarHasilat)
        bookSheetWrite.write(rowNumber, 20, ExcelRowClass.bilancoDonemiDolarHasilatDegisimi)
        bookSheetWrite.write(rowNumber, 21, ExcelRowClass.oncekiBilancoDonemiDolarHasilatDegisimi)
        bookSheetWrite.write(rowNumber, 22, ExcelRowClass.bilancoDonemiDolarHasilatDegisimiGecmeDurumu)
        bookSheetWrite.write(rowNumber, 23, bool(ExcelRowClass.oncekiBilancoDonemiDolarHasilatDegisimiGecmeDurumu))
        bookSheetWrite.write(rowNumber, 24, ExcelRowClass.bilancoDonemiDolarFaaliyetKari)
        bookSheetWrite.write(rowNumber, 25, ExcelRowClass.dortOncekiBilancoDonemiDolarFaaliyetKari)
        bookSheetWrite.write(rowNumber, 26, ExcelRowClass.bilancoDonemiDolarFaaliyetKariDegisimi)
        bookSheetWrite.write(rowNumber, 27, ExcelRowClass.oncekiBilancoDonemiDolarFaaliyetKariDegisimi)
        bookSheetWrite.write(rowNumber, 28, ExcelRowClass.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu)
        bookSheetWrite.write(rowNumber, 29, bool(ExcelRowClass.oncekiBilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu))
        bookSheetWrite.write(rowNumber, 30, ExcelRowClass.sermaye)
        bookSheetWrite.write(rowNumber, 31, ExcelRowClass.anaOrtaklikPayi)
        bookSheetWrite.write(rowNumber, 32, ExcelRowClass.sonDortBilancoDonemiHasilatToplami)
        bookSheetWrite.write(rowNumber, 33, ExcelRowClass.onumuzdekiDortBilancoDonemiHasilatTahmini)
        bookSheetWrite.write(rowNumber, 34, ExcelRowClass.onumuzdekiDortBilancoDonemiFaaliyetKarMarjiTahmini)
        bookSheetWrite.write(rowNumber, 35, ExcelRowClass.faaliyetKariTahmini1)
        bookSheetWrite.write(rowNumber, 36, ExcelRowClass.faaliyetKariTahmini2)
        bookSheetWrite.write(rowNumber, 37, ExcelRowClass.ortalamaFaaliyetKariTahmini)
        bookSheetWrite.write(rowNumber, 38, ExcelRowClass.hisseBasinaOrtalamaKarTahmini)
        bookSheetWrite.write(rowNumber, 39, ExcelRowClass.bilancoEtkisi)
        bookSheetWrite.write(rowNumber, 40, ExcelRowClass.bilancoTarihiHisseFiyati)
        bookSheetWrite.write(rowNumber, 41, ExcelRowClass.gercekHisseDegeri)
        bookSheetWrite.write(rowNumber, 42, ExcelRowClass.targetBuy)
        bookSheetWrite.write(rowNumber, 43, ExcelRowClass.gercekFiyataUzaklik)
        bookSheetWrite.write(rowNumber, 44, ExcelRowClass.netProKriteri)
        bookSheetWrite.write(rowNumber, 45, ExcelRowClass.forwardPeKriteri)

        bookSheetWrite.write(rowNumber, 46, ExcelRowClass.gercekHisseDegeriNfk)
        bookSheetWrite.write(rowNumber, 47, ExcelRowClass.targetBuyNfk)
        bookSheetWrite.write(rowNumber, 48, ExcelRowClass.gercekFiyataUzaklikNfk)
        bookSheetWrite.write(rowNumber, 49, ExcelRowClass.netProKriteriNfk)
        bookSheetWrite.write(rowNumber, 50, ExcelRowClass.forwardPeKriteriNfk)

        bookSheetWrite.write(rowNumber, 51, ExcelRowClass.tarih)
        bookSheetWrite.write(rowNumber, 52, ExcelRowClass.netKarBuyumeYillik)
        bookSheetWrite.write(rowNumber, 53, ExcelRowClass.netKarBuyume4OncekiCeyregeGore)
        bookSheetWrite.write(rowNumber, 54, ExcelRowClass.esasFaaliyetKariBuyumeYillik)
        bookSheetWrite.write(rowNumber, 55, ExcelRowClass.hasilatBuyumeYillik)
        bookSheetWrite.write(rowNumber, 56, ExcelRowClass.favokBuyumeYillik)
        bookSheetWrite.write(rowNumber, 57, ExcelRowClass.fkOrani)
        bookSheetWrite.write(rowNumber, 58, ExcelRowClass.nakitPd)
        bookSheetWrite.write(rowNumber, 59, ExcelRowClass.nakitFd)
        bookSheetWrite.write(rowNumber, 60, ExcelRowClass.pdDd)
        bookSheetWrite.write(rowNumber, 61, ExcelRowClass.pegOrani)
        bookSheetWrite.write(rowNumber, 62, ExcelRowClass.fdSatislar)
        bookSheetWrite.write(rowNumber, 63, ExcelRowClass.fdFavok)
        bookSheetWrite.write(rowNumber, 64, ExcelRowClass.pdEfk)
        bookSheetWrite.write(rowNumber, 65, ExcelRowClass.cariOran)
        bookSheetWrite.write(rowNumber, 66, ExcelRowClass.likitOrani)
        bookSheetWrite.write(rowNumber, 67, ExcelRowClass.nakitOrani)
        bookSheetWrite.write(rowNumber, 68, ExcelRowClass.asitTestOrani)
        bookSheetWrite.write(rowNumber, 69, ExcelRowClass.roe)
        bookSheetWrite.write(rowNumber, 70, ExcelRowClass.roa)
        bookSheetWrite.write(rowNumber, 71, ExcelRowClass.yillikNetKarMarji)
        bookSheetWrite.write(rowNumber, 72, ExcelRowClass.sonCeyrekNetKarMarji)
        bookSheetWrite.write(rowNumber, 73, ExcelRowClass.aktifDevirHizi)
        bookSheetWrite.write(rowNumber, 74, ExcelRowClass.borcKaynak)
        bookSheetWrite.write(rowNumber, 75, ExcelRowClass.ozsermayeBuyumesi)
        bookSheetWrite.write(rowNumber, 76, ExcelRowClass.halkaAciklikOrani)
        bookSheetWrite.write(rowNumber, 77, ExcelRowClass.piyasaDegeri)
        bookSheetWrite.write(rowNumber, 78, ExcelRowClass.sermaye)
        bookSheetWrite.write(rowNumber, 79, ExcelRowClass.sermayeArtirimPotansiyeli)



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
