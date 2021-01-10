import xlrd
from RC1_ExcelRowClass import ExcelRowClass
from RC1_Rapor_Olustur import exportReportExcel
from prettytable import PrettyTable
import logging
import sys
from RC1_BilancoOrtalamaDolarDegeri import altiAylikBilancoDonemiOrtalamaDolarDegeriBul


def runAlgoritma6Aylik(bilancoDosyasi, bilancoDonemi, bondYield, hisseFiyati, reportFile, logPath, logLevel):

    hisseAdiTemp = bilancoDosyasi[47:]
    hisseAdi = hisseAdiTemp[:-5]

    my_logger = logging.getLogger()
    my_logger.setLevel(logLevel)
    output_file_handler = logging.FileHandler(logPath + hisseAdi + ".txt")
    output_file_handler.level = logging.INFO
    # output_file_handler.setFormatter(logging.Formatter("%(asctime)s — %(name)s — %(levelname)s — %(message)s"))
    stdout_handler = logging.StreamHandler(sys.stdout)
    # stdout_handler.setFormatter(logging.Formatter("%(asctime)s — %(name)s — %(levelname)s — %(message)s"))
    my_logger.addHandler(output_file_handler)
    my_logger.addHandler(stdout_handler)

    my_logger.info ("-------------------------------- %s ------------------------", hisseAdi)

    def birOncekiBilancoDoneminiHesapla(dnm):
        yil = int(dnm / 100)
        ceyrek = int(dnm % 100)

        if ceyrek == 6:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + 6

    my_logger.debug("Bilanco Donemi: %d", bilancoDonemi)

    birOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(bilancoDonemi)
    my_logger.debug("Bir Onceki Bilanco Donemi: %d", birOncekiBilancoDonemi)

    ikiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(birOncekiBilancoDonemi)
    my_logger.debug("Iki Onceki Bilanco Donemi: %d", ikiOncekiBilancoDonemi)

    ucOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ikiOncekiBilancoDonemi)
    my_logger.debug("Uc Onceki Bilanco Donemi: %d", ucOncekiBilancoDonemi)

    wb = xlrd.open_workbook(bilancoDosyasi)
    sheet = wb.sheet_by_index(0)

    def donemColumnFind(col):
        for columni in range(sheet.ncols):
            cell = sheet.cell(0, columni)
            if cell.value == col:
                return columni
        my_logger.info("Uygun Donem Bulunamadi!!!")
        return -1

    bilancoDonemiColumn = donemColumnFind(bilancoDonemi)
    birOncekibilancoDonemiColumn = donemColumnFind(birOncekiBilancoDonemi)
    ikiOncekibilancoDonemiColumn = donemColumnFind(ikiOncekiBilancoDonemi)
    ucOncekibilancoDonemiColumn = donemColumnFind(ucOncekiBilancoDonemi)

    def getBilancoDegeri(label, column):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == label:
                if sheet.cell_value(rowi, column)=="":
                    my_logger.info (label + " :Bilanço alanı boş!")
                    return 0
                else:
                    return sheet.cell_value(rowi, column)
        my_logger.info("Uygun bilanco degeri bulunamadi: %s", label)
        return 0


    def getBilancoTitleRow(title):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == title:
                return rowi
        my_logger.info("Uygun baslik bulunamadi!")
        return -1

    hasilatRow = getBilancoTitleRow("Hasılat")
    faaliyetKariRow = getBilancoTitleRow("ESAS FAALİYET KARI (ZARARI)")
    # netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı")
    netKarRow = getBilancoTitleRow("DÖNEM KARI (ZARARI)")
    brutKarRow = getBilancoTitleRow("BRÜT KAR (ZARAR)")

    # TODO: Bir önceki çeyrek bilançosunun olmasını garanti edecek şekilde düzenle

    def altiAyDegeriHesapla(r, c):
        quarter = (sheet.cell_value(0, c)) % (100)
        if (quarter == 6):
            return sheet.cell_value(r, c)
        else:
            if (sheet.cell_value(0, c) - sheet.cell_value(0, (c - 1)) == 6):
                return (sheet.cell_value(r, c) - sheet.cell_value(r, (c - 1)))
            else:
                my_logger.info("EKSİK BİLANÇO VAR!")
                return -1


    def oncekiYilAyniAltiAyDegisimiHesapla(row, donem):
        my_logger.debug("fonksiyon: oncekiYilAyniAltiAyDegisimiHesapla")
        donemColumn = donemColumnFind(donem)
        my_logger.debug ("DonemColumn: %s", donemColumn)
        oncekiYilAyniDonemColumn = donemColumnFind(donem - 100)
        my_logger.debug("Onceki Yıl Aynı DonemColumn: %s", oncekiYilAyniDonemColumn)
        my_logger.debug("Row: %d Column: %d",row ,donemColumn)
        ceyrekDegeri = altiAyDegeriHesapla(row, donemColumn)
        my_logger.debug("Altı Ay Değeri: %d", ceyrekDegeri)
        oncekiCeyrekDegeri = altiAyDegeriHesapla(row, oncekiYilAyniDonemColumn)
        my_logger.debug ("Önceki Altı Ay Değeri: %d", oncekiCeyrekDegeri)
        degisimSonucu = ceyrekDegeri / oncekiCeyrekDegeri - 1
        my_logger.debug("%d %s %d", sheet.cell_value(0, donemColumn), sheet.cell_value(row, 0), ceyrekDegeri)
        my_logger.debug("%d %s %d" ,sheet.cell_value(0, oncekiYilAyniDonemColumn), sheet.cell_value(row, 0), oncekiCeyrekDegeri)
        return degisimSonucu

    def likidasyonDegeriHesapla(ceyrek):
        nakit = getBilancoDegeri("Nakit ve Nakit Benzerleri", bilancoDonemiColumn)
        alacaklar = getBilancoDegeri("Ticari Alacaklar", bilancoDonemiColumn) + getBilancoDegeri("Diğer Alacaklar",
                                                                                                 bilancoDonemiColumn) + getBilancoDegeri(
            "Ticari Alacaklar1", bilancoDonemiColumn)
        stoklar = getBilancoDegeri("Stoklar", bilancoDonemiColumn)
        digerVarliklar = getBilancoDegeri("Diğer Dönen Varlıklar", bilancoDonemiColumn)
        finansalVarliklar = getBilancoDegeri("Finansal Yatırımlar", bilancoDonemiColumn) + getBilancoDegeri(
            "Finansal Yatırımlar1", bilancoDonemiColumn) + getBilancoDegeri("Özkaynak Yöntemiyle Değerlenen Yatırımlar",
                                                                            bilancoDonemiColumn)
        maddiDuranVarliklar = getBilancoDegeri("Maddi Duran Varlıklar", bilancoDonemiColumn)

        likidasyonDegeri = nakit + (alacaklar * 0.7) + (stoklar * 0.5) + (digerVarliklar * 0.7) + (
                    finansalVarliklar * 0.7) + (maddiDuranVarliklar * 0.2)

        return likidasyonDegeri


    # Bilanço Dönemi Satış(Hasılat) Gelirleri
    my_logger.info("")
    my_logger.info("")
    my_logger.info("--------------------HASILAT(SATIŞ) GELİRLERİ---------------------------")
    my_logger.info("")

    bilancoDonemiHasilat = altiAyDegeriHesapla(hasilatRow,bilancoDonemiColumn)
    birOncekiBilancoDonemiHasilat = altiAyDegeriHesapla(hasilatRow,birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiHasilat = altiAyDegeriHesapla(hasilatRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiHasilat = altiAyDegeriHesapla(hasilatRow, ucOncekibilancoDonemiColumn)

    bilancoDonemiHasilatPrint = "{:,.0f}".format(bilancoDonemiHasilat).replace(",", ".")
    ikiOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiHasilat).replace(",", ".")
    bilancoDonemiHasilatDegisimi = oncekiYilAyniAltiAyDegisimiHesapla(hasilatRow, bilancoDonemi)
    bilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(bilancoDonemiHasilatDegisimi)

    birOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(birOncekiBilancoDonemiHasilat).replace(",", ".")
    ucOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(ucOncekiBilancoDonemiHasilat).replace(",", ".")
    birOncekiBilancoDonemiHasilatDegisimi = oncekiYilAyniAltiAyDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)
    birOncekiBilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiHasilatDegisimi)

    satisTablosu = PrettyTable()
    satisTablosu.field_names = ["DÖNEM", "SATIŞ", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ", "YÜZDE DEĞİŞİM"]
    satisTablosu.align["SATIŞ"] = "r"
    satisTablosu.align["ÖNCEKİ YIL SATIŞ"] = "r"
    satisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    satisTablosu.add_row([bilancoDonemi, bilancoDonemiHasilatPrint, ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiHasilatPrint, bilancoDonemiHasilatDegisimiPrint])
    satisTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiHasilatPrint, ucOncekiBilancoDonemi, ucOncekiBilancoDonemiHasilatPrint, birOncekiBilancoDonemiHasilatDegisimiPrint])
    my_logger.info(satisTablosu)

    # Bilanço Dönemi Saış Geliri Artış Kriteri
    bilancoDonemiHasilatGelirArtisi = oncekiYilAyniAltiAyDegisimiHesapla(hasilatRow, bilancoDonemi)
    bilancoDonemiHasilatGelirArtisiGecmeDurumu = (bilancoDonemiHasilatGelirArtisi > 0.1)
    printText = "Bilanço Dönemi Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(bilancoDonemiHasilatGelirArtisi) + " >? 10% " + " " + str(bilancoDonemiHasilatGelirArtisiGecmeDurumu)
    my_logger.info(printText)

    # Önceki Dönem Hasılat Geliri Artış Kriteri
    oncekiDonemHasilatGelirArtisi = oncekiYilAyniAltiAyDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    if (bilancoDonemiHasilatGelirArtisi >= 1):
        my_logger.info ("Bilanço Dönemi Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak.")
        oncekiDonemHasilatGelirArtisiGecmeDurumu = True
        my_logger.info ("Önceki Dönem Satış Gelir Artışı Geçme Durumu: %s", oncekiDonemHasilatGelirArtisiGecmeDurumu)

    else:
        oncekiDonemHasilatGelirArtisiGecmeDurumu = (oncekiDonemHasilatGelirArtisi<bilancoDonemiHasilatGelirArtisi)
        printText = "Önceki Dönem Satış Gelir Artışı Bilanço Döneminden Düşük Mü: " + "{:.2%}".format(oncekiDonemHasilatGelirArtisi) + " <? " + "{:.2%}".format(bilancoDonemiHasilatGelirArtisi) + " " + str(oncekiDonemHasilatGelirArtisiGecmeDurumu)
        my_logger.info(printText)



    # Faaliyet Karı Gelirleri
    my_logger.info("")
    my_logger.info("")
    my_logger.info("--------------------------FAALİYET KARI---------------------------------")
    my_logger.info("")

    bilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow,bilancoDonemiColumn)
    birOncekiBilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow,birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)

    bilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(bilancoDonemiFaaliyetKari).replace(",", ".")
    ikiOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    bilancoDonemiFaaliyetKariDegisimi = oncekiYilAyniAltiAyDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
    bilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(bilancoDonemiFaaliyetKariDegisimi)

    birOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(birOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    ucOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(ucOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    birOncekiBilancoDonemiFaaliyetKariDegisimi = oncekiYilAyniAltiAyDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
    birOncekiBilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiFaaliyetKariDegisimi)

    faaliyetKariTablosu = PrettyTable()
    faaliyetKariTablosu.field_names = ["DÖNEM", "FAALİYET KARI", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI", "YÜZDE DEĞİŞİM"]
    faaliyetKariTablosu.align["FAALİYET KARI"] = "r"
    faaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI"] = "r"
    faaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    faaliyetKariTablosu.add_row([bilancoDonemi, bilancoDonemiFaaliyetKariPrint, ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiFaaliyetKariPrint, bilancoDonemiFaaliyetKariDegisimiPrint])
    faaliyetKariTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiFaaliyetKariPrint, ucOncekiBilancoDonemi, ucOncekiBilancoDonemiFaaliyetKariPrint, birOncekiBilancoDonemiFaaliyetKariDegisimiPrint])
    my_logger.info(faaliyetKariTablosu)


    # Bilanço Dönemi Faaliyet Kar Artış Kriteri
    if altiAyDegeriHesapla(netKarRow, bilancoDonemiColumn) < 0:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniAltiAyDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = False
        my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Çeyrek Net Kar Negatif", str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu))

    elif altiAyDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) < 0:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniAltiAyDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = False
        my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Ceyrek Faaliyet Kari Negatif", str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu))

    elif ((altiAyDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) > 0) and (altiAyDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)) < 0):
        bilancoDonemiFaaliyetKariArtisi = 0
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = True
        my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Dönem Faaliyet Karı Negatiften Pozitife Geçmiş", str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu))

    else:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniAltiAyDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = (bilancoDonemiFaaliyetKariArtisi > 0.15)
        printText = "Bilanço Dönemi Faaliyet Karı Artışı:" + "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi) + " >? 15% " + str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu)
        my_logger.info(printText)

    # Önceki Dönem Faaliyet Kar Artış Kriteri

    if bilancoDonemiFaaliyetKariArtisi >= 1:
        oncekiDonemFaaliyetKariArtisi = oncekiYilAyniAltiAyDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiDonemFaaliyetKarArtisiGecmeDurumu = True
        printText = "Önceki Dönem Faaliyet Kar Artışı: Bilanço Dönemi Faaliyet Karı Artışı 100%'ün Üzerinde, Karşılaştırma Yapılmayacak: " + "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi) + " " + str(oncekiDonemFaaliyetKarArtisiGecmeDurumu)
        my_logger.info(printText)

    else:
        oncekiDonemFaaliyetKariArtisi = oncekiYilAyniAltiAyDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiDonemFaaliyetKarArtisiGecmeDurumu = (oncekiDonemFaaliyetKariArtisi < bilancoDonemiFaaliyetKariArtisi)
        my_logger.info("Önceki Dönem Faaliyet Kar Artışı:", "{:.2%}".format(oncekiDonemFaaliyetKariArtisi),
          "<?" , "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi) , oncekiDonemFaaliyetKarArtisiGecmeDurumu)



    # Net Kar Hesabı
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("-------------------NET KAR (DÖNEM KARI/ZARARI)--------------------------")
    my_logger.info("")

    bilancoDonemiNetKar = altiAyDegeriHesapla(netKarRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiNetKar = altiAyDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiNetKar = altiAyDegeriHesapla(netKarRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiNetKar = altiAyDegeriHesapla(netKarRow, ucOncekibilancoDonemiColumn)

    bilancoDonemiNetKarPrint = "{:,.0f}".format(bilancoDonemiNetKar).replace(",", ".")
    ikiOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiNetKar).replace(",", ".")
    bilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniAltiAyDegisimiHesapla(netKarRow, bilancoDonemi))

    birOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(birOncekiBilancoDonemiNetKar).replace(",", ".")
    ucOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(ucOncekiBilancoDonemiNetKar).replace(",", ".")
    birOncekiBilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniAltiAyDegisimiHesapla(netKarRow, birOncekiBilancoDonemi))

    netKarTablosu = PrettyTable()
    netKarTablosu.field_names = ["DÖNEM", "NET KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL NET KAR",
                                       "YÜZDE DEĞİŞİM"]
    netKarTablosu.align["NET KAR"] = "r"
    netKarTablosu.align["ÖNCEKİ YIL NET KAR"] = "r"
    netKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    netKarTablosu.add_row([bilancoDonemi, bilancoDonemiNetKarPrint, ikiOncekiBilancoDonemi,
                                 ikiOncekiBilancoDonemiNetKarPrint, bilancoDonemiNetKarDegisimiPrint])
    netKarTablosu.add_row(
        [birOncekiBilancoDonemi, birOncekiBilancoDonemiNetKarPrint, ucOncekiBilancoDonemi,
         ucOncekiBilancoDonemiNetKarPrint, birOncekiBilancoDonemiNetKarDegisimiPrint])

    my_logger.info(netKarTablosu)



    # Brüt Kar Hesabı
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("-------------------BRÜT KAR (BRÜT KAR/ZARAR)--------------------------")
    my_logger.info("")

    bilancoDonemiBrutKar = altiAyDegeriHesapla(brutKarRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiBrutKar = altiAyDegeriHesapla(brutKarRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiBrutKar = altiAyDegeriHesapla(brutKarRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiBrutKar = altiAyDegeriHesapla(brutKarRow, ucOncekibilancoDonemiColumn)

    bilancoDonemiBrutKarPrint = "{:,.0f}".format(bilancoDonemiBrutKar).replace(",", ".")
    ikiOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiBrutKar).replace(",", ".")
    bilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniAltiAyDegisimiHesapla(brutKarRow, bilancoDonemi))

    birOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(birOncekiBilancoDonemiBrutKar).replace(",", ".")
    ucOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ucOncekiBilancoDonemiBrutKar).replace(",", ".")
    birOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniAltiAyDegisimiHesapla(brutKarRow, birOncekiBilancoDonemi))

    brutKarTablosu = PrettyTable()
    brutKarTablosu.field_names = ["DÖNEM", "BRÜT KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL BRÜT KAR", "YÜZDE DEĞİŞİM"]
    brutKarTablosu.align["BRÜT KAR"] = "r"
    brutKarTablosu.align["ÖNCEKİ YIL BRÜT KAR"] = "r"
    brutKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    brutKarTablosu.add_row([bilancoDonemi, bilancoDonemiBrutKarPrint, ikiOncekiBilancoDonemi,
                            ikiOncekiBilancoDonemiBrutKarPrint, bilancoDonemiBrutKarDegisimiPrint])
    brutKarTablosu.add_row(
        [birOncekiBilancoDonemi, birOncekiBilancoDonemiBrutKarPrint, ucOncekiBilancoDonemi,
         ucOncekiBilancoDonemiBrutKarPrint, birOncekiBilancoDonemiBrutKarDegisimiPrint])
    my_logger.info(brutKarTablosu)



    # Gerçek Deger Hesaplama
    my_logger.info("")
    my_logger.info("")
    my_logger.info("----------------GERÇEK DEĞER HESABI--------------------------------------------")

    sermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)
    my_logger.info("Sermaye: %s TL", "{:,.0f}".format(sermaye).replace(",","."))

    anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", bilancoDonemiColumn) / getBilancoDegeri(
        "DÖNEM KARI (ZARARI)", bilancoDonemiColumn)
    my_logger.info("Ana Ortaklık Payı: %s", "{:.3f}".format(anaOrtaklikPayi))

    sonAltiAySatisArtisYuzdesi = oncekiYilAyniAltiAyDegisimiHesapla(hasilatRow, bilancoDonemi)
    birOncekiAltiAySatisArtisYuzdesi = oncekiYilAyniAltiAyDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    sonIkiDonemHasilatToplami = birOncekiBilancoDonemiHasilat + bilancoDonemiHasilat

    my_logger.info("Son 2 Dönem Hasılat Toplamı: %s TL", "{:,.0f}".format(sonIkiDonemHasilatToplami).replace(",","."))

    onumuzdekiIkiDonemHasilatTahmini = (
                (((sonAltiAySatisArtisYuzdesi + birOncekiAltiAySatisArtisYuzdesi) / 2) + 1) * sonIkiDonemHasilatToplami)
    my_logger.info("Önümüzdeki 2 Dönem Hasılat Tahmini: %s TL", "{:,.0f}".format(onumuzdekiIkiDonemHasilatTahmini).replace(",","."))

    # HASILAT TAHMININI MANUEL DEGISTIRMEK ICIN
    #onumuzdekiIkiDonemHasilatTahmini = 5000000000

    ucOncekibilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow, birOncekibilancoDonemiColumn)
    bilancoDonemiFaaliyetKari = altiAyDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)

    onumuzdekiIkiDonemFaaliyetKarMarjiTahmini = (birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) / (
                bilancoDonemiHasilat + birOncekiBilancoDonemiHasilat)
    my_logger.info("Önümüzdeki 2 Dönem Faaliyet Kar Marjı Tahmini: %s TL", "{:.2%}".format(onumuzdekiIkiDonemFaaliyetKarMarjiTahmini))

    faaliyetKariTahmini1 = onumuzdekiIkiDonemHasilatTahmini * onumuzdekiIkiDonemFaaliyetKarMarjiTahmini
    my_logger.info("Faaliyet Kar Tahmini1: %s TL", "{:,.0f}".format(faaliyetKariTahmini1).replace(",","."))

    faaliyetKariTahmini2 = ((birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 2 * 0.3) + (
                bilancoDonemiFaaliyetKari * 4 * 0.5) + \
                           ((
                                        ucOncekibilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 0.2)
    my_logger.info("Faaliyet Kar Tahmini2: %s TL", "{:,.0f}".format(faaliyetKariTahmini2).replace(",","."))

    ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1 + faaliyetKariTahmini2) / 2
    my_logger.info("Ortalama Faaliyet Kari Tahmini: %s TL", "{:,.0f}".format(ortalamaFaaliyetKariTahmini).replace(",","."))

    hisseBasinaOrtalamaKarTahmini = (ortalamaFaaliyetKariTahmini * anaOrtaklikPayi) / sermaye
    my_logger.info("Hisse Başına Ortalama Kar Tahmini: %s TL", format(hisseBasinaOrtalamaKarTahmini, ".2f"))

    likidasyonDegeri = likidasyonDegeriHesapla(bilancoDonemi)
    my_logger.info("Likidasyon Değeri: %s TL", "{:,.0f}".format(likidasyonDegeri).replace(",","."))

    borclar = int(getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", bilancoDonemiColumn))
    my_logger.info("Borçlar: %s TL", "{:,.0f}".format(borclar).replace(",","."))

    bilancoEtkisi = (likidasyonDegeri - borclar) / sermaye * anaOrtaklikPayi
    my_logger.info("Bilanço Etkisi: %s TL", format(bilancoEtkisi, ".2f"))

    gercekDeger = (hisseBasinaOrtalamaKarTahmini * 7) + bilancoEtkisi
    my_logger.info("Gerçek Hisse Değeri: %s TL", format(gercekDeger, ".2f"))

    targetBuy = gercekDeger * 0.66
    my_logger.info("Target Buy: %s TL", format(targetBuy, ".2f"))

    my_logger.info("Bilanço Tarihindeki Hisse Fiyatı: %s TL", format(hisseFiyati, ".2f"))

    gercekFiyataUzaklik = hisseFiyati / targetBuy
    my_logger.info("Gerçek Fiyata Uzaklık Oranı: %s", "{:.2%}".format(gercekFiyataUzaklik))

    gercekFiyataUzaklikTl = hisseFiyati - targetBuy
    my_logger.info("Gerçek Fiyata Uzaklık %s TL:", format(gercekFiyataUzaklikTl, ".2f"))


    # Netpro Hesapla
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("----------------NETPRO KRİTERİ-----------------------------------------------------------------")

    sonIkiDonemFaaliyetKariToplami = bilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari

    birOncekiBilancoDonemiNetKari = altiAyDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    bilancoDonemiNetKari = altiAyDegeriHesapla(netKarRow, bilancoDonemiColumn)

    sonIkiDonemNetKarToplami = bilancoDonemiNetKari + birOncekiBilancoDonemiNetKari

    my_logger.info ("Son 2 Dönem Net Kar Toplamı: %s TL", "{:,.0f}".format(sonIkiDonemNetKarToplami).replace(",", "."))
    my_logger.info ("Son 2 Dönem Faaliyet Karı Toplamı: %s TL", "{:,.0f}".format(sonIkiDonemFaaliyetKariToplami).replace(",", "."))

    fkOrani = hisseFiyati/((sonIkiDonemNetKarToplami*anaOrtaklikPayi)/(sermaye))
    my_logger.info("F/K Oranı: %s", "{:,.2f}".format(fkOrani))

    hbkOrani = sonIkiDonemNetKarToplami/(sermaye)
    my_logger.info ("HBK Oranı: %s", "{:,.2f}".format(hbkOrani))

    netProEstDegeri = ((ortalamaFaaliyetKariTahmini / sonIkiDonemFaaliyetKariToplami) * sonIkiDonemNetKarToplami) * anaOrtaklikPayi
    my_logger.info("NetPro Est Değeri: %s TL", "{:,.0f}".format(netProEstDegeri).replace(",","."))

    piyasaDegeri = (bilancoEtkisi * sermaye * -1) + (hisseFiyati * sermaye)

    my_logger.info("Piyasa Değeri: %s TL", "{:,.0f}".format(piyasaDegeri).replace(",","."))
    my_logger.info("BondYield: %s", "{:.2%}".format(bondYield))

    netProKriteri = (netProEstDegeri / piyasaDegeri) / bondYield
    netProKriteriGecmeDurumu = (netProKriteri > 2)
    my_logger.info("NetPro Kriteri (2'den Büyük Olmalı): %s %s", format(netProKriteri, ".2f"), str(netProKriteriGecmeDurumu))

    minNetProIcinHisseFiyati  = (netProEstDegeri / (1.9 * bondYield) - (bilancoEtkisi * sermaye * -1))/sermaye
    my_logger.info("NetPro 1.9 Olması İçin Olması Gereken Hisse Fiyatı: %s", format(minNetProIcinHisseFiyati, ".2f"))



    # Forward PE Hesapla
    my_logger.info("")
    my_logger.info("")
    my_logger.info("----------------FORWARD PE HESAPLAMA--------------------------------------------------------")

    forwardPeKriteri = (piyasaDegeri) / netProEstDegeri

    forwardPeKriteriGecmeDurumu = (forwardPeKriteri < 4)
    printText = "Forward PE Kriteri (4'ten Küçük Olmalı): " + format(forwardPeKriteri, ".2f") + " " + str(forwardPeKriteriGecmeDurumu)
    my_logger.info(printText)



    # Ek Hesaplama ve Tablolar
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("-------------------EK HESAPLAMA ve TABLOLAR--------------------------")
    my_logger.info("")

    bilancoDonemiBrutKarMarji = bilancoDonemiBrutKar/bilancoDonemiHasilat
    my_logger.info("Bilanço Dönemi Brüt Kar Marjı: %s", "{:.2%}".format(bilancoDonemiBrutKarMarji))

    bilancoDonemiFaaliyetKarMarji = bilancoDonemiFaaliyetKari/bilancoDonemiHasilat
    my_logger.info("Bilanço Dönemi Faaliyet Kar Marjı: %s", "{:.2%}".format(bilancoDonemiFaaliyetKarMarji))

    bilancoDonemiNetKarMarji = bilancoDonemiNetKari/bilancoDonemiHasilat
    my_logger.info("Bilanço Dönemi Net Kar Marjı: %s", "{:.2%}".format(bilancoDonemiNetKarMarji))

    bilancoDonemiOzsermayeKarliligi = bilancoDonemiNetKari/getBilancoDegeri("TOPLAM ÖZKAYNAKLAR", bilancoDonemiColumn)
    my_logger.info("Bilanço Dönemi Özsermaye Karlılığı: %s", "{:.2%}".format(bilancoDonemiOzsermayeKarliligi))





    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("----------------BİLANÇO DOLAR HESABI-------------------------------------")
    my_logger.info("")
    bilancoDonemiOrtalamaDolarKuru = altiAylikBilancoDonemiOrtalamaDolarDegeriBul(bilancoDonemi)

    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , bilancoDonemi , "{:,.2f}".format(bilancoDonemiOrtalamaDolarKuru))

    birOncekiBilancoDonemiOrtalamaDolarKuru = altiAylikBilancoDonemiOrtalamaDolarDegeriBul(birOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , birOncekiBilancoDonemi , "{:,.2f}".format(birOncekiBilancoDonemiOrtalamaDolarKuru))

    ikiOncekiBilancoDonemiOrtalamaDolarKuru = altiAylikBilancoDonemiOrtalamaDolarDegeriBul(ikiOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , ikiOncekiBilancoDonemi ,"{:,.2f}".format(ikiOncekiBilancoDonemiOrtalamaDolarKuru))

    ucOncekiBilancoDonemiOrtalamaDolarKuru = altiAylikBilancoDonemiOrtalamaDolarDegeriBul(ucOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , ucOncekiBilancoDonemi ,"{:,.2f}".format(ucOncekiBilancoDonemiOrtalamaDolarKuru))


    # Bilanço Dönemi Satış(Hasılat) Gelirleri (DOLAR)
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("--------------------HASILAT(SATIŞ) GELİRLERİ (DOLAR)----------------------")
    my_logger.info("")

    bilancoDonemiDolarHasilat = bilancoDonemiHasilat/bilancoDonemiOrtalamaDolarKuru
    birOncekiBilancoDonemiDolarHasilat = birOncekiBilancoDonemiHasilat/birOncekiBilancoDonemiOrtalamaDolarKuru
    ikiOncekiBilancoDonemiDolarHasilat = ikiOncekiBilancoDonemiHasilat/ikiOncekiBilancoDonemiOrtalamaDolarKuru
    ucOncekiBilancoDonemiDolarHasilat = ucOncekiBilancoDonemiHasilat/ucOncekiBilancoDonemiOrtalamaDolarKuru

    bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(bilancoDonemiDolarHasilat).replace(",", ".")
    ikiOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    bilancoDonemiDolarHasilatDegisimi = bilancoDonemiDolarHasilat/ikiOncekiBilancoDonemiDolarHasilat-1
    bilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(bilancoDonemiDolarHasilatDegisimi)

    birOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(birOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    ucOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(ucOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    birOncekiBilancoDonemiDolarHasilatDegisimi = birOncekiBilancoDonemiDolarHasilat/ucOncekiBilancoDonemiDolarHasilat-1
    birOncekiBilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiDolarHasilatDegisimi)

    dolarSatisTablosu = PrettyTable()
    dolarSatisTablosu.field_names = ["DÖNEM", "SATIŞ (USD)", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ (USD)", "YÜZDE DEĞİŞİM"]
    dolarSatisTablosu.align["SATIŞ (USD)"] = "r"
    dolarSatisTablosu.align["ÖNCEKİ YIL SATIŞ (USD)"] = "r"
    dolarSatisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    dolarSatisTablosu.add_row([bilancoDonemi, bilancoDonemiDolarHasilatPrint, ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiDolarHasilatPrint, bilancoDonemiDolarHasilatDegisimiPrint])
    dolarSatisTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiDolarHasilatPrint, ucOncekiBilancoDonemi, ucOncekiBilancoDonemiDolarHasilatPrint, birOncekiBilancoDonemiDolarHasilatDegisimiPrint])
    my_logger.info (dolarSatisTablosu)

    # Bilanço Dönemi (DOLAR) Satış Geliri Artış Kriteri
    bilancoDonemiDolarHasilatGelirArtisi = bilancoDonemiDolarHasilat/ikiOncekiBilancoDonemiDolarHasilat-1
    bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (bilancoDonemiDolarHasilatGelirArtisi > 0.1)

    printText = "Bilanço Dönemi (DOLAR) Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(bilancoDonemiDolarHasilatGelirArtisi) + " >? 10%" + " " + str(bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
    my_logger.info(printText)


    # Önceki Dönem (DOLAR) Hasılat Geliri Artış Kriteri
    oncekiDonemDolarHasilatGelirArtisi = birOncekiBilancoDonemiDolarHasilat/ucOncekiBilancoDonemiDolarHasilat-1
    #
    if (bilancoDonemiDolarHasilatGelirArtisi >= 1):
        printText = "Bilanço Dönemi (DOLAR) Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak."
        my_logger.info (printText)
        oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = True
        printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Geçme Durumu: " + str(oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
        my_logger.info (printText)

    else:
        oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (oncekiDonemDolarHasilatGelirArtisi<bilancoDonemiDolarHasilatGelirArtisi)
        printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Bilanço Döneminden Düşük Mü: " + "{:.2%}".format(oncekiDonemDolarHasilatGelirArtisi) + \
                    " <? " + "{:.2%}".format(bilancoDonemiDolarHasilatGelirArtisi) + " " + str(oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
        my_logger.info(printText)



    # Bilanço Dönemi Faaliyet Karı Gelirleri (DOLAR)
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("--------------------------FAALİYET KARI (DOLAR)-------------------------")
    my_logger.info("")

    bilancoDonemiDolarFaaliyetKari = bilancoDonemiFaaliyetKari/bilancoDonemiOrtalamaDolarKuru
    birOncekiBilancoDonemiDolarFaaliyetKari = birOncekiBilancoDonemiFaaliyetKari/birOncekiBilancoDonemiOrtalamaDolarKuru
    ikiOncekiBilancoDonemiDolarFaaliyetKari = ikiOncekiBilancoDonemiFaaliyetKari/ikiOncekiBilancoDonemiOrtalamaDolarKuru
    ucOncekiBilancoDonemiDolarFaaliyetKari = ucOncekiBilancoDonemiFaaliyetKari/ucOncekiBilancoDonemiOrtalamaDolarKuru

    bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(bilancoDonemiDolarFaaliyetKari).replace(",", ".")
    ikiOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiDolarFaaliyetKari).replace(",",".")
    bilancoDonemiDolarFaaliyetKariDegisimi = bilancoDonemiDolarFaaliyetKari/ikiOncekiBilancoDonemiDolarFaaliyetKari-1
    bilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(bilancoDonemiDolarFaaliyetKariDegisimi)

    birOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(birOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    ucOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(ucOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    birOncekiBilancoDonemiDolarFaaliyetKariDegisimi = birOncekiBilancoDonemiDolarFaaliyetKari/ucOncekiBilancoDonemiDolarFaaliyetKari-1
    birOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiDolarFaaliyetKariDegisimi)

    dolarFaaliyetKariTablosu = PrettyTable()
    dolarFaaliyetKariTablosu.field_names = ["DÖNEM", "FAALİYET KARI (DOLAR)", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI (DOLAR)", "YÜZDE DEĞİŞİM"]
    dolarFaaliyetKariTablosu.align["FAALİYET KARI (DOLAR)"] = "r"
    dolarFaaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI (DOLAR)"] = "r"
    dolarFaaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    dolarFaaliyetKariTablosu.add_row([bilancoDonemi, bilancoDonemiDolarFaaliyetKariPrint, ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiDolarFaaliyetKariPrint, bilancoDonemiDolarFaaliyetKariDegisimiPrint])
    dolarFaaliyetKariTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiDolarFaaliyetKariPrint, ucOncekiBilancoDonemi, ucOncekiBilancoDonemiDolarFaaliyetKariPrint, birOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint])
    my_logger.info (dolarFaaliyetKariTablosu)

    # Bilanço Dönem Faaliyet Kar Artış Kriteri (DOLAR)
    if altiAyDegeriHesapla(netKarRow, bilancoDonemiColumn) < 0:
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Dönem Net Kar Negatif"
        my_logger.info (printText)

    elif (bilancoDonemiDolarFaaliyetKari < 0):
        bilancoDonemiDolarFaaliyetKariArtisi = bilancoDonemiDolarFaaliyetKari/ikiOncekiBilancoDonemiDolarFaaliyetKari -1
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Dönem Dolar Faaliyet Kari Negatif"
        my_logger.info (printText)

    elif (bilancoDonemiDolarFaaliyetKari > 0) and (ikiOncekiBilancoDonemiDolarFaaliyetKari < 0):
        bilancoDonemiDolarFaaliyetKariArtisi = 0
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = True
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Dönem Dolar Faaliyet Karı Negatiften Pozitife Geçmiş"
        my_logger.info (printText)

    else:
        bilancoDonemiDolarFaaliyetKariArtisi = bilancoDonemiDolarFaaliyetKari/ikiOncekiBilancoDonemiDolarFaaliyetKari -1
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = (bilancoDonemiDolarFaaliyetKariArtisi > 0.15)
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + "{:.2%}".format(bilancoDonemiDolarFaaliyetKariArtisi) + " >? 15% " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu)
        my_logger.info(printText)

    # Önceki Dönem Faaliyet Kar Artış Kriteri (DOLAR)

    if bilancoDonemiDolarFaaliyetKariArtisi >= 1:
        birOncekiBilancoDonemiDolarFaaliyetKariArtisi = oncekiDonemFaaliyetKariArtisi/birOncekiBilancoDonemiOrtalamaDolarKuru
        birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = True
        printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: Bilanço Dönemi Artış " + "{:.2%}".format(bilancoDonemiDolarFaaliyetKariArtisi) + \
                    " > 100%, Karşılaştırma Yapılmayacak: " + str(birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)
        my_logger.info(printText)

    else:
        birOncekiBilancoDonemiDolarFaaliyetKariArtisi = oncekiDonemFaaliyetKariArtisi/birOncekiBilancoDonemiOrtalamaDolarKuru
        birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = (birOncekiBilancoDonemiDolarFaaliyetKariArtisi < bilancoDonemiDolarFaaliyetKariArtisi)
        printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: " + "{:.2%}".format(birOncekiBilancoDonemiDolarFaaliyetKariArtisi) + \
                    " <? ", "{:.2%}".format(bilancoDonemiDolarFaaliyetKariArtisi) + str(birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)
        my_logger.info(printText)


    my_logger.debug("")
    my_logger.debug("")
    my_logger.debug("----------------RAPOR DOSYASI OLUŞTURMA/GÜNCELLEME-------------------------------------")

    my_logger.debug (hisseAdi)

    excelRow = ExcelRowClass()

    excelRow.bilancoDonemiHasilat = bilancoDonemiHasilat
    excelRow.oncekiYilAyniCeyrekHasilat = ikiOncekiBilancoDonemiHasilat
    excelRow.bilancoDonemiHasilatDegisimi = bilancoDonemiHasilatDegisimi
    excelRow.birOncekiBilancoDonemiHasilat = birOncekiBilancoDonemiHasilat
    excelRow.besOncekiBilancoDonemiHasilat = ucOncekiBilancoDonemiHasilat
    excelRow.birOncekiBilancoDonemiHasilatDegisimi = birOncekiBilancoDonemiHasilatDegisimi
    excelRow.bilancoDonemiHasilatGelirArtisiGecmeDurumu = bilancoDonemiHasilatGelirArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiHasilatGelirArtisiGecmeDurumu = oncekiDonemHasilatGelirArtisiGecmeDurumu
    excelRow.bilancoDonemiFaaliyetKari = bilancoDonemiFaaliyetKari
    excelRow.oncekiYilAyniCeyrekFaaliyetKari = ikiOncekiBilancoDonemiFaaliyetKari
    excelRow.bilancoDonemiFaaliyetKariDegisimi = bilancoDonemiFaaliyetKariDegisimi
    excelRow.birOncekiBilancoDonemiFaaliyetKari = birOncekiBilancoDonemiFaaliyetKari
    excelRow.besOncekiBilancoDonemiFaaliyetKari = ucOncekiBilancoDonemiFaaliyetKari
    excelRow.oncekiBilancoDonemiFaaliyetKariDegisimi = birOncekiBilancoDonemiFaaliyetKariDegisimi
    excelRow.bilancoDonemiFaaliyetKariArtisiGecmeDurumu = bilancoDonemiFaaliyetKariArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiFaaliyetKarArtisiGecmeDurumu = oncekiDonemFaaliyetKarArtisiGecmeDurumu

    excelRow.bilancoDonemiOrtalamaDolarKuru = bilancoDonemiOrtalamaDolarKuru
    excelRow.bilancoDonemiDolarHasilat = bilancoDonemiDolarHasilat
    excelRow.oncekiYilAyniCeyrekDolarHasilat = ikiOncekiBilancoDonemiDolarHasilat
    excelRow.bilancoDonemiDolarHasilatDegisimi = bilancoDonemiDolarHasilatDegisimi
    excelRow.birOncekiBilancoDonemiDolarHasilatDegisimi = birOncekiBilancoDonemiDolarHasilatDegisimi
    excelRow.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu
    excelRow.bilancoDonemiDolarFaaliyetKari = bilancoDonemiDolarFaaliyetKari
    excelRow.dortOncekiBilancoDonemiDolarFaaliyetKari = ikiOncekiBilancoDonemiDolarFaaliyetKari
    excelRow.bilancoDonemiDolarFaaliyetKariDegisimi = bilancoDonemiDolarFaaliyetKariDegisimi
    excelRow.birOncekiBilancoDonemiDolarFaaliyetKariDegisimi = birOncekiBilancoDonemiDolarFaaliyetKariDegisimi
    excelRow.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu

    excelRow.sermaye = sermaye
    excelRow.anaOrtaklikPayi = anaOrtaklikPayi
    excelRow.sonDortBilancoDonemiHasilatToplami = sonIkiDonemHasilatToplami
    excelRow.onumuzdekiDortBilancoDonemiHasilatTahmini = onumuzdekiIkiDonemHasilatTahmini
    excelRow.onumuzdekiDortBilancoDonemiFaaliyetKarMarjiTahmini = onumuzdekiIkiDonemFaaliyetKarMarjiTahmini
    excelRow.faaliyetKariTahmini1 = faaliyetKariTahmini1
    excelRow.faaliyetKariTahmini2 = faaliyetKariTahmini2
    excelRow.ortalamaFaaliyetKariTahmini = ortalamaFaaliyetKariTahmini
    excelRow.hisseBasinaOrtalamaKarTahmini = hisseBasinaOrtalamaKarTahmini
    excelRow.bilancoEtkisi = bilancoEtkisi
    excelRow.bilancoTarihiHisseFiyati = hisseFiyati
    excelRow.gercekHisseDegeri = gercekDeger
    excelRow.targetBuy = targetBuy
    excelRow.gercekFiyataUzaklik = gercekFiyataUzaklik
    excelRow.netProKriteri = netProKriteri
    excelRow.forwardPeKriteri = forwardPeKriteri

    exportReportExcel(hisseAdi, reportFile, bilancoDonemi, excelRow)

    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info ("-------------------------------- %s ------------------------", hisseAdi)
