import xlrd
from RC1_ExcelRowClass import ExcelRowClass
from RC1_Rapor_Olustur import exportReportExcel
from prettytable import PrettyTable
import logging
import sys
from RC1_BilancoOrtalamaDolarDegeri import ucAylikBilancoDonemiOrtalamaDolarDegeriBul


def runAlgoritma(bilancoDosyasi, bilancoDonemi, bondYield, hisseFiyati, reportFile, logPath, logLevel):

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

        if ceyrek == 3:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + (ceyrek - 3)

    my_logger.debug("Bilanco Donemi: %d", bilancoDonemi)

    birOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(bilancoDonemi)
    my_logger.debug("Bir Onceki Bilanco Donemi: %d", birOncekiBilancoDonemi)

    ikiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(birOncekiBilancoDonemi)
    my_logger.debug("Iki Onceki Bilanco Donemi: %d", ikiOncekiBilancoDonemi)

    ucOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ikiOncekiBilancoDonemi)
    my_logger.debug("Uc Onceki Bilanco Donemi: %d", ucOncekiBilancoDonemi)

    dortOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ucOncekiBilancoDonemi)
    my_logger.debug("Dort Onceki Bilanco Donemi: %d", dortOncekiBilancoDonemi)

    besOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(dortOncekiBilancoDonemi)
    my_logger.debug("Bes Onceki Bilanco Donemi: %d", besOncekiBilancoDonemi)

    altiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(besOncekiBilancoDonemi)
    my_logger.debug("Alti Onceki Bilanco Donemi: %d", altiOncekiBilancoDonemi)

    yediOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(altiOncekiBilancoDonemi)
    my_logger.debug("Yedi Onceki Bilanco Donemi: %d", yediOncekiBilancoDonemi)

    sekizOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(yediOncekiBilancoDonemi)
    my_logger.debug("Sekiz Onceki Bilanco Donemi: %d", sekizOncekiBilancoDonemi)

    wb = xlrd.open_workbook(bilancoDosyasi)
    sheet = wb.sheet_by_index(0)

    def donemColumnFind(col):
        for columni in range(sheet.ncols):
            cell = sheet.cell(0, columni)
            if cell.value == col:
                return columni
        my_logger.info("Uygun Ceyrek Bulunamadi!!!")
        return -1

    bilancoDonemiColumn = donemColumnFind(bilancoDonemi)
    birOncekibilancoDonemiColumn = donemColumnFind(birOncekiBilancoDonemi)
    ikiOncekibilancoDonemiColumn = donemColumnFind(ikiOncekiBilancoDonemi)
    ucOncekibilancoDonemiColumn = donemColumnFind(ucOncekiBilancoDonemi)
    dortOncekibilancoDonemiColumn = donemColumnFind(dortOncekiBilancoDonemi)
    besOncekibilancoDonemiColumn = donemColumnFind(besOncekiBilancoDonemi)
    altiOncekibilancoDonemiColumn = donemColumnFind(altiOncekiBilancoDonemi)
    yediOncekibilancoDonemiColumn = donemColumnFind(yediOncekiBilancoDonemi)

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
    faaliyetKariRow = getBilancoTitleRow("ESAS FAALİYET KARI (ZARARI)");
    # netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı");
    netKarRow = getBilancoTitleRow("DÖNEM KARI (ZARARI)");
    brutKarRow = getBilancoTitleRow("BRÜT KAR (ZARAR)");

    # TODO: Bir önceki çeyrek bilançosunun olmasını garanti edecek şekilde düzenle

    def ceyrekDegeriHesapla(r, c):
        quarter = (sheet.cell_value(0, c)) % (100)
        if (quarter == 3):
            return sheet.cell_value(r, c)
        else:
            if (sheet.cell_value(0,c)-sheet.cell_value(0,(c-1)) == 3):
                return (sheet.cell_value(r, c) - sheet.cell_value(r, (c - 1)))
            else:
                my_logger.info("EKSİK BİLANÇO VAR!")
                return -1


    def oncekiYilAyniCeyrekDegisimiHesapla(row, donem):
        my_logger.debug("fonksiyon: oncekiYilAyniCeyrekDegisimiHesapla")
        donemColumn = donemColumnFind(donem)
        my_logger.debug ("DonemColumn: %s", donemColumn)
        oncekiYilAyniDonemColumn = donemColumnFind(donem - 100)
        my_logger.debug("Onceki Yıl Aynı DonemColumn: %s", oncekiYilAyniDonemColumn)
        my_logger.debug("Row: %d Column: %d",row ,donemColumn)
        ceyrekDegeri = ceyrekDegeriHesapla(row, donemColumn)
        my_logger.debug("Çeyrek Değeri: %d", ceyrekDegeri)
        oncekiCeyrekDegeri = ceyrekDegeriHesapla(row, oncekiYilAyniDonemColumn)
        my_logger.debug ("Önceki Çeyrek Değeri: %d", oncekiCeyrekDegeri)
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

    bilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow,bilancoDonemiColumn)
    birOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow,birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, ucOncekibilancoDonemiColumn)
    dortOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, dortOncekibilancoDonemiColumn)
    besOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, besOncekibilancoDonemiColumn)
    altiOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, altiOncekibilancoDonemiColumn)
    yediOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, yediOncekibilancoDonemiColumn)

    bilancoDonemiHasilatPrint = "{:,.0f}".format(bilancoDonemiHasilat).replace(",", ".")
    dortOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(dortOncekiBilancoDonemiHasilat).replace(",", ".")
    bilancoDonemiHasilatDegisimi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    bilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(bilancoDonemiHasilatDegisimi)

    birOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(birOncekiBilancoDonemiHasilat).replace(",", ".")
    besOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(besOncekiBilancoDonemiHasilat).replace(",", ".")
    birOncekiBilancoDonemiHasilatDegisimi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)
    birOncekiBilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiHasilatDegisimi)

    ikiOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiHasilat).replace(",", ".")
    altiOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(altiOncekiBilancoDonemiHasilat).replace(",", ".")
    ikiOncekiBilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, ikiOncekiBilancoDonemi))

    ucOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(ucOncekiBilancoDonemiHasilat).replace(",", ".")
    yediOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(yediOncekiBilancoDonemiHasilat).replace(",", ".")
    ucOncekiBilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, ucOncekiBilancoDonemi))

    satisTablosu = PrettyTable()
    satisTablosu.field_names = ["ÇEYREK", "SATIŞ", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ", "YÜZDE DEĞİŞİM"]
    satisTablosu.align["SATIŞ"] = "r"
    satisTablosu.align["ÖNCEKİ YIL SATIŞ"] = "r"
    satisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    satisTablosu.add_row([bilancoDonemi, bilancoDonemiHasilatPrint, dortOncekiBilancoDonemi, dortOncekiBilancoDonemiHasilatPrint, bilancoDonemiHasilatDegisimiPrint])
    satisTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiHasilatPrint, besOncekiBilancoDonemi, besOncekiBilancoDonemiHasilatPrint, birOncekiBilancoDonemiHasilatDegisimiPrint])
    satisTablosu.add_row([ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiHasilatPrint, altiOncekiBilancoDonemi, altiOncekiBilancoDonemiHasilatPrint, ikiOncekiBilancoDonemiHasilatDegisimiPrint])
    satisTablosu.add_row([ucOncekiBilancoDonemi, ucOncekiBilancoDonemiHasilatPrint, yediOncekiBilancoDonemi,yediOncekiBilancoDonemiHasilatPrint, ucOncekiBilancoDonemiHasilatDegisimiPrint])
    my_logger.info(satisTablosu)

    # Bilanço Dönemi Saış Geliri Artış Kriteri
    bilancoDonemiHasilatGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    bilancoDonemiHasilatGelirArtisiGecmeDurumu = (bilancoDonemiHasilatGelirArtisi > 0.1)
    printText = "Bilanço Dönemi Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(bilancoDonemiHasilatGelirArtisi) + " >? 10% " + " " + str(bilancoDonemiHasilatGelirArtisiGecmeDurumu)
    my_logger.info(printText)

    # Önceki Dönem Hasılat Geliri Artış Kriteri
    oncekiDonemHasilatGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

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

    bilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow,bilancoDonemiColumn)
    birOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow,birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)
    dortOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, dortOncekibilancoDonemiColumn)
    besOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, besOncekibilancoDonemiColumn)
    altiOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, altiOncekibilancoDonemiColumn)
    yediOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, yediOncekibilancoDonemiColumn)

    bilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(bilancoDonemiFaaliyetKari).replace(",", ".")
    dortOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(dortOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    bilancoDonemiFaaliyetKariDegisimi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
    bilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(bilancoDonemiFaaliyetKariDegisimi)

    birOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(birOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    besOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(besOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    birOncekiBilancoDonemiFaaliyetKariDegisimi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
    birOncekiBilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiFaaliyetKariDegisimi)

    ikiOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    altiOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(altiOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    ikiOncekiBilancoDonemiFaaliyetKariDegisimi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, ikiOncekiBilancoDonemi)
    ikiOncekiBilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(ikiOncekiBilancoDonemiFaaliyetKariDegisimi)

    ucOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(ucOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    yediOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(yediOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    ucOncekiBilancoDonemiFaaliyetKariDegisimiPrint = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, ucOncekiBilancoDonemi)
    ucOncekiBilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(ucOncekiBilancoDonemiFaaliyetKariDegisimiPrint)

    faaliyetKariTablosu = PrettyTable()
    faaliyetKariTablosu.field_names = ["ÇEYREK", "FAALİYET KARI", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI", "YÜZDE DEĞİŞİM"]
    faaliyetKariTablosu.align["FAALİYET KARI"] = "r"
    faaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI"] = "r"
    faaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    faaliyetKariTablosu.add_row([bilancoDonemi, bilancoDonemiFaaliyetKariPrint, dortOncekiBilancoDonemi, dortOncekiBilancoDonemiFaaliyetKariPrint, bilancoDonemiFaaliyetKariDegisimiPrint])
    faaliyetKariTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiFaaliyetKariPrint, besOncekiBilancoDonemi, besOncekiBilancoDonemiFaaliyetKariPrint, birOncekiBilancoDonemiFaaliyetKariDegisimiPrint])
    faaliyetKariTablosu.add_row([ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiFaaliyetKariPrint, altiOncekiBilancoDonemi, altiOncekiBilancoDonemiFaaliyetKariPrint, ikiOncekiBilancoDonemiFaaliyetKariDegisimiPrint])
    faaliyetKariTablosu.add_row([ucOncekiBilancoDonemi, ucOncekiBilancoDonemiFaaliyetKariPrint, yediOncekiBilancoDonemi,yediOncekiBilancoDonemiFaaliyetKariPrint, ucOncekiBilancoDonemiFaaliyetKariDegisimiPrint])
    my_logger.info(faaliyetKariTablosu)


    # Bilanço Dönemi Faaliyet Kar Artış Kriteri
    if ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn) < 0:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = False
        my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Çeyrek Net Kar Negatif", str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu))

    elif ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) < 0:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = False
        my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Ceyrek Faaliyet Kari Negatif", str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu))

    elif ((ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) > 0) and (ceyrekDegeriHesapla(faaliyetKariRow, dortOncekibilancoDonemiColumn)) < 0):
        bilancoDonemiFaaliyetKariArtisi = 0
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = True
        my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Çeyrek Faaliyet Karı Negatiften Pozitife Geçmiş", str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu))

    else:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = (bilancoDonemiFaaliyetKariArtisi > 0.15)
        printText = "Bilanço Dönemi Faaliyet Karı Artışı:" + "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi) + " >? 15% " + str(bilancoDonemiFaaliyetKariArtisiGecmeDurumu)
        my_logger.info(printText)

    # Önceki Dönem Faaliyet Kar Artış Kriteri

    if bilancoDonemiFaaliyetKariArtisi >= 1:
        oncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiCeyrekFaaliyetKarArtisiGecmeDurumu = True
        printText = "Önceki Dönem Faaliyet Kar Artışı: Bilanço Dönemi Faaliyet Karı Artışı 100%'ün Üzerinde, Karşılaştırma Yapılmayacak: " + "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi) + " " + str(oncekiCeyrekFaaliyetKarArtisiGecmeDurumu)
        my_logger.info(printText)

    else:
        oncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiCeyrekFaaliyetKarArtisiGecmeDurumu = (oncekiCeyrekFaaliyetKariArtisi < bilancoDonemiFaaliyetKariArtisi)
        printText = "Önceki Dönem Faaliyet Kar Artışı:" + "{:.2%}".format(oncekiCeyrekFaaliyetKariArtisi) + " < ? " + "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi) + str(oncekiCeyrekFaaliyetKarArtisiGecmeDurumu)
        my_logger.info(printText)



    # Net Kar Hesabı
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("-------------------NET KAR (DÖNEM KARI/ZARARI)--------------------------")
    my_logger.info("")

    bilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, ucOncekibilancoDonemiColumn)
    dortOncekiBilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, dortOncekibilancoDonemiColumn)
    besOncekiBilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, besOncekibilancoDonemiColumn)
    altiOncekiBilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, altiOncekibilancoDonemiColumn)
    yediOncekiBilancoDonemiNetKar = ceyrekDegeriHesapla(netKarRow, yediOncekibilancoDonemiColumn)

    bilancoDonemiNetKarPrint = "{:,.0f}".format(bilancoDonemiNetKar).replace(",", ".")
    dortOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(dortOncekiBilancoDonemiNetKar).replace(",", ".")
    bilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(netKarRow, bilancoDonemi))

    birOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(birOncekiBilancoDonemiNetKar).replace(",", ".")
    besOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(besOncekiBilancoDonemiNetKar).replace(",", ".")
    birOncekiBilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(netKarRow, birOncekiBilancoDonemi))

    ikiOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiNetKar).replace(",", ".")
    altiOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(altiOncekiBilancoDonemiNetKar).replace(",", ".")
    ikiOncekiBilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(netKarRow, ikiOncekiBilancoDonemi))

    ucOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(ucOncekiBilancoDonemiNetKar).replace(",", ".")
    yediOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(yediOncekiBilancoDonemiNetKar).replace(",",".")
    ucOncekiBilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(netKarRow, ucOncekiBilancoDonemi))

    netKarTablosu = PrettyTable()
    netKarTablosu.field_names = ["ÇEYREK", "NET KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL NET KAR",
                                       "YÜZDE DEĞİŞİM"]
    netKarTablosu.align["NET KAR"] = "r"
    netKarTablosu.align["ÖNCEKİ YIL NET KAR"] = "r"
    netKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    netKarTablosu.add_row([bilancoDonemi, bilancoDonemiNetKarPrint, dortOncekiBilancoDonemi,
                                 dortOncekiBilancoDonemiNetKarPrint, bilancoDonemiNetKarDegisimiPrint])
    netKarTablosu.add_row(
        [birOncekiBilancoDonemi, birOncekiBilancoDonemiNetKarPrint, besOncekiBilancoDonemi,
         besOncekiBilancoDonemiNetKarPrint, birOncekiBilancoDonemiNetKarDegisimiPrint])
    netKarTablosu.add_row(
        [ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiNetKarPrint, altiOncekiBilancoDonemi,
         altiOncekiBilancoDonemiNetKarPrint, ikiOncekiBilancoDonemiNetKarDegisimiPrint])
    netKarTablosu.add_row(
        [ucOncekiBilancoDonemi, ucOncekiBilancoDonemiNetKarPrint, yediOncekiBilancoDonemi,
         yediOncekiBilancoDonemiNetKarPrint, ucOncekiBilancoDonemiNetKarDegisimiPrint])
    my_logger.info(netKarTablosu)



    # Brüt Kar Hesabı
    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("-------------------BRÜT KAR (BRÜT KAR/ZARAR)--------------------------")
    my_logger.info("")

    bilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, ucOncekibilancoDonemiColumn)
    dortOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, dortOncekibilancoDonemiColumn)
    besOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, besOncekibilancoDonemiColumn)
    altiOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, altiOncekibilancoDonemiColumn)
    yediOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, yediOncekibilancoDonemiColumn)

    bilancoDonemiBrutKarPrint = "{:,.0f}".format(bilancoDonemiBrutKar).replace(",", ".")
    dortOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(dortOncekiBilancoDonemiBrutKar).replace(",", ".")
    bilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, bilancoDonemi))

    birOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(birOncekiBilancoDonemiBrutKar).replace(",", ".")
    besOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(besOncekiBilancoDonemiBrutKar).replace(",", ".")
    birOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, birOncekiBilancoDonemi))

    ikiOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiBrutKar).replace(",", ".")
    altiOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(altiOncekiBilancoDonemiBrutKar).replace(",", ".")
    ikiOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, ikiOncekiBilancoDonemi))

    ucOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ucOncekiBilancoDonemiBrutKar).replace(",", ".")
    yediOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(yediOncekiBilancoDonemiBrutKar).replace(",", ".")
    ucOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, ucOncekiBilancoDonemi))

    brutKarTablosu = PrettyTable()
    brutKarTablosu.field_names = ["ÇEYREK", "BRÜT KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL BRÜT KAR", "YÜZDE DEĞİŞİM"]
    brutKarTablosu.align["BRÜT KAR"] = "r"
    brutKarTablosu.align["ÖNCEKİ YIL BRÜT KAR"] = "r"
    brutKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    brutKarTablosu.add_row([bilancoDonemi, bilancoDonemiBrutKarPrint, dortOncekiBilancoDonemi,
                            dortOncekiBilancoDonemiBrutKarPrint, bilancoDonemiBrutKarDegisimiPrint])
    brutKarTablosu.add_row(
        [birOncekiBilancoDonemi, birOncekiBilancoDonemiBrutKarPrint, besOncekiBilancoDonemi,
         besOncekiBilancoDonemiBrutKarPrint, birOncekiBilancoDonemiBrutKarDegisimiPrint])
    brutKarTablosu.add_row(
        [ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiBrutKarPrint, altiOncekiBilancoDonemi,
         altiOncekiBilancoDonemiBrutKarPrint, ikiOncekiBilancoDonemiBrutKarDegisimiPrint])
    brutKarTablosu.add_row(
        [ucOncekiBilancoDonemi, ucOncekiBilancoDonemiBrutKarPrint, yediOncekiBilancoDonemi,
         yediOncekiBilancoDonemiBrutKarPrint, ucOncekiBilancoDonemiBrutKarDegisimiPrint])
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

    sonCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    birOncekiCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    sonDortCeyrekHasilatToplami = ucOncekiBilancoDonemiHasilat + ikiOncekiBilancoDonemiHasilat + birOncekiBilancoDonemiHasilat + bilancoDonemiHasilat

    my_logger.info("Son 4 Çeyrek Hasılat Toplamı: %s TL", "{:,.0f}".format(sonDortCeyrekHasilatToplami).replace(",","."))

    onumuzdekiDortCeyrekHasilatTahmini = (
                (((sonCeyrekSatisArtisYuzdesi + birOncekiCeyrekSatisArtisYuzdesi) / 2) + 1) * sonDortCeyrekHasilatToplami)

    # HASILAT TAHMININI MANUEL DEGISTIRMEK ICIN
    # onumuzdekiDortCeyrekHasilatTahmini = 4000000000

    my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini: %s TL", "{:,.0f}".format(onumuzdekiDortCeyrekHasilatTahmini).replace(",","."))

    ucOncekibilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, birOncekibilancoDonemiColumn)
    bilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)

    onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) / (
                bilancoDonemiHasilat + birOncekiBilancoDonemiHasilat)
    my_logger.info("Önümüzdeki 4 Çeyrek Faaliyet Kar Marjı Tahmini: %s TL", "{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))

    faaliyetKariTahmini1 = onumuzdekiDortCeyrekHasilatTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
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

    sonDortDonemFaaliyetKariToplami = bilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + ucOncekibilancoDonemiFaaliyetKari

    ucOncekibilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    bilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn)

    sonDortDonemNetKarToplami = bilancoDonemiNetKari + birOncekiBilancoDonemiNetKari + ikiOncekiBilancoDonemiNetKari + ucOncekibilancoDonemiNetKari

    my_logger.info ("Son 4 Dönem Net Kar Toplamı: %s TL", "{:,.0f}".format(sonDortDonemNetKarToplami).replace(",", "."))
    my_logger.info ("Son 4 Dönem Faaliyet Karı Toplamı: %s TL", "{:,.0f}".format(sonDortDonemFaaliyetKariToplami).replace(",", "."))

    fkOrani = hisseFiyati/((sonDortDonemNetKarToplami*anaOrtaklikPayi)/(sermaye))
    my_logger.info("F/K Oranı: %s", "{:,.2f}".format(fkOrani))

    hbkOrani = sonDortDonemNetKarToplami/(sermaye)
    my_logger.info ("HBK Oranı: %s", "{:,.2f}".format(hbkOrani))

    netProEstDegeri = ((ortalamaFaaliyetKariTahmini / sonDortDonemFaaliyetKariToplami) * sonDortDonemNetKarToplami) * anaOrtaklikPayi
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

    if (bilancoDonemiHasilat != 0):
        bilancoDonemiBrutKarMarji = bilancoDonemiBrutKar / bilancoDonemiHasilat
        bilancoDonemiFaaliyetKarMarji = bilancoDonemiFaaliyetKari / bilancoDonemiHasilat
        bilancoDonemiNetKarMarji = bilancoDonemiNetKari / bilancoDonemiHasilat
    else:
        bilancoDonemiBrutKarMarji = 0
        bilancoDonemiFaaliyetKarMarji = 0
        bilancoDonemiNetKarMarji = 0

    my_logger.info("Bilanço Dönemi Brüt Kar Marjı: %s", "{:.2%}".format(bilancoDonemiBrutKarMarji))
    my_logger.info("Bilanço Dönemi Faaliyet Kar Marjı: %s", "{:.2%}".format(bilancoDonemiFaaliyetKarMarji))
    my_logger.info("Bilanço Dönemi Net Kar Marjı: %s", "{:.2%}".format(bilancoDonemiNetKarMarji))

    bilancoDonemiOzsermayeKarliligi = bilancoDonemiNetKari/getBilancoDegeri("TOPLAM ÖZKAYNAKLAR", bilancoDonemiColumn)
    my_logger.info("Bilanço Dönemi Özsermaye Karlılığı: %s", "{:.2%}".format(bilancoDonemiOzsermayeKarliligi))



    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info("----------------BİLANÇO DOLAR HESABI-------------------------------------")
    my_logger.info("")
    bilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(bilancoDonemi)

    my_logger.info ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , bilancoDonemi , "{:,.2f}".format(bilancoDonemiOrtalamaDolarKuru))

    birOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(birOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , birOncekiBilancoDonemi , "{:,.2f}".format(birOncekiBilancoDonemiOrtalamaDolarKuru))

    ikiOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(ikiOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , ikiOncekiBilancoDonemi ,"{:,.2f}".format(ikiOncekiBilancoDonemiOrtalamaDolarKuru))

    ucOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(ucOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , ucOncekiBilancoDonemi ,"{:,.2f}".format(ucOncekiBilancoDonemiOrtalamaDolarKuru))

    dortOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(dortOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , dortOncekiBilancoDonemi ,"{:,.2f}".format(dortOncekiBilancoDonemiOrtalamaDolarKuru))

    besOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(besOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL", besOncekiBilancoDonemi, "{:,.2f}".format(besOncekiBilancoDonemiOrtalamaDolarKuru))

    altiOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(altiOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , altiOncekiBilancoDonemi, "{:,.2f}".format(altiOncekiBilancoDonemiOrtalamaDolarKuru))

    yediOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(yediOncekiBilancoDonemi)
    my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL", yediOncekiBilancoDonemi, "{:,.2f}".format(yediOncekiBilancoDonemiOrtalamaDolarKuru))



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
    oncekiYilAyniCeyrekDolarHasilat = dortOncekiBilancoDonemiHasilat/dortOncekiBilancoDonemiOrtalamaDolarKuru
    besOncekiBilancoDonemiDolarHasilat = besOncekiBilancoDonemiHasilat/besOncekiBilancoDonemiOrtalamaDolarKuru
    altiOncekiBilancoDonemiDolarHasilat = altiOncekiBilancoDonemiHasilat/altiOncekiBilancoDonemiOrtalamaDolarKuru
    yediOncekiBilancoDonemiDolarHasilat = yediOncekiBilancoDonemiHasilat/yediOncekiBilancoDonemiOrtalamaDolarKuru

    bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(bilancoDonemiDolarHasilat).replace(",", ".")
    dortOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(oncekiYilAyniCeyrekDolarHasilat).replace(",", ".")
    bilancoDonemiDolarHasilatDegisimi = bilancoDonemiDolarHasilat/oncekiYilAyniCeyrekDolarHasilat-1
    bilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(bilancoDonemiDolarHasilatDegisimi)

    birOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(birOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    besOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(besOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    birOncekiBilancoDonemiDolarHasilatDegisimi = birOncekiBilancoDonemiDolarHasilat/besOncekiBilancoDonemiDolarHasilat-1
    birOncekiBilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiDolarHasilatDegisimi)

    ikiOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    altiOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(altiOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    ikiOncekiBilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(ikiOncekiBilancoDonemiDolarHasilat/altiOncekiBilancoDonemiDolarHasilat-1)

    ucOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(ucOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    yediOncekiBilancoDonemiDolarHasilatPrint = "{:,.0f}".format(yediOncekiBilancoDonemiDolarHasilat).replace(",", ".")
    ucOncekiBilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(ucOncekiBilancoDonemiDolarHasilat/yediOncekiBilancoDonemiDolarHasilat-1)

    dolarSatisTablosu = PrettyTable()
    dolarSatisTablosu.field_names = ["ÇEYREK", "SATIŞ (USD)", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ (USD)", "YÜZDE DEĞİŞİM"]
    dolarSatisTablosu.align["SATIŞ (USD)"] = "r"
    dolarSatisTablosu.align["ÖNCEKİ YIL SATIŞ (USD)"] = "r"
    dolarSatisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    dolarSatisTablosu.add_row([bilancoDonemi, bilancoDonemiDolarHasilatPrint, dortOncekiBilancoDonemi, dortOncekiBilancoDonemiDolarHasilatPrint, bilancoDonemiDolarHasilatDegisimiPrint])
    dolarSatisTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiDolarHasilatPrint, besOncekiBilancoDonemi, besOncekiBilancoDonemiDolarHasilatPrint, birOncekiBilancoDonemiDolarHasilatDegisimiPrint])
    dolarSatisTablosu.add_row([ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiDolarHasilatPrint, altiOncekiBilancoDonemi, altiOncekiBilancoDonemiDolarHasilatPrint, ikiOncekiBilancoDonemiDolarHasilatDegisimiPrint])
    dolarSatisTablosu.add_row([ucOncekiBilancoDonemi, ucOncekiBilancoDonemiDolarHasilatPrint, yediOncekiBilancoDonemi,yediOncekiBilancoDonemiDolarHasilatPrint, ucOncekiBilancoDonemiDolarHasilatDegisimiPrint])
    my_logger.info (dolarSatisTablosu)

    # Bilanço Dönemi (DOLAR) Satış Geliri Artış Kriteri
    bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (bilancoDonemiDolarHasilatDegisimi > 0.1)

    printText = "Bilanço Dönemi (DOLAR) Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(bilancoDonemiDolarHasilatDegisimi) + " >? 10%" + " " + str(bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
    my_logger.info(printText)



    # Önceki Dönem (DOLAR) Hasılat Geliri Artış Kriteri
    #
    if (bilancoDonemiDolarHasilatDegisimi >= 1):
        printText = "Bilanço Dönemi (DOLAR) Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak."
        my_logger.info (printText)
        oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = True
        printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Geçme Durumu: " + str(oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
        my_logger.info (printText)

    else:
        oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (birOncekiBilancoDonemiDolarHasilatDegisimi<bilancoDonemiDolarHasilatDegisimi)
        printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Bilanço Döneminden Düşük Mü: " + "{:.2%}".format(birOncekiBilancoDonemiDolarHasilatDegisimi) + \
                    " <? " + "{:.2%}".format(bilancoDonemiDolarHasilatDegisimi) + " " + str(oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
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
    dortOncekiBilancoDonemiDolarFaaliyetKari = dortOncekiBilancoDonemiFaaliyetKari/dortOncekiBilancoDonemiOrtalamaDolarKuru
    besOncekiBilancoDonemiDolarFaaliyetKari = besOncekiBilancoDonemiFaaliyetKari/besOncekiBilancoDonemiOrtalamaDolarKuru
    altiOncekiBilancoDonemiDolarFaaliyetKari = altiOncekiBilancoDonemiFaaliyetKari/altiOncekiBilancoDonemiOrtalamaDolarKuru
    yediOncekiBilancoDonemiDolarFaaliyetKari = yediOncekiBilancoDonemiFaaliyetKari/yediOncekiBilancoDonemiOrtalamaDolarKuru

    bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(bilancoDonemiDolarFaaliyetKari).replace(",", ".")
    dortOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(dortOncekiBilancoDonemiDolarFaaliyetKari).replace(",",".")
    bilancoDonemiDolarFaaliyetKariDegisimi = bilancoDonemiDolarFaaliyetKari/dortOncekiBilancoDonemiDolarFaaliyetKari-1
    bilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(bilancoDonemiDolarFaaliyetKariDegisimi)

    birOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(birOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    besOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(besOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    birOncekiBilancoDonemiDolarFaaliyetKariDegisimi = birOncekiBilancoDonemiDolarFaaliyetKari/besOncekiBilancoDonemiDolarFaaliyetKari-1
    birOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(birOncekiBilancoDonemiDolarFaaliyetKariDegisimi)

    ikiOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    altiOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(altiOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    ikiOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(ikiOncekiBilancoDonemiDolarFaaliyetKari/altiOncekiBilancoDonemiDolarFaaliyetKari-1)

    ucOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(ucOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    yediOncekiBilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(yediOncekiBilancoDonemiDolarFaaliyetKari).replace(",", ".")
    ucOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(ucOncekiBilancoDonemiDolarFaaliyetKari/yediOncekiBilancoDonemiDolarFaaliyetKari-1)

    dolarFaaliyetKariTablosu = PrettyTable()
    dolarFaaliyetKariTablosu.field_names = ["ÇEYREK", "FAALİYET KARI (DOLAR)", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI (DOLAR)", "YÜZDE DEĞİŞİM"]
    dolarFaaliyetKariTablosu.align["FAALİYET KARI (DOLAR)"] = "r"
    dolarFaaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI (DOLAR)"] = "r"
    dolarFaaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    dolarFaaliyetKariTablosu.add_row([bilancoDonemi, bilancoDonemiDolarFaaliyetKariPrint, dortOncekiBilancoDonemi, dortOncekiBilancoDonemiDolarFaaliyetKariPrint, bilancoDonemiDolarFaaliyetKariDegisimiPrint])
    dolarFaaliyetKariTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiDolarFaaliyetKariPrint, besOncekiBilancoDonemi,besOncekiBilancoDonemiDolarFaaliyetKariPrint, birOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint])
    dolarFaaliyetKariTablosu.add_row([ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiDolarFaaliyetKariPrint, altiOncekiBilancoDonemi, altiOncekiBilancoDonemiDolarFaaliyetKariPrint, ikiOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint])
    dolarFaaliyetKariTablosu.add_row([ucOncekiBilancoDonemi, ucOncekiBilancoDonemiDolarFaaliyetKariPrint, yediOncekiBilancoDonemi, yediOncekiBilancoDonemiDolarFaaliyetKariPrint, ucOncekiBilancoDonemiDolarFaaliyetKariDegisimiPrint])
    my_logger.info (dolarFaaliyetKariTablosu)

    # Bilanço Dönem Faaliyet Kar Artış Kriteri (DOLAR)
    if ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn) < 0:
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Çeyrek Net Kar Negatif"
        my_logger.info (printText)

    elif (bilancoDonemiDolarFaaliyetKari < 0):
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Ceyrek Dolar Faaliyet Kari Negatif"
        my_logger.info (printText)

    elif (bilancoDonemiDolarFaaliyetKari > 0) and (dortOncekiBilancoDonemiDolarFaaliyetKari < 0):
        bilancoDonemiDolarFaaliyetKariArtisi = 0
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = True
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Çeyrek Dolar Faaliyet Karı Negatiften Pozitife Geçmiş"
        my_logger.info (printText)

    else:
        bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = (bilancoDonemiDolarFaaliyetKariDegisimi > 0.15)
        printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + "{:.2%}".format(bilancoDonemiDolarFaaliyetKariDegisimi) + " >? 15% " + str(bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu)
        my_logger.info(printText)

    # Önceki Dönem Faaliyet Kar Artış Kriteri (DOLAR)

    if bilancoDonemiDolarFaaliyetKariDegisimi >= 1:
        birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = True
        printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: Bilanço Dönemi Artış " + "{:.2%}".format(bilancoDonemiDolarFaaliyetKariDegisimi) + \
                    " > 100%, Karşılaştırma Yapılmayacak: " + str(birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)
        my_logger.info(printText)


    else:
        birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = (birOncekiBilancoDonemiDolarFaaliyetKariDegisimi < bilancoDonemiDolarFaaliyetKariDegisimi)
        printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: " + "{:.2%}".format(birOncekiBilancoDonemiDolarFaaliyetKariDegisimi) + \
                    " <? " + "{:.2%}".format(bilancoDonemiDolarFaaliyetKariDegisimi) + " " + str(birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)
        my_logger.info(printText)


    my_logger.debug("")
    my_logger.debug("")
    my_logger.debug("----------------RAPOR DOSYASI OLUŞTURMA/GÜNCELLEME-------------------------------------")

    my_logger.debug (hisseAdi)

    excelRow = ExcelRowClass()

    excelRow.bilancoDonemiHasilat = bilancoDonemiHasilat
    excelRow.oncekiYilAyniCeyrekHasilat = dortOncekiBilancoDonemiHasilat
    excelRow.bilancoDonemiHasilatDegisimi = bilancoDonemiHasilatDegisimi
    excelRow.birOncekiBilancoDonemiHasilat = birOncekiBilancoDonemiHasilat
    excelRow.besOncekiBilancoDonemiHasilat = besOncekiBilancoDonemiHasilat
    excelRow.birOncekiBilancoDonemiHasilatDegisimi = birOncekiBilancoDonemiHasilatDegisimi
    excelRow.bilancoDonemiHasilatGelirArtisiGecmeDurumu = bilancoDonemiHasilatGelirArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiHasilatGelirArtisiGecmeDurumu = oncekiDonemHasilatGelirArtisiGecmeDurumu
    excelRow.bilancoDonemiFaaliyetKari = bilancoDonemiFaaliyetKari
    excelRow.oncekiYilAyniCeyrekFaaliyetKari = dortOncekiBilancoDonemiFaaliyetKari
    excelRow.bilancoDonemiFaaliyetKariDegisimi = bilancoDonemiFaaliyetKariDegisimi
    excelRow.birOncekiBilancoDonemiFaaliyetKari = birOncekiBilancoDonemiFaaliyetKari
    excelRow.besOncekiBilancoDonemiFaaliyetKari = besOncekiBilancoDonemiFaaliyetKari
    excelRow.oncekiBilancoDonemiFaaliyetKariDegisimi = birOncekiBilancoDonemiFaaliyetKariDegisimi
    excelRow.bilancoDonemiFaaliyetKariArtisiGecmeDurumu = bilancoDonemiFaaliyetKariArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiFaaliyetKarArtisiGecmeDurumu = oncekiCeyrekFaaliyetKarArtisiGecmeDurumu

    excelRow.bilancoDonemiOrtalamaDolarKuru = bilancoDonemiOrtalamaDolarKuru
    excelRow.bilancoDonemiDolarHasilat = bilancoDonemiDolarHasilat
    excelRow.oncekiYilAyniCeyrekDolarHasilat = oncekiYilAyniCeyrekDolarHasilat
    excelRow.bilancoDonemiDolarHasilatDegisimi = bilancoDonemiDolarHasilatDegisimi
    excelRow.birOncekiBilancoDonemiDolarHasilatDegisimi = birOncekiBilancoDonemiDolarHasilatDegisimi
    excelRow.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu
    excelRow.bilancoDonemiDolarFaaliyetKari = bilancoDonemiDolarFaaliyetKari
    excelRow.dortOncekiBilancoDonemiDolarFaaliyetKari = dortOncekiBilancoDonemiDolarFaaliyetKari
    excelRow.bilancoDonemiDolarFaaliyetKariDegisimi = bilancoDonemiDolarFaaliyetKariDegisimi
    excelRow.birOncekiBilancoDonemiDolarFaaliyetKariDegisimi = birOncekiBilancoDonemiDolarFaaliyetKariDegisimi
    excelRow.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu
    excelRow.oncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu

    excelRow.sermaye = sermaye
    excelRow.anaOrtaklikPayi = anaOrtaklikPayi
    excelRow.sonDortBilancoDonemiHasilatToplami = sonDortCeyrekHasilatToplami
    excelRow.onumuzdekiDortBilancoDonemiHasilatTahmini = onumuzdekiDortCeyrekHasilatTahmini
    excelRow.onumuzdekiDortBilancoDonemiFaaliyetKarMarjiTahmini = onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
    excelRow.faaliyetKariTahmini1 = faaliyetKariTahmini1
    excelRow.faaliyetKariTahmini2 = faaliyetKariTahmini2
    excelRow.ortalamaFaaliyetKariTahmini = ortalamaFaaliyetKariTahmini
    excelRow.hisseBasinaOrtalamaKarTahmini = hisseBasinaOrtalamaKarTahmini
    excelRow.bilancoEtkisi = bilancoEtkisi
    excelRow.bilancoTarihiHisseFiyati = hisseFiyati
    excelRow.gercekHisseDegeri = gercekDeger
    excelRow.targetBuy = targetBuy
    excelRow.gercekFiyataUzaklik = gercekFiyataUzaklik
    excelRow.fkOrani = fkOrani
    excelRow.hbkOrani = hbkOrani

    excelRow.netProKriteri = netProKriteri
    excelRow.forwardPeKriteri = forwardPeKriteri

    exportReportExcel(hisseAdi, reportFile, bilancoDonemi, excelRow)

    my_logger.info("")
    my_logger.info("")
    my_logger.info("")
    my_logger.info ("-------------------------------- %s ------------------------", hisseAdi)

    my_logger.removeHandler(output_file_handler)
    my_logger.removeHandler(stdout_handler)
