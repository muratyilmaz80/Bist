import xlrd
from RC1_ExcelRowClass import ExcelRowClass
from RC1_Rapor_Olustur import exportReportExcel
from prettytable import PrettyTable
import logging
import sys
from RC1_BilancoOrtalamaDolarDegeri import ucAylikBilancoDonemiOrtalamaDolarDegeriBul


def runAlgoritma(bilancoDosyasi, bilancoDonemi, bondYield, hisseFiyati, reportFile):

    hisseAdiTemp = bilancoDosyasi[47:]
    hisseAdi = hisseAdiTemp[:-5]

    print ("--------------------------------", hisseAdi, "--------------------------------")

    def birOncekiBilancoDoneminiHesapla(dnm):
        yil = int(dnm / 100)
        ceyrek = int(dnm % 100)

        if ceyrek == 3:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + (ceyrek - 3)

    logging.debug("Bilanco Donemi: %d", bilancoDonemi)

    birOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(bilancoDonemi)
    logging.debug("Bir Onceki Bilanco Donemi: %d", birOncekiBilancoDonemi)

    ikiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(birOncekiBilancoDonemi)
    logging.debug("Iki Onceki Bilanco Donemi: %d", ikiOncekiBilancoDonemi)

    ucOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ikiOncekiBilancoDonemi)
    logging.debug("Uc Onceki Bilanco Donemi: %d", ucOncekiBilancoDonemi)

    dortOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ucOncekiBilancoDonemi)
    logging.debug("Dort Onceki Bilanco Donemi: %d", dortOncekiBilancoDonemi)

    besOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(dortOncekiBilancoDonemi)
    logging.debug("Bes Onceki Bilanco Donemi: %d", besOncekiBilancoDonemi)

    altiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(besOncekiBilancoDonemi)
    logging.debug("Alti Onceki Bilanco Donemi: %d", altiOncekiBilancoDonemi)

    yediOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(altiOncekiBilancoDonemi)
    logging.debug("Yedi Onceki Bilanco Donemi: %d", yediOncekiBilancoDonemi)

    sekizOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(yediOncekiBilancoDonemi)
    logging.debug("Sekiz Onceki Bilanco Donemi: %d", sekizOncekiBilancoDonemi)

    wb = xlrd.open_workbook(bilancoDosyasi)
    sheet = wb.sheet_by_index(0)

    def donemColumnFind(col):
        for columni in range(sheet.ncols):
            cell = sheet.cell(0, columni)
            if cell.value == col:
                return columni
        print("Uygun Ceyrek Bulunamadi!!!")
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
                    print (label + " :Bilanço alanı boş!")
                    return 0
                else:
                    return sheet.cell_value(rowi, column)
        print("Uygun bilanco degeri bulunamadi:", label)
        return 0


    def getBilancoTitleRow(title):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == title:
                return rowi
        print("Uygun baslik bulunamadi!")
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
                print ("EKSİK BİLANÇO VAR!")
                return -1


    def oncekiYilAyniCeyrekDegisimiHesapla(row, donem):
        logging.debug("fonksiyon: oncekiYilAyniCeyrekDegisimiHesapla")
        donemColumn = donemColumnFind(donem)
        logging.debug ("DonemColumn: %s", donemColumn)
        oncekiYilAyniDonemColumn = donemColumnFind(donem - 100)
        logging.debug("Onceki Yıl Aynı DonemColumn: %s", oncekiYilAyniDonemColumn)
        logging.debug("Row: %d Column: %d",row ,donemColumn)
        ceyrekDegeri = ceyrekDegeriHesapla(row, donemColumn)
        logging.debug("Çeyrek Değeri: %d", ceyrekDegeri)
        oncekiCeyrekDegeri = ceyrekDegeriHesapla(row, oncekiYilAyniDonemColumn)
        logging.debug ("Önceki Çeyrek Değeri: %d", oncekiCeyrekDegeri)
        degisimSonucu = ceyrekDegeri / oncekiCeyrekDegeri - 1
        logging.debug("%d %s %d", sheet.cell_value(0, donemColumn), sheet.cell_value(row, 0), ceyrekDegeri)
        logging.debug("%d %s %d" ,sheet.cell_value(0, oncekiYilAyniDonemColumn), sheet.cell_value(row, 0), oncekiCeyrekDegeri)
        #print(int(sheet.cell_value(0, donemColumn)), sheet.cell_value(row, 0), "{:,.0f}".format(ceyrekDegeri).replace(",","."), "TL")
        #print(int(sheet.cell_value(0, oncekiYilAyniDonemColumn)), sheet.cell_value(row, 0), "{:,.0f}".format(oncekiCeyrekDegeri).replace(",","."), "TL")
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
    print("")
    print("--------------------HASILAT(SATIŞ) GELİRLERİ---------------------------")
    print("")

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
    print(satisTablosu)

    # Bilanço Dönemi Saış Geliri Artış Kriteri
    bilancoDonemiHasilatGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    bilancoDonemiHasilatGelirArtisiGecmeDurumu = (bilancoDonemiHasilatGelirArtisi > 0.1)
    print("Bilanço Dönemi Satış Geliri Artışı %10'dan Büyük Mü:", "{:.2%}".format(bilancoDonemiHasilatGelirArtisi), ">? 10%", bilancoDonemiHasilatGelirArtisiGecmeDurumu)

    # Önceki Dönem Hasılat Geliri Artış Kriteri
    oncekiDonemHasilatGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    if (bilancoDonemiHasilatGelirArtisi >= 1):
        print ("Bilanço Dönemi Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak.")
        oncekiDonemHasilatGelirArtisiGecmeDurumu = True
        print ("Önceki Dönem Satış Gelir Artışı Geçme Durumu:", oncekiDonemHasilatGelirArtisiGecmeDurumu)

    else:
        oncekiDonemHasilatGelirArtisiGecmeDurumu = (oncekiDonemHasilatGelirArtisi<bilancoDonemiHasilatGelirArtisi)
        print("Önceki Dönem Satış Gelir Artışı Bilanço Döneminden Düşük Mü:", "{:.2%}".format(oncekiDonemHasilatGelirArtisi),"<?","{:.2%}".format(bilancoDonemiHasilatGelirArtisi), oncekiDonemHasilatGelirArtisiGecmeDurumu)


    # Faaliyet Karı Gelirleri
    print("")
    print("--------------------------FAALİYET KARI---------------------------------")
    print("")

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
    print(faaliyetKariTablosu)


    # Bilanço Dönemi Faaliyet Kar Artış Kriteri
    if ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn) < 0:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = False
        print("Bilanço Dönemi Faaliyet Kari Artisi:", bilancoDonemiFaaliyetKariArtisiGecmeDurumu, "Son Çeyrek Net Kar Negatif")

    elif ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) < 0:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = False
        print("Bilanço Dönemi Faaliyet Kari Artisi:", bilancoDonemiFaaliyetKariArtisiGecmeDurumu, "Son Ceyrek Faaliyet Kari Negatif")

    elif ((ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) > 0) and (ceyrekDegeriHesapla(faaliyetKariRow, dortOncekibilancoDonemiColumn)) < 0):
        bilancoDonemiFaaliyetKariArtisi = 0
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = True
        print("Bilanço Dönemi Faaliyet Kari Artisi:", bilancoDonemiFaaliyetKariArtisiGecmeDurumu, "Son Çeyrek Faaliyet Karı Negatiften Pozitife Geçmiş")

    else:
        bilancoDonemiFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        bilancoDonemiFaaliyetKariArtisiGecmeDurumu = (bilancoDonemiFaaliyetKariArtisi > 0.15)
        print("Bilanço Dönemi Faaliyet Kari Artisi:", "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi), ">? 15%", bilancoDonemiFaaliyetKariArtisiGecmeDurumu)

    # Önceki Dönem Faaliyet Kar Artış Kriteri

    if bilancoDonemiFaaliyetKariArtisi >= 1:
        oncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiCeyrekFaaliyetKarArtisiGecmeDurumu = True
        print("Önceki Dönem Faaliyet Kar Artışı: "
              "Bilanço Dönemi Faaliyet Karı Artışı %100'ün Üzerinde, Karşılaştırma Yapılmayacak:", "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi), oncekiCeyrekFaaliyetKarArtisiGecmeDurumu)

    else:
        oncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiCeyrekFaaliyetKarArtisiGecmeDurumu = (oncekiCeyrekFaaliyetKariArtisi < bilancoDonemiFaaliyetKariArtisi)
        print("Önceki Dönem Faaliyet Kar Artışı:", "{:.2%}".format(oncekiCeyrekFaaliyetKariArtisi),
          "<?" , "{:.2%}".format(bilancoDonemiFaaliyetKariArtisi) , oncekiCeyrekFaaliyetKarArtisiGecmeDurumu)



    # Net Kar Hesabı
    print("")
    print("-------------------NET KAR (DÖNEM KARI/ZARARI)--------------------------")
    print("")

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
    print(netKarTablosu)





    # Brüt Kar Hesabı
    print("")
    print("-------------------BRÜT KAR (BRÜT KAR/ZARAR)--------------------------")
    print("")

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
    print(brutKarTablosu)





    # Gerçek Deger Hesaplama
    print("")
    print("")
    print("----------------GERÇEK DEĞER HESABI--------------------------------------------")

    sermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)
    print("Sermaye:", "{:,.0f}".format(sermaye).replace(",","."), "TL")

    anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", bilancoDonemiColumn) / getBilancoDegeri(
        "DÖNEM KARI (ZARARI)", bilancoDonemiColumn)
    print("Ana Ortaklık Payı:", "{:.3f}".format(anaOrtaklikPayi))

    sonCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    birOncekiCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    sonDortCeyrekHasilatToplami = ucOncekiBilancoDonemiHasilat + ikiOncekiBilancoDonemiHasilat + birOncekiBilancoDonemiHasilat + bilancoDonemiHasilat

    print("Son 4 Çeyrek Hasılat Toplamı:", "{:,.0f}".format(sonDortCeyrekHasilatToplami).replace(",","."), "TL")

    onumuzdekiDortCeyrekHasilatTahmini = (
                (((sonCeyrekSatisArtisYuzdesi + birOncekiCeyrekSatisArtisYuzdesi) / 2) + 1) * sonDortCeyrekHasilatToplami)
    print("Önümüzdeki 4 Çeyrek Hasılat Tahmini:", "{:,.0f}".format(onumuzdekiDortCeyrekHasilatTahmini).replace(",","."), "TL")

    # HASILAT TAHMININI MANUEL DEGISTIRMEK ICIN
    #onumuzdekiDortCeyrekHasilatTahmini = 5000000000

    ucOncekibilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, birOncekibilancoDonemiColumn)
    bilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)

    onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) / (
                bilancoDonemiHasilat + birOncekiBilancoDonemiHasilat)
    print("Önümüzdeki 4 çeyrek faaliyet kar marjı tahmini:",
          "{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))

    faaliyetKariTahmini1 = onumuzdekiDortCeyrekHasilatTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
    print("Faaliyet Kar Tahmini1:", "{:,.0f}".format(faaliyetKariTahmini1).replace(",","."), "TL")

    faaliyetKariTahmini2 = ((birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 2 * 0.3) + (
                bilancoDonemiFaaliyetKari * 4 * 0.5) + \
                           ((
                                        ucOncekibilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 0.2)
    print("Faaliyet Kar Tahmini2:", "{:,.0f}".format(faaliyetKariTahmini2).replace(",","."), "TL")

    ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1 + faaliyetKariTahmini2) / 2
    print("Ortalama Faaliyet Kari Tahmini:", "{:,.0f}".format(ortalamaFaaliyetKariTahmini).replace(",","."), "TL")

    hisseBasinaOrtalamaKarTahmini = (ortalamaFaaliyetKariTahmini * anaOrtaklikPayi) / sermaye
    print("Hisse başına ortalama kar tahmini:", format(hisseBasinaOrtalamaKarTahmini, ".2f") ,"TL")

    likidasyonDegeri = likidasyonDegeriHesapla(bilancoDonemi)
    print("Likidasyon değeri:", "{:,.0f}".format(likidasyonDegeri).replace(",","."), "TL")

    borclar = int(getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", bilancoDonemiColumn))
    print("Borçlar:", "{:,.0f}".format(borclar).replace(",","."), "TL")

    bilancoEtkisi = (likidasyonDegeri - borclar) / sermaye * anaOrtaklikPayi
    print("Bilanço Etkisi:", format(bilancoEtkisi, ".2f"), "TL")

    gercekDeger = (hisseBasinaOrtalamaKarTahmini * 7) + bilancoEtkisi
    print("Gerçek hisse değeri:", format(gercekDeger, ".2f"), "TL")

    targetBuy = gercekDeger * 0.66
    print("Target buy:", format(targetBuy, ".2f"), "TL")

    print("Bilanço tarihindeki hisse fiyatı:", format(hisseFiyati, ".2f"), "TL")

    gercekFiyataUzaklik = hisseFiyati / targetBuy
    print("Gerçek fiyata uzaklık oranı:", "{:.2%}".format(gercekFiyataUzaklik))

    gercekFiyataUzaklikTl = hisseFiyati - targetBuy
    print("Gerçek fiyata uzaklık TL:", format(gercekFiyataUzaklikTl, ".2f"))

    # Netpro Hesapla
    print("")
    print("")
    print("")
    print("----------------NetPro Kriteri-----------------------------------------------------------------")

    sonDortDonemFaaliyetKariToplami = bilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + ucOncekibilancoDonemiFaaliyetKari

    ucOncekibilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    bilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn)

    sonDortDonemNetKarToplami = bilancoDonemiNetKari + birOncekiBilancoDonemiNetKari + ikiOncekiBilancoDonemiNetKari + ucOncekibilancoDonemiNetKari

    print ("Son 4 Dönem Net Kar Toplamı:", "{:,.0f}".format(sonDortDonemNetKarToplami).replace(",", "."), "TL")
    print ("Son 4 Dönem Faaliyet Karı Toplamı:", "{:,.0f}".format(sonDortDonemFaaliyetKariToplami).replace(",", "."), "TL")


    fkOrani = hisseFiyati/((sonDortDonemNetKarToplami*anaOrtaklikPayi)/(sermaye))
    print("F/K Oranı:", "{:,.2f}".format(fkOrani))

    hbkOrani = sonDortDonemNetKarToplami/(sermaye)
    print ("HBK Oranı:", "{:,.2f}".format(hbkOrani))

    netProEstDegeri = ((ortalamaFaaliyetKariTahmini / sonDortDonemFaaliyetKariToplami) * sonDortDonemNetKarToplami) * anaOrtaklikPayi
    print("NetPro Est Değeri:", "{:,.0f}".format(netProEstDegeri).replace(",","."), "TL")

    piyasaDegeri = (bilancoEtkisi * sermaye * -1) + (hisseFiyati * sermaye)

    print("Piyasa Değeri:", "{:,.0f}".format(piyasaDegeri).replace(",","."), "TL")
    print("BondYield:", "{:.2%}".format(bondYield))

    netProKriteri = (netProEstDegeri / piyasaDegeri) / bondYield
    netProKriteriGecmeDurumu = (netProKriteri > 2)
    print("NetPro Kriteri (2'den Büyük Olmalı):", format(netProKriteri, ".2f"), netProKriteriGecmeDurumu)

    minNetProIcinHisseFiyati  = (netProEstDegeri / (1.9 * bondYield) - (bilancoEtkisi * sermaye * -1))/sermaye
    print("NetPro 1.9 Olmasi Icin Hisse Fiyatı:", format(minNetProIcinHisseFiyati, ".2f"), )



    # Forward PE Hesapla
    print("")
    print("")
    print("----------------FORWARD PE HESAPLAMA--------------------------------------------------------")

    forwardPeKriteri = (piyasaDegeri) / netProEstDegeri

    forwardPeKriteriGecmeDurumu = (forwardPeKriteri < 4)
    print("Forward PE Kriteri (4'ten Küçük Olmalı):", format(forwardPeKriteri, ".2f"), forwardPeKriteriGecmeDurumu)





    # Ek Hesaplama ve Tablolar
    print("")
    print("-------------------EK HESAPLAMA ve TABLOLAR--------------------------")
    print("")

    bilancoDonemiBrutKarMarji = bilancoDonemiBrutKar/bilancoDonemiHasilat;
    print("Bilanço Dönemi Brüt Kar Marjı:", bilancoDonemiBrutKarMarji)

    bilancoDonemiFaaliyetKarMarji = bilancoDonemiFaaliyetKari/bilancoDonemiHasilat;
    print("Bilanço Dönemi Faaliyet Kar Marjı:", bilancoDonemiFaaliyetKarMarji)

    bilancoDonemiNetKarMarji = bilancoDonemiNetKari/bilancoDonemiHasilat;
    print("Bilanço Dönemi Net Kar Marjı:", bilancoDonemiNetKarMarji)

    bilancoDonemiOzsermayeKarliligi = bilancoDonemiNetKari/getBilancoDegeri("TOPLAM ÖZKAYNAKLAR", bilancoDonemiColumn)
    print("Bilanço Dönemi Özsermaye Karlılığı:", bilancoDonemiOzsermayeKarliligi)




    print("")
    print("")
    print("----------------BİLANÇO DOLAR HESABI-------------------------------------")
    print("")
    bilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(bilancoDonemi)

    #print (bilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ", "{:,.3f}".format(bilancoDonemiOrtalamaDolarKuru))

    birOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(birOncekiBilancoDonemi)
    #print(birOncekiBilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ", "{:,.3f}".format(birOncekiBilancoDonemiOrtalamaDolarKuru))

    ikiOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(ikiOncekiBilancoDonemi)
    #print(ikiOncekiBilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ","{:,.3f}".format(ikiOncekiBilancoDonemiOrtalamaDolarKuru))

    ucOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(ucOncekiBilancoDonemi)
    #print(ucOncekiBilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ","{:,.3f}".format(ucOncekiBilancoDonemiOrtalamaDolarKuru))

    dortOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(dortOncekiBilancoDonemi)
    #print(dortOncekiBilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ","{:,.3f}".format(dortOncekiBilancoDonemiOrtalamaDolarKuru))

    besOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(besOncekiBilancoDonemi)
    #print(besOncekiBilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ","{:,.3f}".format(besOncekiBilancoDonemiOrtalamaDolarKuru))

    altiOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(altiOncekiBilancoDonemi)
    #print(altiOncekiBilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ","{:,.3f}".format(altiOncekiBilancoDonemiOrtalamaDolarKuru))

    yediOncekiBilancoDonemiOrtalamaDolarKuru = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(yediOncekiBilancoDonemi)
    #print(yediOncekiBilancoDonemi, "Bilanço Dönemi Ortalama Dolar Kuru: ","{:,.3f}".format(yediOncekiBilancoDonemiOrtalamaDolarKuru))





    # Bilanço Dönemi Satış(Hasılat) Gelirleri (DOLAR)
    print("")
    print("--------------------HASILAT(SATIŞ) GELİRLERİ (DOLAR)----------------------")
    print("")

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
    #
    dolarSatisTablosu = PrettyTable()
    dolarSatisTablosu.field_names = ["ÇEYREK", "SATIŞ (USD)", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ (USD)", "YÜZDE DEĞİŞİM"]
    dolarSatisTablosu.align["SATIŞ (USD)"] = "r"
    dolarSatisTablosu.align["ÖNCEKİ YIL SATIŞ (USD)"] = "r"
    dolarSatisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
    dolarSatisTablosu.add_row([bilancoDonemi, bilancoDonemiDolarHasilatPrint, dortOncekiBilancoDonemi, dortOncekiBilancoDonemiDolarHasilatPrint, bilancoDonemiDolarHasilatDegisimiPrint])
    dolarSatisTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiDolarHasilatPrint, besOncekiBilancoDonemi, besOncekiBilancoDonemiDolarHasilatPrint, birOncekiBilancoDonemiDolarHasilatDegisimiPrint])
    dolarSatisTablosu.add_row([ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiDolarHasilatPrint, altiOncekiBilancoDonemi, altiOncekiBilancoDonemiDolarHasilatPrint, ikiOncekiBilancoDonemiDolarHasilatDegisimiPrint])
    dolarSatisTablosu.add_row([ucOncekiBilancoDonemi, ucOncekiBilancoDonemiDolarHasilatPrint, yediOncekiBilancoDonemi,yediOncekiBilancoDonemiDolarHasilatPrint, ucOncekiBilancoDonemiDolarHasilatDegisimiPrint])
    print(dolarSatisTablosu)

    # Bilanço Dönemi (DOLAR) Satış Geliri Artış Kriteri
    bilancoDonemiDolarHasilatGelirArtisi = bilancoDonemiDolarHasilat/oncekiYilAyniCeyrekDolarHasilat-1
    bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (bilancoDonemiDolarHasilatGelirArtisi > 0.1)
    print("Bilanço Dönemi (DOLAR) Satış Geliri Artışı %10'dan Büyük Mü:", "{:.2%}".format(bilancoDonemiDolarHasilatGelirArtisi), ">? 10%", bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)

    # Önceki Dönem (DOLAR) Hasılat Geliri Artış Kriteri
    oncekiDonemDolarHasilatGelirArtisi = birOncekiBilancoDonemiDolarHasilat/besOncekiBilancoDonemiDolarHasilat-1
    #
    if (bilancoDonemiDolarHasilatGelirArtisi >= 1):
        print ("Bilanço Dönemi (DOLAR) Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak.")
        oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = True
        print ("Önceki Dönem (DOLAR) Satış Gelir Artışı Geçme Durumu:", oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)

    else:
        oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (oncekiDonemDolarHasilatGelirArtisi<bilancoDonemiDolarHasilatGelirArtisi)
        print("Önceki Dönem Satış Gelir Artışı Bilançp Döneminden Düşük Mü:", "{:.2%}".format(oncekiDonemDolarHasilatGelirArtisi),"<?","{:.2%}".format(bilancoDonemiDolarHasilatGelirArtisi), oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)




        # Faaliyet Karı Gelirleri (DOLAR)
        print("")
        print("--------------------------FAALİYET KARI (DOLAR)-------------------------")
        print("")

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
        print(dolarFaaliyetKariTablosu)

        # Bilanço Dönem Faaliyet Kar Artış Kriteri (DOLAR)
        if ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn) < 0:
            bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
            print("Bilanço Dönemi Dolar Faaliyet Kari Artisi:", bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu, "Son Çeyrek Net Kar Negatif")

        elif (bilancoDonemiDolarFaaliyetKari < 0):
            bilancoDonemiDolarFaaliyetKariArtisi = bilancoDonemiDolarFaaliyetKari/dortOncekiBilancoDonemiDolarFaaliyetKari -1
            bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
            print("Bilanço Dönemi Dolar Faaliyet Kari Artisi:", bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu, "Son Ceyrek Dolar Faaliyet Kari Negatif")

        elif (bilancoDonemiDolarFaaliyetKari > 0) and (dortOncekiBilancoDonemiDolarFaaliyetKari < 0):
            bilancoDonemiDolarFaaliyetKariArtisi = 0
            bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = True
            print("Bilanço Dönemi Dolar Faaliyet Kari Artisi:", bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu, "Son Çeyrek Dolar Faaliyet Karı Negatiften Pozitife Geçmiş")

        else:
            bilancoDonemiDolarFaaliyetKariArtisi = bilancoDonemiDolarFaaliyetKari/dortOncekiBilancoDonemiDolarFaaliyetKari -1
            bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = (bilancoDonemiDolarFaaliyetKariArtisi > 0.15)
            print("Bilanço Dönemi Dolar Faaliyet Kari Artisi:", "{:.2%}".format(bilancoDonemiDolarFaaliyetKariArtisi), ">? 15%",bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu)

        # Önceki Dönem Faaliyet Kar Artış Kriteri (DOLAR)

        if bilancoDonemiDolarFaaliyetKariArtisi >= 1:
            birOncekiBilancoDonemiDolarFaaliyetKariArtisi = oncekiCeyrekFaaliyetKariArtisi/birOncekiBilancoDonemiOrtalamaDolarKuru
            birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = True
            print("Önceki Bilanço Dönemi Dolar Faaliyet Kar Artışı: "
                  "Bilanço Dönemi Dolar Faaliyet Karı Artışı %100'ün Üzerinde, Karşılaştırma Yapılmayacak:",
                  "{:.2%}".format(bilancoDonemiDolarFaaliyetKariArtisi), birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)

        else:
            birOncekiBilancoDonemiDolarFaaliyetKariArtisi = oncekiCeyrekFaaliyetKariArtisi/birOncekiBilancoDonemiOrtalamaDolarKuru
            birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = (birOncekiBilancoDonemiDolarFaaliyetKariArtisi < bilancoDonemiDolarFaaliyetKariArtisi)
            print("Önceki Bilanço Dönemi Dolar Faaliyet Kar Artışı:", "{:.2%}".format(birOncekiBilancoDonemiDolarFaaliyetKariArtisi),
                  "<?", "{:.2%}".format(bilancoDonemiDolarFaaliyetKariArtisi), birOncekiBilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)



    print("")
    print("")
    print("----------------RAPOR DOSYASI OLUŞTURMA/GÜNCELLEME-------------------------------------")

    print (hisseAdi)

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
    excelRow.netProKriteri = netProKriteri
    excelRow.forwardPeKriteri = forwardPeKriteri

    exportReportExcel(hisseAdi, reportFile, bilancoDonemi, excelRow)