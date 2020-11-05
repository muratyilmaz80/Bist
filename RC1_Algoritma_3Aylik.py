import xlrd
from ExcelRowClass import ExcelRowClass
from Rapor_Olustur import exportReportExcel
from prettytable import PrettyTable
import logging

#logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.INFO, format='%(levelname)s - %(message)s')

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

    # Cari Dönem Satış(Hasılat) Gelirleri
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
    bilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi))

    birOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(birOncekiBilancoDonemiHasilat).replace(",", ".")
    besOncekiBilancoDonemiHasilatPrint = "{:,.0f}".format(besOncekiBilancoDonemiHasilat).replace(",", ".")
    birOncekiBilancoDonemiHasilatDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi))

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
    satisTablosu.add_row([bilancoDonemi, bilancoDonemiHasilatPrint, dortOncekiBilancoDonemi, dortOncekiBilancoDonemiHasilatPrint, bilancoDonemiHasilatDegisimiPrint])
    satisTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiHasilatPrint, besOncekiBilancoDonemi, besOncekiBilancoDonemiHasilatPrint, birOncekiBilancoDonemiHasilatDegisimiPrint])
    satisTablosu.add_row([ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiHasilatPrint, altiOncekiBilancoDonemi, altiOncekiBilancoDonemiHasilatPrint, ikiOncekiBilancoDonemiHasilatDegisimiPrint])
    satisTablosu.add_row([ucOncekiBilancoDonemi, ucOncekiBilancoDonemiHasilatPrint, yediOncekiBilancoDonemi,yediOncekiBilancoDonemiHasilatPrint, ucOncekiBilancoDonemiHasilatDegisimiPrint])
    print(satisTablosu)

    # Cari Dönem Saış Geliri Artış Kriteri
    cariDonemHasilatGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    cariDonemHasilatGelirArtisiGecmeDurumu = (cariDonemHasilatGelirArtisi > 0.1)
    print("Cari Dönem Satış Geliri Artışı %10'dan Büyük Mü:", "{:.2%}".format(cariDonemHasilatGelirArtisi), ">? 10%", cariDonemHasilatGelirArtisiGecmeDurumu)

    # Önceki Dönem Hasılat Geliri Artış Kriteri
    oncekiDonemHasilatGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    if (cariDonemHasilatGelirArtisi >= 1):
        print ("Cari Dönem Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak.")
        oncekiDonemHasilatGelirArtisiGecmeDurumu = True
        print ("Önceki Dönem Satış Gelir Artışı Geçme Durumu:", oncekiDonemHasilatGelirArtisiGecmeDurumu)

    else:
        oncekiDonemHasilatGelirArtisiGecmeDurumu = (oncekiDonemHasilatGelirArtisi<cariDonemHasilatGelirArtisi)
        print("Önceki Dönem Satış Gelir Artışı Cari Dönemden Düşük Mü:", "{:.2%}".format(oncekiDonemHasilatGelirArtisi),"<?","{:.2%}".format(cariDonemHasilatGelirArtisi), oncekiDonemHasilatGelirArtisiGecmeDurumu)


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
    bilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi))

    birOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(birOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    besOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(besOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    birOncekiBilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi))

    ikiOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    altiOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(altiOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    ikiOncekiBilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, ikiOncekiBilancoDonemi))

    ucOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(ucOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    yediOncekiBilancoDonemiFaaliyetKariPrint = "{:,.0f}".format(yediOncekiBilancoDonemiFaaliyetKari).replace(",", ".")
    ucOncekiBilancoDonemiFaaliyetKariDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, ucOncekiBilancoDonemi))

    faaliyetKariTablosu = PrettyTable()
    faaliyetKariTablosu.field_names = ["ÇEYREK", "FAALİYET KARI", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI", "YÜZDE DEĞİŞİM"]
    faaliyetKariTablosu.align["FAALİYET KARI"] = "r"
    faaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI"] = "r"
    faaliyetKariTablosu.add_row([bilancoDonemi, bilancoDonemiFaaliyetKariPrint, dortOncekiBilancoDonemi, dortOncekiBilancoDonemiFaaliyetKariPrint, bilancoDonemiFaaliyetKariDegisimiPrint])
    faaliyetKariTablosu.add_row([birOncekiBilancoDonemi, birOncekiBilancoDonemiFaaliyetKariPrint, besOncekiBilancoDonemi, besOncekiBilancoDonemiFaaliyetKariPrint, birOncekiBilancoDonemiFaaliyetKariDegisimiPrint])
    faaliyetKariTablosu.add_row([ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiFaaliyetKariPrint, altiOncekiBilancoDonemi, altiOncekiBilancoDonemiFaaliyetKariPrint, ikiOncekiBilancoDonemiFaaliyetKariDegisimiPrint])
    faaliyetKariTablosu.add_row([ucOncekiBilancoDonemi, ucOncekiBilancoDonemiFaaliyetKariPrint, yediOncekiBilancoDonemi,yediOncekiBilancoDonemiFaaliyetKariPrint, ucOncekiBilancoDonemiFaaliyetKariDegisimiPrint])
    print(faaliyetKariTablosu)


    # Cari Dönem Faaliyet Kar Artış Kriteri
    if ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn) < 0:
        cariDonemFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        cariDonemFaaliyetKariArtisiGecmeDurumu = False
        print("Cari Dönem Faaliyet Kari Artisi:", cariDonemFaaliyetKariArtisiGecmeDurumu, "Son Çeyrek Net Kar Negatif")

    elif ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) < 0:
        cariDonemFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        cariDonemFaaliyetKariArtisiGecmeDurumu = False
        print("Cari Dönem Faaliyet Kari Artisi:", cariDonemFaaliyetKariArtisiGecmeDurumu, "Son Ceyrek Faaliyet Kari Negatif")

    elif ((ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) > 0) and (ceyrekDegeriHesapla(faaliyetKariRow, dortOncekibilancoDonemiColumn)) < 0):
        cariDonemFaaliyetKariArtisi = 0
        cariDonemFaaliyetKariArtisiGecmeDurumu = True
        print("Cari Dönem Faaliyet Kari Artisi:", cariDonemFaaliyetKariArtisiGecmeDurumu, "Son Çeyrek Faaliyet Karı Negatiften Pozitife Geçmiş")

    else:
        cariDonemFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        cariDonemFaaliyetKariArtisiGecmeDurumu = (cariDonemFaaliyetKariArtisi > 0.15)
        print("Cari Dönem Faaliyet Kari Artisi:", "{:.2%}".format(cariDonemFaaliyetKariArtisi), ">? 15%", cariDonemFaaliyetKariArtisiGecmeDurumu)

    # Önceki Dönem Faaliyet Kar Artış Kriteri

    if cariDonemFaaliyetKariArtisi >= 1:
        oncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiCeyrekFaaliyetKarArtisiGecmeDurumu = True
        print("Önceki Dönem Faaliyet Kar Artışı: "
              "Cari Dönem Faaliyet Karı Artışı %100'ün Üzerinde, Karşılaştırma Yapılmayacak:", "{:.2%}".format(cariDonemFaaliyetKariArtisi), oncekiCeyrekFaaliyetKarArtisiGecmeDurumu)

    else:
        oncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        oncekiCeyrekFaaliyetKarArtisiGecmeDurumu = (oncekiCeyrekFaaliyetKariArtisi < cariDonemFaaliyetKariArtisi)
        print("Önceki Dönem Faaliyet Kar Artışı:", "{:.2%}".format(oncekiCeyrekFaaliyetKariArtisi),
          "<?" , "{:.2%}".format(cariDonemFaaliyetKariArtisi) , oncekiCeyrekFaaliyetKarArtisiGecmeDurumu)



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
    birOncekiBilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(
        oncekiYilAyniCeyrekDegisimiHesapla(netKarRow, birOncekiBilancoDonemi))

    ikiOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiNetKar).replace(",", ".")
    altiOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(altiOncekiBilancoDonemiNetKar).replace(",", ".")
    ikiOncekiBilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(
        oncekiYilAyniCeyrekDegisimiHesapla(netKarRow, ikiOncekiBilancoDonemi))

    ucOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(ucOncekiBilancoDonemiNetKar).replace(",", ".")
    yediOncekiBilancoDonemiNetKarPrint = "{:,.0f}".format(yediOncekiBilancoDonemiNetKar).replace(",",".")
    ucOncekiBilancoDonemiNetKarDegisimiPrint = "{:.2%}".format(
        oncekiYilAyniCeyrekDegisimiHesapla(netKarRow, ucOncekiBilancoDonemi))

    netKarTablosu = PrettyTable()
    netKarTablosu.field_names = ["ÇEYREK", "NET KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL NET KAR",
                                       "YÜZDE DEĞİŞİM"]
    netKarTablosu.align["NET KAR"] = "r"
    netKarTablosu.align["ÖNCEKİ YIL NET KAR"] = "r"
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
    birOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(
        oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, birOncekiBilancoDonemi))

    ikiOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiBrutKar).replace(",", ".")
    altiOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(altiOncekiBilancoDonemiBrutKar).replace(",", ".")
    ikiOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(
        oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, ikiOncekiBilancoDonemi))

    ucOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ucOncekiBilancoDonemiBrutKar).replace(",", ".")
    yediOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(yediOncekiBilancoDonemiBrutKar).replace(",", ".")
    ucOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(
        oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, ucOncekiBilancoDonemi))

    brutKarTablosu = PrettyTable()
    brutKarTablosu.field_names = ["ÇEYREK", "BRÜT KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL BRÜT KAR", "YÜZDE DEĞİŞİM"]
    brutKarTablosu.align["BRÜT KAR"] = "r"
    brutKarTablosu.align["ÖNCEKİ YIL BRÜT KAR"] = "r"
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

    #ucOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, ucOncekibilancoDonemiColumn)
    #ikiOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, ikiOncekibilancoDonemiColumn)
    #birOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, birOncekibilancoDonemiColumn)
    #bilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, bilancoDonemiColumn)

    #sonDortCeyrekSatisToplami = ucOncekiBilancoDonemiSatis + ikiOncekiBilancoDonemiSatis + birOncekiBilancoDonemiSatis + bilancoDonemiSatis

    sonDortCeyrekHasilatToplami = ucOncekiBilancoDonemiHasilat + ikiOncekiBilancoDonemiHasilat + birOncekiBilancoDonemiHasilat + bilancoDonemiHasilat

    print("Son 4 Çeyrek Hasılat Toplamı:", "{:,.0f}".format(sonDortCeyrekHasilatToplami).replace(",","."), "TL")

    onumuzdekiDortCeyrekHasilatTahmini = (
                (((sonCeyrekSatisArtisYuzdesi + birOncekiCeyrekSatisArtisYuzdesi) / 2) + 1) * sonDortCeyrekHasilatToplami)
    print("Önümüzdeki 4 Çeyrek Hasılat Tahmini:", "{:,.0f}".format(onumuzdekiDortCeyrekHasilatTahmini).replace(",","."), "TL")

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


    # bilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, bilancoDonemiColumn)
    # birOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, birOncekibilancoDonemiColumn)
    # ikiOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, ikiOncekibilancoDonemiColumn)
    # ucOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, ucOncekibilancoDonemiColumn)
    # dortOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, dortOncekibilancoDonemiColumn)
    # besOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, besOncekibilancoDonemiColumn)
    # altiOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, altiOncekibilancoDonemiColumn)
    # yediOncekiBilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, yediOncekibilancoDonemiColumn)
    #
    # bilancoDonemiBrutKarPrint = "{:,.0f}".format(bilancoDonemiBrutKar).replace(",", ".")
    # dortOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(dortOncekiBilancoDonemiBrutKar).replace(",", ".")
    # bilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, bilancoDonemi))
    #
    # birOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(birOncekiBilancoDonemiBrutKar).replace(",", ".")
    # besOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(besOncekiBilancoDonemiBrutKar).replace(",", ".")
    # birOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(
    #     oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, birOncekiBilancoDonemi))
    #
    # ikiOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ikiOncekiBilancoDonemiBrutKar).replace(",", ".")
    # altiOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(altiOncekiBilancoDonemiBrutKar).replace(",", ".")
    # ikiOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(
    #     oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, ikiOncekiBilancoDonemi))
    #
    # ucOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(ucOncekiBilancoDonemiBrutKar).replace(",", ".")
    # yediOncekiBilancoDonemiBrutKarPrint = "{:,.0f}".format(yediOncekiBilancoDonemiBrutKar).replace(",", ".")
    # ucOncekiBilancoDonemiBrutKarDegisimiPrint = "{:.2%}".format(
    #     oncekiYilAyniCeyrekDegisimiHesapla(brutKarRow, ucOncekiBilancoDonemi))
    #
    # brutKarTablosu = PrettyTable()
    # brutKarTablosu.field_names = ["ÇEYREK", "BRÜT KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL BRÜT KAR", "YÜZDE DEĞİŞİM"]
    # brutKarTablosu.align["BRÜT KAR"] = "r"
    # brutKarTablosu.align["ÖNCEKİ YIL BRÜT KAR"] = "r"
    # brutKarTablosu.add_row([bilancoDonemi, bilancoDonemiBrutKarPrint, dortOncekiBilancoDonemi,
    #                         dortOncekiBilancoDonemiBrutKarPrint, bilancoDonemiBrutKarDegisimiPrint])
    # brutKarTablosu.add_row(
    #     [birOncekiBilancoDonemi, birOncekiBilancoDonemiBrutKarPrint, besOncekiBilancoDonemi,
    #      besOncekiBilancoDonemiBrutKarPrint, birOncekiBilancoDonemiBrutKarDegisimiPrint])
    # brutKarTablosu.add_row(
    #     [ikiOncekiBilancoDonemi, ikiOncekiBilancoDonemiBrutKarPrint, altiOncekiBilancoDonemi,
    #      altiOncekiBilancoDonemiBrutKarPrint, ikiOncekiBilancoDonemiBrutKarDegisimiPrint])
    # brutKarTablosu.add_row(
    #     [ucOncekiBilancoDonemi, ucOncekiBilancoDonemiBrutKarPrint, yediOncekiBilancoDonemi,
    #      yediOncekiBilancoDonemiBrutKarPrint, ucOncekiBilancoDonemiBrutKarDegisimiPrint])
    # print(brutKarTablosu)
    #
    #







    print("")
    print("")
    print("----------------RAPOR DOSYASI OLUŞTURMA/GÜNCELLEME-------------------------------------")

    print (hisseAdi)

    excelRow = ExcelRowClass()

    excelRow.sonCeyrekHasilat = ceyrekDegeriHesapla(hasilatRow, bilancoDonemiColumn)
    excelRow.oncekiYilAyniCeyrekHasilat = ceyrekDegeriHesapla(hasilatRow, dortOncekibilancoDonemiColumn)
    excelRow.hasilatArtisi = cariDonemHasilatGelirArtisi
    excelRow.birOncekiCeyrekHasilatArtisi = oncekiDonemHasilatGelirArtisi
    excelRow.kriter1 = cariDonemHasilatGelirArtisiGecmeDurumu
    excelRow.kriter3 = oncekiDonemHasilatGelirArtisiGecmeDurumu
    excelRow.sonCeyrekFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)
    excelRow.oncekiYilAyniCeyrekFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, dortOncekibilancoDonemiColumn)
    excelRow.faaliyetKarArtisi = cariDonemFaaliyetKariArtisi
    excelRow.birOncekiCeyrekFaaliyetKarArtisi = oncekiCeyrekFaaliyetKariArtisi
    excelRow.kriter2 = cariDonemFaaliyetKariArtisiGecmeDurumu
    excelRow.kriter4 = oncekiCeyrekFaaliyetKarArtisiGecmeDurumu
    excelRow.sermaye = sermaye
    excelRow.anaOrtaklikPayi = anaOrtaklikPayi
    excelRow.son4CeyrekSatisToplami = sonDortCeyrekHasilatToplami
    excelRow.onumuzdeki4CeyrekSatisTahmini = onumuzdekiDortCeyrekHasilatTahmini
    excelRow.onumuzdeki4CeyrekFaaliyetKarMarjiTahmini = onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
    excelRow.faaliyetKarTahmini1 = faaliyetKariTahmini1
    excelRow.faaliyetKarTahmini2 = faaliyetKariTahmini2
    excelRow.ortalamaFaaliyetKarTahmini = ortalamaFaaliyetKariTahmini
    excelRow.hisseBasinaKarTahmini = hisseBasinaOrtalamaKarTahmini
    excelRow.bilancoEtkisi = bilancoEtkisi
    excelRow.bilancoTarihiHisseFiyati = hisseFiyati
    excelRow.gercekHisseDegeri = gercekDeger
    excelRow.targetBuy = targetBuy
    excelRow.gercekFiyataUzaklik = gercekFiyataUzaklik
    excelRow.netPro = netProKriteri
    excelRow.forwardPe = forwardPeKriteri

    exportReportExcel(hisseAdi, reportFile, bilancoDonemi, excelRow)
