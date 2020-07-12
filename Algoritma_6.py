import xlrd
import xlwt
from xlutils.copy import copy
import os.path
import os

varHisseAdi = "BFREN"
#varBilancoDosyasi = ("D:\\bist\\bilancolar\\" + varHisseAdi + ".xlsx")
varBilancoDonemi = 202003
varBondYield = 0.0907
varHisseFiyati = 274.90


reportFile = "D:\\bist\\Report_2020_03.xls"

def runAlgoritma(bilancoDosyasi, bilancoDonemi, bondYield, hisseFiyati):
    def birOncekiBilancoDoneminiHesapla(dnm):
        yil = int(dnm / 100)
        ceyrek = int(dnm % 100)

        if ceyrek == 3:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + (ceyrek - 3)

    birOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(bilancoDonemi)
    print("Bir Onceki Bilanco Donemi:", birOncekiBilancoDonemi)

    ikiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(birOncekiBilancoDonemi)
    print("Iki Onceki Bilanco Donemi:", ikiOncekiBilancoDonemi)

    ucOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ikiOncekiBilancoDonemi)
    print("Uc Onceki Bilanco Donemi:", ucOncekiBilancoDonemi)

    dortOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ucOncekiBilancoDonemi)
    print("Dort Onceki Bilanco Donemi:", dortOncekiBilancoDonemi)

    besOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(dortOncekiBilancoDonemi)
    print("Bes Onceki Bilanco Donemi:", besOncekiBilancoDonemi)

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


    def getBilancoDegeri(label, column):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == label:
                if sheet.cell_value(rowi, column)=="":
                    print ("Bilanço alanı boş!")
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
    netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı");

    def ceyrekDegeriHesapla(r, c):
        quarter = int(sheet.cell_value(0, c)) % (100)
        if (quarter == 3):
            return sheet.cell_value(r, c)
        else:
            return (sheet.cell_value(r, c) - sheet.cell_value(r, (c - 1)))

    def oncekiYilAyniCeyrekDegisimiHesapla(row, donem):
        donemColumn = donemColumnFind(donem)
        oncekiYilAyniDonemColumn = donemColumnFind(donem - 100)
        ceyrekDegeri = ceyrekDegeriHesapla(row, donemColumn)
        oncekiCeyrekDegeri = ceyrekDegeriHesapla(row, oncekiYilAyniDonemColumn)
        degisimSonucu = ceyrekDegeri / oncekiCeyrekDegeri - 1
        print(int(sheet.cell_value(0, donemColumn)), sheet.cell_value(row, 0), int(ceyrekDegeri))
        print(int(sheet.cell_value(0, oncekiYilAyniDonemColumn)), sheet.cell_value(row, 0), int(oncekiCeyrekDegeri))
        return degisimSonucu

    # def yilCeyrekAyir (a):
    #     yil = int (a/100)
    #     ceyrek = int (a % 100)
    #     return (yil, ceyrek)
    #
    # hesaplanacakYil, hesaplanacakCeyrek = yilCeyrekAyir(hesaplanacakDonem)
    # print ("Hesaplanacak Yıl:", hesaplanacakYil, "Hesaplanacak Çeyrek:", hesaplanacakCeyrek)
    #
    #
    # def yilCeyrekBirlestir (yil, ceyrek):
    #     return 100*yil + ceyrek

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

    # 1.kriter hesabi
    print("---------------------------------------------------------------------------------")
    print("1.Kriter: Son ceyrek satisi onceki yil ayni ceyrege göre en az %10 fazla olacak")

    kriter1SatisGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    kriter1GecmeDurumu = (kriter1SatisGelirArtisi > 0.1)
    print("Kriter1: Satis Geliri Artisi:", "{:.2%}".format(kriter1SatisGelirArtisi), kriter1GecmeDurumu)

    # 2.kriter hesabi
    print("---------------------------------------------------------------------------------")
    print("2.Kriter: Son ceyrek faaliyet kari onceki yil ayni ceyrege göre en az %15 fazla olacak")

    if getBilancoDegeri("Net Dönem Karı veya Zararı",bilancoDonemiColumn)<0:
        kriter2FaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        kriter2GecmeDurumu = False
        print("Kriter2: Faaliyet Kari Artisi:", kriter2GecmeDurumu, "Son Ceyrek Faaliyet Kari Negatif")

    else:
        kriter2FaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        kriter2GecmeDurumu = (kriter2FaaliyetKariArtisi > 0.15)
        print("Kriter2: Faaliyet Kari Artisi:", "{:.2%}".format(kriter2FaaliyetKariArtisi), kriter2GecmeDurumu)


    # 3.kriter hesabı
    print("---------------------------------------------------------------------------------")
    print("3.Kriter: Bir Onceki Ceyrek Satis Artis Yuzdesi Cari Donemden Dusuk Olmali")

    kriter3OncekiCeyrekArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)
    kriter3GecmeDurumu = (kriter3OncekiCeyrekArtisi < kriter1SatisGelirArtisi)
    print("Kriter3: Onceki Ceyrek Satis Geliri Artisi:", "{:.2%}".format(kriter3OncekiCeyrekArtisi), kriter3GecmeDurumu)

    # 4.kriter hesabi
    print("---------------------------------------------------------------------------------")
    print("4.Kriter: Bir Onceki Ceyrek Faaliyet Kar Artis Yuzdesi Cari Donemden Dusuk Olmali")
    kriter4OncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
    kriter4GecmeDurumu = (kriter4OncekiCeyrekFaaliyetKariArtisi < kriter2FaaliyetKariArtisi)
    print("Kriter4: Onceki Yila Gore Faaliyet Kari Artisi:", "{:.2%}".format(kriter4OncekiCeyrekFaaliyetKariArtisi),
          kriter4GecmeDurumu)

    # Gercek Deger Hesapla
    print("----------------Gercek Deger Hesabi-----------------------------------------------------------------")

    sermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)
    print("Sermaye:", sermaye)

    anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", bilancoDonemiColumn) / getBilancoDegeri(
        "DÖNEM KARI (ZARARI)", bilancoDonemiColumn)
    print("Ana Ortaklık Payı:", anaOrtaklikPayi)

    sonCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    birOncekiCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    ucOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, birOncekibilancoDonemiColumn)
    bilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, bilancoDonemiColumn)

    sonDortCeyrekSatisToplami = ucOncekiBilancoDonemiSatis + ikiOncekiBilancoDonemiSatis + birOncekiBilancoDonemiSatis + bilancoDonemiSatis
    print("Son 4 ceyrek satış toplamı:", sonDortCeyrekSatisToplami)

    onumuzdekiDortCeyrekSatisTahmini = (
                (((sonCeyrekSatisArtisYuzdesi + birOncekiCeyrekSatisArtisYuzdesi) / 2) + 1) * sonDortCeyrekSatisToplami)
    print("Önümüzdeki 4 çeyrek satış tahmini:", onumuzdekiDortCeyrekSatisTahmini)

    ucOncekibilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, birOncekibilancoDonemiColumn)
    bilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)

    onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) / (
                bilancoDonemiSatis + birOncekiBilancoDonemiSatis)
    print("Önümüzdeki 4 çeyrek faaliyet kar marjı tahmini:",
          "{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))

    faaliyetKariTahmini1 = onumuzdekiDortCeyrekSatisTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
    print("Faaliyet Kar Tahmini1:", faaliyetKariTahmini1)

    faaliyetKariTahmini2 = ((birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 2 * 0.3) + (
                bilancoDonemiFaaliyetKari * 4 * 0.5) + \
                           ((
                                        ucOncekibilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 0.2)
    print("Faaliyet Kar Tahmini2:", faaliyetKariTahmini2)

    ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1 + faaliyetKariTahmini2) / 2
    print("Ortalama Faaliyet Kari Tahmini:", ortalamaFaaliyetKariTahmini)

    hisseBasinaOrtalamaKarTahmini = (ortalamaFaaliyetKariTahmini * anaOrtaklikPayi) / sermaye
    print("Hisse başına ortalama kar tahmini:", format(hisseBasinaOrtalamaKarTahmini, ".2f"))

    likidasyonDegeri = likidasyonDegeriHesapla(bilancoDonemi)
    print("Likidasyon değeri:", likidasyonDegeri)

    borclar = int(getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", bilancoDonemiColumn))
    print("Borçlar:", format(borclar, ",").replace(',', '.'))

    bilancoEtkisi = (likidasyonDegeri - borclar) / sermaye * anaOrtaklikPayi
    print("Bilanço Etkisi:", format(bilancoEtkisi, ".2f"))

    gercekDeger = (hisseBasinaOrtalamaKarTahmini * 7) + bilancoEtkisi
    print("Gerçek hisse değeri:", format(gercekDeger, ".2f"))

    targetBuy = gercekDeger * 0.66
    print("Target buy:", format(targetBuy, ".2f"))

    # Netpro Hesapla
    print("----------------NetPro Kriteri-----------------------------------------------------------------")

    sonDortDonemFaaliyetKariToplami = bilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + ucOncekibilancoDonemiFaaliyetKari

    ucOncekibilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    bilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn)
    sonDortDonemNetKar = bilancoDonemiNetKari + birOncekiBilancoDonemiNetKari + ikiOncekiBilancoDonemiNetKari + ucOncekibilancoDonemiNetKari

    netProEstDegeri = ((
                                   ortalamaFaaliyetKariTahmini / sonDortDonemFaaliyetKariToplami) * sonDortDonemNetKar) * anaOrtaklikPayi
    print("NetPro Est Değeri:", netProEstDegeri)

    piyasaDegeri = (bilancoEtkisi * sermaye * -1) + (hisseFiyati * sermaye)

    netProKriteri = (netProEstDegeri / piyasaDegeri) / bondYield
    netProKriteriGecmeDurumu = (netProKriteri > 2)
    print("NetPro Kriteri:", format(netProKriteri, ".2f"), netProKriteriGecmeDurumu)

    # Forward PE Hesapla
    print("----------------Forward PE Kriteri-----------------------------------------------------------------")

    forwardPeKriteri = (piyasaDegeri) / netProEstDegeri

    forwardPeKriteriGecmeDurumu = (forwardPeKriteri < 4)
    print("Forward PE Kriteri:", format(forwardPeKriteri, ".2f"), forwardPeKriteriGecmeDurumu)

    def exportReportExcel(hisse, b):

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
            bookSheetWrite.write(0, 23, "Gerçek Hisse Değeri")
            bookSheetWrite.write(0, 24, "Target Buy")
            bookSheetWrite.write(0, 25, "NET Pro")
            bookSheetWrite.write(0, 26, "Forward PE")

        def reportResults(rowNumber):

            bookSheetWrite.write(rowNumber, 0, hisse)
            bookSheetWrite.write(rowNumber, 1, ceyrekDegeriHesapla(hasilatRow, bilancoDonemiColumn))
            bookSheetWrite.write(rowNumber, 2, ceyrekDegeriHesapla(hasilatRow, dortOncekibilancoDonemiColumn))
            bookSheetWrite.write(rowNumber, 3, kriter1SatisGelirArtisi)
            bookSheetWrite.write(rowNumber, 4, kriter3OncekiCeyrekArtisi)
            bookSheetWrite.write(rowNumber, 5, kriter1GecmeDurumu)
            bookSheetWrite.write(rowNumber, 6, kriter3GecmeDurumu)
            bookSheetWrite.write(rowNumber, 7, ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn))
            bookSheetWrite.write(rowNumber, 8, ceyrekDegeriHesapla(faaliyetKariRow, dortOncekibilancoDonemiColumn))
            bookSheetWrite.write(rowNumber, 9, kriter2FaaliyetKariArtisi)
            bookSheetWrite.write(rowNumber, 10, kriter4OncekiCeyrekFaaliyetKariArtisi)
            bookSheetWrite.write(rowNumber, 11, kriter2GecmeDurumu)
            bookSheetWrite.write(rowNumber, 12, kriter4GecmeDurumu)
            bookSheetWrite.write(rowNumber, 13, sermaye)
            bookSheetWrite.write(rowNumber, 14, anaOrtaklikPayi)
            bookSheetWrite.write(rowNumber, 15, sonDortCeyrekSatisToplami)
            bookSheetWrite.write(rowNumber, 16, onumuzdekiDortCeyrekSatisTahmini)
            bookSheetWrite.write(rowNumber, 17, onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini)
            bookSheetWrite.write(rowNumber, 18, faaliyetKariTahmini1)
            bookSheetWrite.write(rowNumber, 19, faaliyetKariTahmini2)
            bookSheetWrite.write(rowNumber, 20, ortalamaFaaliyetKariTahmini)
            bookSheetWrite.write(rowNumber, 21, hisseBasinaOrtalamaKarTahmini)
            bookSheetWrite.write(rowNumber, 22, bilancoEtkisi)
            bookSheetWrite.write(rowNumber, 23, gercekDeger)
            bookSheetWrite.write(rowNumber, 24, targetBuy)
            bookSheetWrite.write(rowNumber, 25, netProKriteri)
            bookSheetWrite.write(rowNumber, 26, forwardPeKriteri)

        if os.path.isfile(reportFile):
            print("Rapor dosyası var, güncellenecek:", reportFile)
            bookRead = xlrd.open_workbook(reportFile, formatting_info=True)
            sheetRead = bookRead.sheet_by_index(0)
            rowNumber = sheetRead.nrows
            bookWrite = copy(bookRead)
            bookSheetWrite = bookWrite.get_sheet(0)
            reportResults(rowNumber)
            bookWrite.save(reportFile)

        else:
            print("Rapor dosyası yeni oluşturulacak: ", reportFile)
            bookWrite = xlwt.Workbook()
            bookSheetWrite = bookWrite.add_sheet("2020_03")
            createTopRow()
            reportResults(1)
            bookWrite.save(reportFile)

    hisseAdiTemp = varBilancoDosyasi[19:]
    hisseAdi = hisseAdiTemp[:-5]
    print (hisseAdi)
    exportReportExcel(hisseAdi, reportFile)


#runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati)



directory = "D:\\bist\\bilancolar"
for filename in os.listdir(directory):
    varBilancoDosyasi = directory + "\\" + filename
    print (varBilancoDosyasi)
    runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati)