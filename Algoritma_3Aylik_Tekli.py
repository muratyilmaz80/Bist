import xlrd
import xlwt
from xlutils.copy import copy
import os.path
from ExcelRowClass import ExcelRowClass
from Rapor_Olustur import exportReportExcel

varBilancoDosyasi = ("//Users//myilmaz//Documents//bist//bilancolar//SEYKM.xlsx")

# AKSUE BNTAS DAGI DOBUR EGSER JANTS TIRE

varBilancoDonemi = 202009
varBondYield = 0.1448
varHisseFiyati = 12.72
varReportFile = "//Users//myilmaz//Documents//bist//Report_2020_09_3Ayliklar_Mac.xls"


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
    # netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı");
    netKarRow = getBilancoTitleRow("DÖNEM KARI (ZARARI)");

    def ceyrekDegeriHesapla(r, c):
        quarter = (sheet.cell_value(0, c)) % (100)
        if (quarter == 3):
            return sheet.cell_value(r, c)
        else:
            return (sheet.cell_value(r, c) - sheet.cell_value(r, (c - 1)))

    def oncekiYilAyniCeyrekDegisimiHesapla(row, donem):
        donemColumn = donemColumnFind(donem)
        #print ("DonemColumn:", donemColumn)
        oncekiYilAyniDonemColumn = donemColumnFind(donem - 100)
        #print("Onceki Yıl Aynı DonemColumn:", oncekiYilAyniDonemColumn)
        #print("Row:",row, "Column:", donemColumn)
        ceyrekDegeri = ceyrekDegeriHesapla(row, donemColumn)
        #print("Çeyrek Değeri:", ceyrekDegeri)
        oncekiCeyrekDegeri = ceyrekDegeriHesapla(row, oncekiYilAyniDonemColumn)
        #print ("Önceki Çeyrek Değeri:", oncekiCeyrekDegeri)
        degisimSonucu = ceyrekDegeri / oncekiCeyrekDegeri - 1
        print(int(sheet.cell_value(0, donemColumn)), sheet.cell_value(row, 0), "{:,.0f}".format(ceyrekDegeri).replace(",","."), "TL")
        print(int(sheet.cell_value(0, oncekiYilAyniDonemColumn)), sheet.cell_value(row, 0), "{:,.0f}".format(oncekiCeyrekDegeri).replace(",","."), "TL")
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
    print("1.Kriter: Satış gelirleri bir önceki yılın aynı dönemine göre en az %10 artmalı")

    kriter1SatisGelirArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    kriter1GecmeDurumu = (kriter1SatisGelirArtisi > 0.1)
    print("Kriter1: Satis Geliri Artisi:", "{:.2%}".format(kriter1SatisGelirArtisi), ">? 10%", kriter1GecmeDurumu)

    # 2.kriter hesabi
    print("---------------------------------------------------------------------------------")
    print("2.Kriter: Son ceyrek faaliyet kari onceki yil ayni ceyrege göre en az %15 fazla olacak")

    if ceyrekDegeriHesapla(netKarRow,bilancoDonemiColumn)<0:
        kriter2FaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        kriter2GecmeDurumu = False
        print("Kriter2: Faaliyet Kari Artisi:", kriter2GecmeDurumu, "Son Ceyrek Net Kar Negatif")

    elif ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) < 0:
        kriter2FaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        kriter2GecmeDurumu = False
        print("Kriter2: Faaliyet Kari Artisi:", kriter2GecmeDurumu, "Son Ceyrek Faaliyet Kari Negatif")

    elif ((ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn) > 0) and (oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi))<0 ):
        kriter2FaaliyetKariArtisi = 0
        kriter2GecmeDurumu = True
        print("Kriter2: Faaliyet Kari Artisi:", kriter2GecmeDurumu, "Son Ceyrek Faaliyet Kari Negatiften Pozitife Geçmiş")

    else:
        kriter2FaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, bilancoDonemi)
        kriter2GecmeDurumu = (kriter2FaaliyetKariArtisi > 0.15)
        print("Kriter2: Faaliyet Kari Artisi:", "{:.2%}".format(kriter2FaaliyetKariArtisi), ">? 15%", kriter2GecmeDurumu)


    # 3.kriter hesabı
    print("---------------------------------------------------------------------------------")
    print("3.Kriter: Bir önceki çeyrekteki satış artış yüzdesi cari dönemden düşük olmalı")

    if kriter1SatisGelirArtisi >= 1:
        kriter3OncekiCeyrekArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)
        kriter3GecmeDurumu = True
        print("Kriter3: Onceki Ceyrek Satis Geliri Artisi %100'ün Üzerinde, Karşılaştırma Yapılmayacak!:", "{:.2%}".format(kriter3OncekiCeyrekArtisi), "<",
              "{:.2%}".format(kriter1SatisGelirArtisi), kriter3GecmeDurumu)

    else:
        kriter3OncekiCeyrekArtisi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)
        kriter3GecmeDurumu = (kriter3OncekiCeyrekArtisi < kriter1SatisGelirArtisi)
        print("Kriter3: Onceki Ceyrek Satis Geliri Artisi:", "{:.2%}".format(kriter3OncekiCeyrekArtisi),"<?","{:.2%}".format(kriter1SatisGelirArtisi), kriter3GecmeDurumu)


    # 4.kriter hesabi
    print("---------------------------------------------------------------------------------")
    print("4.Kriter: Bir önceki çeyrekteki faaliyet karı artış yüzdesi cari dönemden düşük olmalı")

    if kriter2FaaliyetKariArtisi >= 1:
        kriter4OncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow,
                                                                                   birOncekiBilancoDonemi)
        kriter4GecmeDurumu = True
        print("Kriter4: Onceki Ceyrek Faaliyet Kari Artisi %100'ün Üzerinde, Karşılaştırma Yapılmayacak:", "{:.2%}".format(kriter4OncekiCeyrekFaaliyetKariArtisi),
              "<?", "{:.2%}".format(kriter2FaaliyetKariArtisi), kriter4GecmeDurumu)


    else:
        kriter4OncekiCeyrekFaaliyetKariArtisi = oncekiYilAyniCeyrekDegisimiHesapla(faaliyetKariRow, birOncekiBilancoDonemi)
        kriter4GecmeDurumu = (kriter4OncekiCeyrekFaaliyetKariArtisi < kriter2FaaliyetKariArtisi)
        print("Kriter4: Onceki Yila Gore Faaliyet Kari Artisi:", "{:.2%}".format(kriter4OncekiCeyrekFaaliyetKariArtisi),
          "<?" , "{:.2%}".format(kriter2FaaliyetKariArtisi) , kriter4GecmeDurumu)

    # Gercek Deger Hesapla
    print("----------------Gercek Deger Hesabi-----------------------------------------------------------------")

    sermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)
    print("Sermaye:", "{:,.0f}".format(sermaye).replace(",","."), "TL")

    anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", bilancoDonemiColumn) / getBilancoDegeri(
        "DÖNEM KARI (ZARARI)", bilancoDonemiColumn)
    print("Ana Ortaklık Payı:", "{:.2f}".format(anaOrtaklikPayi))

    sonCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
    birOncekiCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

    ucOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, birOncekibilancoDonemiColumn)
    bilancoDonemiSatis = ceyrekDegeriHesapla(hasilatRow, bilancoDonemiColumn)

    sonDortCeyrekSatisToplami = ucOncekiBilancoDonemiSatis + ikiOncekiBilancoDonemiSatis + birOncekiBilancoDonemiSatis + bilancoDonemiSatis
    print("Son 4 ceyrek satış toplamı:", "{:,.0f}".format(sonDortCeyrekSatisToplami).replace(",","."), "TL")

    onumuzdekiDortCeyrekSatisTahmini = (
                (((sonCeyrekSatisArtisYuzdesi + birOncekiCeyrekSatisArtisYuzdesi) / 2) + 1) * sonDortCeyrekSatisToplami)
    print("Önümüzdeki 4 çeyrek satış tahmini:", "{:,.0f}".format(onumuzdekiDortCeyrekSatisTahmini).replace(",","."), "TL")

    #onumuzdekiDortCeyrekSatisTahmini = 5000000000

    ucOncekibilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, birOncekibilancoDonemiColumn)
    bilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)

    onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) / (
                bilancoDonemiSatis + birOncekiBilancoDonemiSatis)
    print("Önümüzdeki 4 çeyrek faaliyet kar marjı tahmini:",
          "{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))

    faaliyetKariTahmini1 = onumuzdekiDortCeyrekSatisTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
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

    print("Bilanço tarihindeki hisse fiyatı:", format(varHisseFiyati, ".2f"), "TL")

    gercekFiyataUzaklik = hisseFiyati / targetBuy
    print("Gerçek fiyata uzaklık:", "{:.2%}".format(gercekFiyataUzaklik))

    # Netpro Hesapla
    print("----------------NetPro Kriteri-----------------------------------------------------------------")

    sonDortDonemFaaliyetKariToplami = bilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + ucOncekibilancoDonemiFaaliyetKari

    ucOncekibilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    bilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn)

    print ("Bilanco Donem Net Kar:", bilancoDonemiNetKari)
    print("birOncekiBilancoDonemiNetKari:", birOncekiBilancoDonemiNetKari)
    print("ikiOncekiBilancoDonemiNetKari:", ikiOncekiBilancoDonemiNetKari)
    print("ucOncekibilancoDonemiNetKari:", ucOncekibilancoDonemiNetKari)

    sonDortDonemNetKar = bilancoDonemiNetKari + birOncekiBilancoDonemiNetKari + ikiOncekiBilancoDonemiNetKari + ucOncekibilancoDonemiNetKari

    print ("Son 4 Dönem Net Kar:", sonDortDonemNetKar)
    print ("Son 4 Dönem Faaliyet Kari Toplami:", sonDortDonemFaaliyetKariToplami)

    netProEstDegeri = ((
                                   ortalamaFaaliyetKariTahmini / sonDortDonemFaaliyetKariToplami) * sonDortDonemNetKar) * anaOrtaklikPayi
    print("NetPro Est Değeri:", "{:,.0f}".format(netProEstDegeri).replace(",","."), "TL")

    piyasaDegeri = (bilancoEtkisi * sermaye * -1) + (hisseFiyati * sermaye)

    print("Piyasa Değeri:", piyasaDegeri)
    print("BondYield:", bondYield)

    netProKriteri = (netProEstDegeri / piyasaDegeri) / bondYield
    netProKriteriGecmeDurumu = (netProKriteri > 2)
    print("NetPro Kriteri:", format(netProKriteri, ".2f"), netProKriteriGecmeDurumu)

    minNetProIcınHisseFiyati  = (netProEstDegeri / (1.9 * bondYield) - (bilancoEtkisi * sermaye * -1))/sermaye
    print("NetPro 1.9 Olmasi Icin Hisse Degeri:", format(minNetProIcınHisseFiyati, ".2f"), )

    # Forward PE Hesapla
    print("----------------Forward PE Kriteri-----------------------------------------------------------------")

    forwardPeKriteri = (piyasaDegeri) / netProEstDegeri

    forwardPeKriteriGecmeDurumu = (forwardPeKriteri < 4)
    print("Forward PE Kriteri:", format(forwardPeKriteri, ".2f"), forwardPeKriteriGecmeDurumu)

    hisseAdiTemp = varBilancoDosyasi[47:]
    hisseAdi = hisseAdiTemp[:-5]
    print (hisseAdi)

    excelRow = ExcelRowClass()

    excelRow.sonCeyrekHasilat = ceyrekDegeriHesapla(hasilatRow, bilancoDonemiColumn)
    excelRow.oncekiYilAyniCeyrekHasilat = ceyrekDegeriHesapla(hasilatRow, dortOncekibilancoDonemiColumn)
    excelRow.hasilatArtisi = kriter1SatisGelirArtisi
    excelRow.birOncekiCeyrekHasilatArtisi = kriter3OncekiCeyrekArtisi
    excelRow.kriter1 = kriter1GecmeDurumu
    excelRow.kriter3 = kriter3GecmeDurumu
    excelRow.sonCeyrekFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)
    excelRow.oncekiYilAyniCeyrekFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, dortOncekibilancoDonemiColumn)
    excelRow.faaliyetKarArtisi = kriter2FaaliyetKariArtisi
    excelRow.birOncekiCeyrekFaaliyetKarArtisi = kriter4OncekiCeyrekFaaliyetKariArtisi
    excelRow.kriter2 = kriter2GecmeDurumu
    excelRow.kriter4 = kriter4GecmeDurumu
    excelRow.sermaye = sermaye
    excelRow.anaOrtaklikPayi = anaOrtaklikPayi
    excelRow.son4CeyrekSatisToplami = sonDortCeyrekSatisToplami
    excelRow.onumuzdeki4CeyrekSatisTahmini = onumuzdekiDortCeyrekSatisTahmini
    excelRow.onumuzdeki4CeyrekFaaliyetKarMarjiTahmini = onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
    excelRow.faaliyetKarTahmini1 = faaliyetKariTahmini1
    excelRow.faaliyetKarTahmini2 = faaliyetKariTahmini2
    excelRow.ortalamaFaaliyetKarTahmini = ortalamaFaaliyetKariTahmini
    excelRow.hisseBasinaKarTahmini = hisseBasinaOrtalamaKarTahmini
    excelRow.bilancoEtkisi = bilancoEtkisi
    excelRow.bilancoTarihiHisseFiyati = varHisseFiyati
    excelRow.gercekHisseDegeri = gercekDeger
    excelRow.targetBuy = targetBuy
    excelRow.gercekFiyataUzaklik = gercekFiyataUzaklik
    excelRow.netPro = netProKriteri
    excelRow.forwardPe = forwardPeKriteri

    exportReportExcel(hisseAdi, varReportFile, varBilancoDonemi, excelRow)

runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati)