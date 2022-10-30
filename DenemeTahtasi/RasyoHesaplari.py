import xlrd
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri
import os

hisseAdi = "ASELS"
bilancoDonemi = 202209
directory = "//Users//myilmaz//Documents//bist//bilancolar"

def hesapla(varHisseAdi, varBilancoDonemi):

    print ("Hisse Adı: ", varHisseAdi)

    bilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"
    wb = xlrd.open_workbook(bilancoDosyasi)
    sheet = wb.sheet_by_index(0)

    def donemColumnFind(col):
        for columni in range(sheet.ncols):
            cell = sheet.cell(0, columni)
            if cell.value == col:
                return columni
        return -1

    def ceyrekDegeriHesapla(r, c):
        quarter = (sheet.cell_value(0, c)) % (100)
        if (quarter == 3):
            return sheet.cell_value(r, c)
        else:
            if (sheet.cell_value(0,c)-sheet.cell_value(0,(c-1)) == 3):
                return (sheet.cell_value(r, c) - sheet.cell_value(r, (c - 1)))
            else:
                return -1


    def getBilancoDegeri(label, column):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == label:
                if sheet.cell_value(rowi, column)=="":
                    # print (label + " :Bilanço alanı boş!")
                    return 0
                else:
                    return sheet.cell_value(rowi, column)
        # print ("Uygun bilanco degeri bulunamadi: %s", label)
        return 0

    def getBilancoTitleRow(title):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == title:
                return rowi
        return -1

    def birOncekiBilancoDoneminiHesapla(dnm):
        yil = int(dnm / 100)
        ceyrek = int(dnm % 100)

        if ceyrek == 3:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + (ceyrek - 3)

    hisseFiyati = returnGuncelHisseDegeri(varHisseAdi)


    birOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(bilancoDonemi)
    ikiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(birOncekiBilancoDonemi)
    ucOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ikiOncekiBilancoDonemi)
    dortOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(ucOncekiBilancoDonemi)
    besOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(dortOncekiBilancoDonemi)
    altiOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(besOncekiBilancoDonemi)
    yediOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(altiOncekiBilancoDonemi)
    sekizOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(yediOncekiBilancoDonemi)

    bilancoDonemiColumn = donemColumnFind(bilancoDonemi)
    birOncekibilancoDonemiColumn = donemColumnFind(birOncekiBilancoDonemi)
    ikiOncekibilancoDonemiColumn = donemColumnFind(ikiOncekiBilancoDonemi)
    ucOncekibilancoDonemiColumn = donemColumnFind(ucOncekiBilancoDonemi)
    dortOncekibilancoDonemiColumn = donemColumnFind(dortOncekiBilancoDonemi)
    besOncekibilancoDonemiColumn = donemColumnFind(besOncekiBilancoDonemi)
    altiOncekibilancoDonemiColumn = donemColumnFind(altiOncekiBilancoDonemi)
    yediOncekibilancoDonemiColumn = donemColumnFind(yediOncekiBilancoDonemi)

    # netKarRow = getBilancoTitleRow("DÖNEM KARI (ZARARI)");
    netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı");


    ucOncekibilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ucOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, ikiOncekibilancoDonemiColumn)
    birOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, birOncekibilancoDonemiColumn)
    bilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, bilancoDonemiColumn)
    sonDortDonemNetKarToplami = bilancoDonemiNetKari + birOncekiBilancoDonemiNetKari + ikiOncekiBilancoDonemiNetKari + ucOncekibilancoDonemiNetKari

    yediOncekibilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, yediOncekibilancoDonemiColumn)
    altiOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, altiOncekibilancoDonemiColumn)
    besOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, besOncekibilancoDonemiColumn)
    dortOncekiBilancoDonemiNetKari = ceyrekDegeriHesapla(netKarRow, dortOncekibilancoDonemiColumn)
    oncekiYilNetKarToplami = dortOncekiBilancoDonemiNetKari + besOncekiBilancoDonemiNetKari + altiOncekiBilancoDonemiNetKari + yediOncekibilancoDonemiNetKari

    anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", bilancoDonemiColumn) / getBilancoDegeri(
        "Net Dönem Karı veya Zararı", bilancoDonemiColumn)

    sermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)

    fkOrani = hisseFiyati / ((sonDortDonemNetKarToplami * anaOrtaklikPayi) / (sermaye))
    print ("F/K Orani: ", "{:,.2f}".format(fkOrani))

    hbkOrani = sonDortDonemNetKarToplami / (sermaye)
    # print ("HBK Oranı:", "{:,.2f}".format(hbkOrani))

    # print("Yıllık Net Kar: ", sonDortDonemNetKarToplami)
    # print("Önceki Yıl Net Kar: ", oncekiYilNetKarToplami)

    netKarBuyumeOraniYillik = (sonDortDonemNetKarToplami/oncekiYilNetKarToplami-1)
    print ("Yıllık Net Kar Büyüme: ", "{:.2%}".format(netKarBuyumeOraniYillik))
    oncekiYilAyniCeyregeGoreNetKarBuyume = (bilancoDonemiNetKari/dortOncekiBilancoDonemiNetKari-1)
    print("Önceki Yıl Aynı Çeyreğe Göre Net Kar Büyüme: ", "{:.2%}".format(oncekiYilAyniCeyregeGoreNetKarBuyume))



    pegOrani = fkOrani / (netKarBuyumeOraniYillik*100)
    print("PEG Orani: ", "{:,.2f}".format(pegOrani))

    piyasaDegeri = sermaye * hisseFiyati;
    # print ("Piyasa Değeri: ", piyasaDegeri)

    defterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", bilancoDonemiColumn)
    # print("Defter Değeri: ", defterDegeri)

    pddd = piyasaDegeri / defterDegeri
    print("PDDD: ", "{:,.2f}".format(pddd))


    # FD/FAVÖK Hesabı:

    brutKarRow = getBilancoTitleRow("BRÜT KAR (ZARAR)");
    pazarlamaGiderleriRow = getBilancoTitleRow("Pazarlama Giderleri")
    genelYonetimGiderleriRow = getBilancoTitleRow("Genel Yönetim Giderleri")
    argeGiderleriRow = getBilancoTitleRow("Araştırma ve Geliştirme Giderleri")
    amortismanlarRow = getBilancoTitleRow("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler")

    bilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiBrutKAr = ceyrekDegeriHesapla(brutKarRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiBrutKAr = ceyrekDegeriHesapla(brutKarRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiBrutKAr = ceyrekDegeriHesapla(brutKarRow, ucOncekibilancoDonemiColumn)
    yillikBrutKar = bilancoDonemiBrutKar + birOncekiBilancoDonemiBrutKAr + ikiOncekiBilancoDonemiBrutKAr + ucOncekiBilancoDonemiBrutKAr

    bilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, ucOncekibilancoDonemiColumn)
    yillikPazarlamaGiderleri = bilancoDonemiPazarlamaGiderleri + birOncekiBilancoDonemiPazarlamaGiderleri + ikiOncekiBilancoDonemiPazarlamaGiderleri + ucOncekiBilancoDonemiPazarlamaGiderleri

    bilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, ucOncekibilancoDonemiColumn)
    yillikGenelYonetimGiderleri = bilancoDonemiGenelYonetimGiderleri + birOncekiBilancoDonemiGenelYonetimGiderleri + ikiOncekiBilancoDonemiGenelYonetimGiderleri + ucOncekiBilancoDonemiGenelYonetimGiderleri

    bilancoDonemiArgeiderleri = ceyrekDegeriHesapla(argeGiderleriRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, ucOncekibilancoDonemiColumn)
    yillikArgeGiderleri = bilancoDonemiArgeiderleri + birOncekiBilancoDonemiArgeGiderleri + ikiOncekiBilancoDonemiArgeGiderleri + ucOncekiBilancoDonemiArgeGiderleri

    bilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, ucOncekibilancoDonemiColumn)
    yillikAmortisman = bilancoDonemiAmortisman + birOncekiBilancoDonemiAmortisman + ikiOncekiBilancoDonemiAmortisman + ucOncekiBilancoDonemiAmortisman

    favok = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
    print ("FAVÖK: ", "{:,.0f}".format(favok).replace(",","."))


    kisaVadeliFinansalBorclar = getBilancoDegeri("Kısa Vadeli Borçlanmalar", bilancoDonemiColumn) + getBilancoDegeri("Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", bilancoDonemiColumn)
    uzunVadeliFinansalBorclar = getBilancoDegeri("Uzun Vadeli Borçlanmalar", bilancoDonemiColumn)

    finansalBorclar = kisaVadeliFinansalBorclar + uzunVadeliFinansalBorclar

    nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", bilancoDonemiColumn)
    # finansalYatirimlar = getBilancoDegeri("Finansal Yatırımlar", bilancoDonemiColumn) + getBilancoDegeri("Finansal Yatırımlar1", bilancoDonemiColumn)
    finansalYatirimlar = getBilancoDegeri("Finansal Yatırımlar", bilancoDonemiColumn)

    netBorc = finansalBorclar - nakitVeNakitBenzerleri + finansalYatirimlar
    firmaDegeri = piyasaDegeri + netBorc

    print("Piyasa Değeri: ", "{:,.0f}".format(piyasaDegeri).replace(",",".") )
    print("Firma Değeri: ", "{:,.0f}".format(firmaDegeri).replace(",",".") )


hesapla(hisseAdi, bilancoDonemi)
