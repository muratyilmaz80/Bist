import xlrd
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri
import os
import sys

hisseAdi = "TUPRS" \
           ""
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
        return 0

    def getBilancoDegeriYeni(label, col):
        column = donemColumnFind(bilancoDonemi)+col
        print ("Bulunan Column: ", column)

        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == label:
                if sheet.cell_value(rowi, column)=="":
                    # print (label + " :Bilanço alanı boş!")
                    return 0
                else:
                    return sheet.cell_value(rowi, column)
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
    print ("Güncel Hisse Fiyatı: ", hisseFiyati)

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

    # netKarRow = getBilancoTitleRow("DÖNEM KARI (ZARARI)")
    netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı")
    hasilatRow = getBilancoTitleRow("Hasılat")
    efkRow = getBilancoTitleRow("ESAS FAALİYET KARI (ZARARI)")

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

    # print("Yıllık Net Kar: ", sonDortDonemNetKarToplami)
    # print("Önceki Yıl Net Kar: ", oncekiYilNetKarToplami)

    netKarBuyumeOraniYillik = (sonDortDonemNetKarToplami/oncekiYilNetKarToplami-1)
    print ("Yıllık Net Kar Büyüme: ", "{:.2%}".format(netKarBuyumeOraniYillik))
    oncekiYilAyniCeyregeGoreNetKarBuyume = (bilancoDonemiNetKari/dortOncekiBilancoDonemiNetKari-1)
    print("Önceki Yıl Aynı Çeyreğe Göre Net Kar Büyüme: ", "{:.2%}".format(oncekiYilAyniCeyregeGoreNetKarBuyume))


    # TEMEL CARPANLAR
    print("")
    print ("---------- TEMEL ORANLAR ----------")

    # F/K
    fkOrani = hisseFiyati / ((sonDortDonemNetKarToplami * anaOrtaklikPayi) / (sermaye))
    print ("F/K Orani: ", "{:,.2f}".format(fkOrani))

    # PD/DD
    piyasaDegeri = sermaye * hisseFiyati;
    print("Piyasa Değeri (PD): ", "{:,.0f}".format(piyasaDegeri).replace(",", "."))
    nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", bilancoDonemiColumn)
    print ("Nakit ve Nakit Benzerleri: ", "{:,.0f}".format(nakitVeNakitBenzerleri).replace(",", "."))
    print ("Nakit / PD: ", "{:,.2f}".format(nakitVeNakitBenzerleri/piyasaDegeri))

    defterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", bilancoDonemiColumn)
    dortOncekiCeyrekDefterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", dortOncekibilancoDonemiColumn)

    pddd = piyasaDegeri / defterDegeri
    print("PD/DD: ", "{:,.2f}".format(pddd))

    pegOrani = fkOrani / (netKarBuyumeOraniYillik*100)
    print("PEG Orani: ", "{:,.4f}".format(pegOrani))


    # Firma Degeri Hesabi
    kisaVadeliFinansalBorclar = getBilancoDegeri("Kısa Vadeli Borçlanmalar", bilancoDonemiColumn) + getBilancoDegeri("Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", bilancoDonemiColumn)
    uzunVadeliFinansalBorclar = getBilancoDegeri("Uzun Vadeli Borçlanmalar", bilancoDonemiColumn)
    finansalBorclar = kisaVadeliFinansalBorclar + uzunVadeliFinansalBorclar
    nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", bilancoDonemiColumn)
    # finansalYatirimlar = getBilancoDegeri("Finansal Yatırımlar", bilancoDonemiColumn) + getBilancoDegeri("Finansal Yatırımlar1", bilancoDonemiColumn)
    finansalYatirimlar = getBilancoDegeri("Finansal Yatırımlar", bilancoDonemiColumn)
    netBorc = finansalBorclar - nakitVeNakitBenzerleri - finansalYatirimlar
    firmaDegeri = piyasaDegeri + netBorc
    print ("Firma Değeri (FD): ", "{:,.0f}".format(firmaDegeri).replace(",","."))
    print("Nakit / FD: ", "{:,.2f}".format(nakitVeNakitBenzerleri / firmaDegeri))


    # Yillik Hasilat Hesabi
    bilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiHasilat = ceyrekDegeriHesapla(hasilatRow, ucOncekibilancoDonemiColumn)
    yillikHasilat = bilancoDonemiHasilat + birOncekiBilancoDonemiHasilat + ikiOncekiBilancoDonemiHasilat + ucOncekiBilancoDonemiHasilat

    # FD/Satislar
    fdSatislar = firmaDegeri / yillikHasilat
    print ("FD/Satışlar: ", "{:,.2f}".format(fdSatislar))


    # FAVÖK Hesabı:

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

    try:
        bilancoDonemiArgeiderleri = ceyrekDegeriHesapla(argeGiderleriRow, bilancoDonemiColumn)
        birOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, birOncekibilancoDonemiColumn)
        ikiOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, ikiOncekibilancoDonemiColumn)
        ucOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, ucOncekibilancoDonemiColumn)
        yillikArgeGiderleri = bilancoDonemiArgeiderleri + birOncekiBilancoDonemiArgeGiderleri + ikiOncekiBilancoDonemiArgeGiderleri + ucOncekiBilancoDonemiArgeGiderleri

    except Exception as e:
        print("Bilançoda AR-GE Giderleri Bulunmamaktadır!")
        yillikArgeGiderleri = 0


    bilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiAmortisman = ceyrekDegeriHesapla(amortismanlarRow, ucOncekibilancoDonemiColumn)
    yillikAmortisman = bilancoDonemiAmortisman + birOncekiBilancoDonemiAmortisman + ikiOncekiBilancoDonemiAmortisman + ucOncekiBilancoDonemiAmortisman

    favok = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
    print ("FAVÖK: ", "{:,.0f}".format(favok).replace(",","."))


    # FD/FAVOK Hesabi
    fdfavok = firmaDegeri/favok
    print("FD/FAVÖK: ", "{:,.2f}".format(fdfavok))

    # EFK Hesabi
    bilancoDonemiEfk = ceyrekDegeriHesapla(efkRow, bilancoDonemiColumn)
    birOncekiBilancoDonemiEfk = ceyrekDegeriHesapla(efkRow, birOncekibilancoDonemiColumn)
    ikiOncekiBilancoDonemiEfk = ceyrekDegeriHesapla(efkRow, ikiOncekibilancoDonemiColumn)
    ucOncekiBilancoDonemiEfk = ceyrekDegeriHesapla(efkRow, ucOncekibilancoDonemiColumn)
    yillikEfk = bilancoDonemiEfk + birOncekiBilancoDonemiEfk + ikiOncekiBilancoDonemiEfk + ucOncekiBilancoDonemiEfk
    print ("Yıllık EFK: ", "{:,.0f}".format(yillikEfk).replace(",","."))

    #PD/EFK Hesabi
    pdefk = piyasaDegeri / yillikEfk
    print ("PD/EFK: ""{:,.2f}".format(pdefk))

    #HBK Hesabi
    hbk = sonDortDonemNetKarToplami / (sermaye)
    print ("HBK:", "{:,.2f}".format(hbk))

    #Ödenmiş Sermaye
    odenmisSermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)
    print("Ödenmiş Sermaye: ", "{:,.0f}".format(odenmisSermaye).replace(",", "."))

    #Net Borc
    print("Net Borç: ", "{:,.0f}".format(netBorc).replace(",", "."))







    #DIGER Rasyolar
    print("")
    print("---------- DİĞER ORANLAR ----------")


    # Cari Oran Hesabı
    donenVarliklar = getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", bilancoDonemiColumn)
    kisaVadeliYukumlulukler = getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", bilancoDonemiColumn)
    cariOran = donenVarliklar / kisaVadeliYukumlulukler
    print("Cari Oran: ", "{:.3}".format(cariOran))


    #DUPONT ANALİZİ ORANLARI
    print("")
    print("---------- DUPONT ANALİZİ ORANLARI ----------")


    # Özsermaye Karlılığı (ROE) Hesabı
    ortDefterDegeri = (defterDegeri + dortOncekiCeyrekDefterDegeri) / 2
    roe = sonDortDonemNetKarToplami/ortDefterDegeri
    # print("Yıllık Net Kar: ", "{:,.0f}".format(sonDortDonemNetKarToplami).replace(",", "."))
    # print("Özsermaye: ", "{:,.0f}".format(defterDegeri).replace(",", "."))
    # print("Yıllık Ortalama Özsermaye: ", "{:,.0f}".format(ortDefterDegeri).replace(",", "."))
    print("ROE (Özsermaye Karlılığı - Özkaynak Getirisi): ", "{:.2%}".format(roe))

    # Aktif Karlılık Hesabı
    bilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", bilancoDonemiColumn)
    dortOncekiBilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", dortOncekibilancoDonemiColumn)
    toplamVarliklar = (bilancoDonemiToplamVarliklar + dortOncekiBilancoDonemiToplamVarliklar) / 2
    aktifKarlilik = sonDortDonemNetKarToplami / toplamVarliklar
    print("ROA (Aktif Karlılık): ", "{:.2%}".format(aktifKarlilik))

    # Kar Marjı Hesabı
    netKarMarji = sonDortDonemNetKarToplami / yillikHasilat
    sonCeyrekNetKarMarji = bilancoDonemiNetKari/bilancoDonemiHasilat
    print ("Yıllık Net Kar Marjı: ", "{:.2%}".format(netKarMarji))
    print("Son Çeyrek Net Kar Marjı: ", "{:.2%}".format(sonCeyrekNetKarMarji))

    # Aktif Devir Hızı Hesabı
    aktifDevirHizi = yillikHasilat / toplamVarliklar
    print ("Aktif Devir Hızı: ", "{:.2}".format(aktifDevirHizi))

    # Borç Kaynak Oranı Hesabı
    borcKaynakOrani = getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", bilancoDonemiColumn) / getBilancoDegeri("TOPLAM KAYNAKLAR", bilancoDonemiColumn)
    print("Borç/Kaynak Oranı: ", "{:.2%}".format(borcKaynakOrani))

    deneme = getBilancoDegeriYeni("TOPLAM YÜKÜMLÜLÜKLER", -1)
    print ("Toplam Yükümlülükler: ", deneme)



hesapla(hisseAdi, bilancoDonemi)
