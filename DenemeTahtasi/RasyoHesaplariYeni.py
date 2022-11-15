import xlrd
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri
import os
import sys

hisseAdi = "USAK    "
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


    def birOncekiBilancoDoneminiHesapla(dnm):
        yil = int(dnm / 100)
        ceyrek = int(dnm % 100)

        if ceyrek == 3:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + (ceyrek - 3)


    def bilancoDoneminiBul (i):
        if (i > 0):
            print ("Hatalı Bilanço Dönemi!")
            return -999;
        elif (i == 0):
            return bilancoDonemi
        else:
            a = bilancoDonemi
            while (i<0):
                a = birOncekiBilancoDoneminiHesapla(a)
                i = i + 1
            return a


    def ceyrekDegeriHesapla(r, col):
        donem = bilancoDoneminiBul(col)
        c = donemColumnFind (donem)
        quarter = (sheet.cell_value(0, c)) % (100)
        if (quarter == 3):
            return sheet.cell_value(r, c)
        else:
            if (sheet.cell_value(0,c)-sheet.cell_value(0,(c-1)) == 3):
                return (sheet.cell_value(r, c) - sheet.cell_value(r, (c - 1)))
            else:
                return -1


    def getBilancoDegeri(label, col):

        column = donemColumnFind(bilancoDonemi) + col
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



    hisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
    print ("Güncel Hisse Fiyatı: ", hisseFiyati)


    # BİLANÇO KALEMLERİ TANIMLAMALARI
    # netKarRow = getBilancoTitleRow("DÖNEM KARI (ZARARI)")
    netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı")
    hasilatRow = getBilancoTitleRow("Hasılat")
    efkRow = getBilancoTitleRow("ESAS FAALİYET KARI (ZARARI)")




    sonDortDonemNetKarToplami = ceyrekDegeriHesapla(netKarRow, -3) + ceyrekDegeriHesapla(netKarRow, -2) + ceyrekDegeriHesapla(netKarRow, -1) + ceyrekDegeriHesapla(netKarRow, 0)
    oncekiYilNetKarToplami = ceyrekDegeriHesapla(netKarRow, -7) + ceyrekDegeriHesapla(netKarRow, -6) + ceyrekDegeriHesapla(netKarRow, -5) + ceyrekDegeriHesapla(netKarRow, -4)

    anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", 0) / getBilancoDegeri("Net Dönem Karı veya Zararı", 0)
    sermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)

    # print("Yıllık Net Kar: ", sonDortDonemNetKarToplami)
    # print("Önceki Yıl Net Kar: ", oncekiYilNetKarToplami)

    netKarBuyumeOraniYillik = (sonDortDonemNetKarToplami/oncekiYilNetKarToplami-1)
    print ("Yıllık Net Kar Büyüme: ", "{:.2%}".format(netKarBuyumeOraniYillik))
    oncekiYilAyniCeyregeGoreNetKarBuyume = (ceyrekDegeriHesapla(netKarRow, 0)/ceyrekDegeriHesapla(netKarRow, -4) - 1)
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
    nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
    print ("Nakit ve Nakit Benzerleri: ", "{:,.0f}".format(nakitVeNakitBenzerleri).replace(",", "."))
    print ("Nakit / PD: ", "{:,.2f}".format(nakitVeNakitBenzerleri/piyasaDegeri))

    defterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
    dortOncekiCeyrekDefterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", -4)

    pddd = piyasaDegeri / defterDegeri
    print("PD/DD: ", "{:,.2f}".format(pddd))

    pegOrani = fkOrani / (netKarBuyumeOraniYillik*100)
    print("PEG Orani: ", "{:,.4f}".format(pegOrani))


    # Firma Degeri Hesabi
    kisaVadeliFinansalBorclar = getBilancoDegeri("Kısa Vadeli Borçlanmalar", 0) + getBilancoDegeri("Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", 0)
    uzunVadeliFinansalBorclar = getBilancoDegeri("Uzun Vadeli Borçlanmalar", 0)
    finansalBorclar = kisaVadeliFinansalBorclar + uzunVadeliFinansalBorclar
    nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
    # finansalYatirimlar = getBilancoDegeri("Finansal Yatırımlar", bilancoDonemiColumn) + getBilancoDegeri("Finansal Yatırımlar1", bilancoDonemiColumn)
    finansalYatirimlar = getBilancoDegeri("Finansal Yatırımlar", 0)
    netBorc = finansalBorclar - nakitVeNakitBenzerleri - finansalYatirimlar
    firmaDegeri = piyasaDegeri + netBorc
    print ("Firma Değeri (FD): ", "{:,.0f}".format(firmaDegeri).replace(",","."))
    print("Nakit / FD: ", "{:,.2f}".format(nakitVeNakitBenzerleri / firmaDegeri))


    # Yillik Hasilat Hesabi
    yillikHasilat = ceyrekDegeriHesapla(hasilatRow, 0) + ceyrekDegeriHesapla(hasilatRow, -1) + ceyrekDegeriHesapla(hasilatRow, -2) + ceyrekDegeriHesapla(hasilatRow, -3)

    # FD/Satislar
    fdSatislar = firmaDegeri / yillikHasilat
    print ("FD/Satışlar: ", "{:,.2f}".format(fdSatislar))


    # FAVÖK Hesabı:

    brutKarRow = getBilancoTitleRow("BRÜT KAR (ZARAR)");
    pazarlamaGiderleriRow = getBilancoTitleRow("Pazarlama Giderleri")
    genelYonetimGiderleriRow = getBilancoTitleRow("Genel Yönetim Giderleri")
    argeGiderleriRow = getBilancoTitleRow("Araştırma ve Geliştirme Giderleri")
    amortismanlarRow = getBilancoTitleRow("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler")

    bilancoDonemiBrutKar = ceyrekDegeriHesapla(brutKarRow, 0)
    birOncekiBilancoDonemiBrutKAr = ceyrekDegeriHesapla(brutKarRow, -1)
    ikiOncekiBilancoDonemiBrutKAr = ceyrekDegeriHesapla(brutKarRow, -2)
    ucOncekiBilancoDonemiBrutKAr = ceyrekDegeriHesapla(brutKarRow, -3)
    yillikBrutKar = bilancoDonemiBrutKar + birOncekiBilancoDonemiBrutKAr + ikiOncekiBilancoDonemiBrutKAr + ucOncekiBilancoDonemiBrutKAr

    bilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, 0)
    birOncekiBilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, -1)
    ikiOncekiBilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, -2)
    ucOncekiBilancoDonemiGenelYonetimGiderleri = ceyrekDegeriHesapla(genelYonetimGiderleriRow, -3)
    yillikGenelYonetimGiderleri = bilancoDonemiGenelYonetimGiderleri + birOncekiBilancoDonemiGenelYonetimGiderleri + ikiOncekiBilancoDonemiGenelYonetimGiderleri + ucOncekiBilancoDonemiGenelYonetimGiderleri

    try:
        bilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, 0)
        birOncekiBilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, -1)
        ikiOncekiBilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, -2)
        ucOncekiBilancoDonemiPazarlamaGiderleri = ceyrekDegeriHesapla(pazarlamaGiderleriRow, -3)
        yillikPazarlamaGiderleri = bilancoDonemiPazarlamaGiderleri + birOncekiBilancoDonemiPazarlamaGiderleri + ikiOncekiBilancoDonemiPazarlamaGiderleri + ucOncekiBilancoDonemiPazarlamaGiderleri
    except Exception as e:
        print("Bilançoda Pazarlama Giderleri Bulunmamaktadır!")
        yillikPazarlamaGiderleri    = 0


    try:
        bilancoDonemiArgeiderleri = ceyrekDegeriHesapla(argeGiderleriRow, 0)
        birOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, -1)
        ikiOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, -2)
        ucOncekiBilancoDonemiArgeGiderleri = ceyrekDegeriHesapla(argeGiderleriRow, -3)
        yillikArgeGiderleri = bilancoDonemiArgeiderleri + birOncekiBilancoDonemiArgeGiderleri + ikiOncekiBilancoDonemiArgeGiderleri + ucOncekiBilancoDonemiArgeGiderleri

    except Exception as e:
        print("Bilançoda AR-GE Giderleri Bulunmamaktadır!")
        yillikArgeGiderleri = 0


    try:
        yillikAmortisman = ceyrekDegeriHesapla(amortismanlarRow, 0) + ceyrekDegeriHesapla(amortismanlarRow,-1) + ceyrekDegeriHesapla(amortismanlarRow, -2) + ceyrekDegeriHesapla(amortismanlarRow, -3)
    except Exception as e:
            print("Bilançoda YILLIK AMORTİSMAN Bulunmamaktadır!")
            yillikAmortisman = 0



    yillikNetEsasFaaliyetKari = yillikBrutKar + yillikGenelYonetimGiderleri + yillikPazarlamaGiderleri + yillikArgeGiderleri

    favok = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
    print ("FAVÖK: ", "{:,.0f}".format(favok).replace(",","."))


    # FD/FAVOK Hesabi
    fdfavok = firmaDegeri/favok
    print("FD/FAVÖK: ", "{:,.2f}".format(fdfavok))

    # EFK Hesabi
    print ("Yıllık EFK: ", "{:,.0f}".format(yillikNetEsasFaaliyetKari).replace(",","."))

    #PD/EFK Hesabi
    pdefk = piyasaDegeri / yillikNetEsasFaaliyetKari
    print ("PD/EFK: ""{:,.2f}".format(pdefk))

    #HBK Hesabi
    hbk = sonDortDonemNetKarToplami / (sermaye)
    print ("HBK:", "{:,.2f}".format(hbk))

    #Ödenmiş Sermaye
    odenmisSermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)
    print("Ödenmiş Sermaye: ", "{:,.0f}".format(odenmisSermaye).replace(",", "."))

    #Net Borc
    print("Net Borç: ", "{:,.0f}".format(netBorc).replace(",", "."))







    #DIGER Rasyolar
    print("")
    print("---------- DİĞER ORANLAR ----------")


    # Cari Oran Hesabı
    donenVarliklar = getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", 0)
    kisaVadeliYukumlulukler = getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
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
    bilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", 0)
    dortOncekiBilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", -4)
    toplamVarliklar = (bilancoDonemiToplamVarliklar + dortOncekiBilancoDonemiToplamVarliklar) / 2
    aktifKarlilik = sonDortDonemNetKarToplami / toplamVarliklar
    print("ROA (Aktif Karlılık): ", "{:.2%}".format(aktifKarlilik))

    # Kar Marjı Hesabı
    netKarMarji = sonDortDonemNetKarToplami / yillikHasilat
    sonCeyrekNetKarMarji = ceyrekDegeriHesapla(netKarRow, 0) / ceyrekDegeriHesapla(hasilatRow, -0)
    print ("Yıllık Net Kar Marjı: ", "{:.2%}".format(netKarMarji))
    print("Son Çeyrek Net Kar Marjı: ", "{:.2%}".format(sonCeyrekNetKarMarji))

    # Aktif Devir Hızı Hesabı
    aktifDevirHizi = yillikHasilat / toplamVarliklar
    print ("Aktif Devir Hızı: ", "{:.2}".format(aktifDevirHizi))

    # Borç Kaynak Oranı Hesabı
    borcKaynakOrani = getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", 0) / getBilancoDegeri("TOPLAM KAYNAKLAR", 0)
    print("Borç/Kaynak Oranı: ", "{:.2%}".format(borcKaynakOrani))

    # ROIC Hesabı
    # ROIC = ((FVÖK * (1 - Vergi Oranı)) / (Alacak + Özsermaye)))
    nopat = getBilancoDegeri("BRÜT KAR (ZARAR)", 0)



hesapla(hisseAdi, bilancoDonemi)
