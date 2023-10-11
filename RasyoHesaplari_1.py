
from GetGuncelHisseDegeri import returnGuncelHisseDegeri
from GetHisseHalkaAciklikOrani import returnHisseHalkaAciklikOrani
import os
import xlrd
import xlwt
from xlutils.copy import copy
import os.path
from datetime import datetime

varReportFile = "//Users//myilmaz//Documents//bist//Report_202306_Rasyolar.xls"

hisseAdi = "SISE"
bilancoDonemi = 202306
directory = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar"
bilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + hisseAdi + ".xlsx"
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


def yilliklandirmisDegerHesapla (row, bd):
    toplam = ceyrekDegeriHesapla(row, bd) + ceyrekDegeriHesapla(row, bd-1) + ceyrekDegeriHesapla(row, bd-2) + ceyrekDegeriHesapla(row, bd-3)
    return toplam


def hesapla(varHisseAdi, varBilancoDonemi):

    print ("Hisse Adı: ", varHisseAdi)
    hisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
    print ("Güncel Hisse Fiyatı: ", hisseFiyati)

    # BİLANÇO KALEMLERİ TANIMLAMALARI
    netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı")
    hasilatRow = getBilancoTitleRow("Hasılat")
    esasFaaliyetKariRow = getBilancoTitleRow("ESAS FAALİYET KARI (ZARARI)")
    donemVergiGideriRow = getBilancoTitleRow("Dönem Vergi (Gideri) Geliri")
    ertelenmisVergiGideriRow = getBilancoTitleRow("Ertelenmiş Vergi (Gideri) Geliri")

    yillikNetKar = ceyrekDegeriHesapla(netKarRow, -3) + ceyrekDegeriHesapla(netKarRow, -2) + ceyrekDegeriHesapla(netKarRow, -1) + ceyrekDegeriHesapla(netKarRow, 0)
    oncekiYilNetKarToplami = ceyrekDegeriHesapla(netKarRow, -7) + ceyrekDegeriHesapla(netKarRow, -6) + ceyrekDegeriHesapla(netKarRow, -5) + ceyrekDegeriHesapla(netKarRow, -4)

    anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", 0) / getBilancoDegeri("Net Dönem Karı veya Zararı", 0)
    sermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)

    # print("Yıllık Net Kar: ", yillikNetKar)
    # print("Önceki Yıl Net Kar: ", oncekiYilNetKarToplami)

    netKarBuyumeOraniYillik = (yillikNetKar/oncekiYilNetKarToplami-1)
    print ("Yıllık Net Kar Büyüme: ", "{:.2%}".format(netKarBuyumeOraniYillik))
    print("Son Dört Çeyrek Net Kar Toplamı: ", "{:,.0f}".format(yillikNetKar).replace(",", "."))
    print("Önceki Yıl Net Kar Toplamı: ", "{:,.0f}".format(oncekiYilNetKarToplami).replace(",", "."))
    oncekiYilAyniCeyregeGoreNetKarBuyume = (ceyrekDegeriHesapla(netKarRow, 0)/ceyrekDegeriHesapla(netKarRow, -4) - 1)
    print("Önceki Yıl Aynı Çeyreğe Göre Net Kar Büyüme: ", "{:.2%}".format(oncekiYilAyniCeyregeGoreNetKarBuyume))
    print("Son Çeyrek Net Kar: ", "{:,.0f}".format(ceyrekDegeriHesapla(netKarRow, 0)).replace(",", "."))
    print("Önceki Yıl Aynı Çeyrek Net Kar: ", "{:,.0f}".format(ceyrekDegeriHesapla(netKarRow, -4)).replace(",", "."))


    # TEMEL CARPANLAR
    print("")
    print ("---------- TEMEL ORANLAR ----------")

    # F/K
    fkOrani = hisseFiyati / ((yillikNetKar * anaOrtaklikPayi) / (sermaye))
    print ("F/K Orani: ", "{:,.2f}".format(fkOrani))

    # PD/DD
    piyasaDegeri = sermaye * hisseFiyati;
    print("Piyasa Değeri (PD): ", "{:,.0f}".format(piyasaDegeri).replace(",", "."))
    nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
    print ("Nakit ve Nakit Benzerleri: ", "{:,.0f}".format(nakitVeNakitBenzerleri).replace(",", "."))
    nakitPd = nakitVeNakitBenzerleri/piyasaDegeri
    print ("Nakit / PD: ", "{:,.2f}".format(nakitPd))

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
    nakitFd = nakitVeNakitBenzerleri / firmaDegeri
    print("Nakit / FD: ", "{:,.2f}".format(nakitFd))


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



    yillikEsasFaaliyetKari = yillikBrutKar + yillikGenelYonetimGiderleri + yillikPazarlamaGiderleri + yillikArgeGiderleri

    favok = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
    print ("FAVÖK: ", "{:,.0f}".format(favok).replace(",","."))


    # FD/FAVOK Hesabi
    fdfavok = firmaDegeri/favok
    print("FD/FAVÖK: ", "{:,.2f}".format(fdfavok))

    # EFK Hesabi
    print ("Yıllık EFK: ", "{:,.0f}".format(yillikEsasFaaliyetKari).replace(",","."))

    #PD/EFK Hesabi
    pdefk = piyasaDegeri / yillikEsasFaaliyetKari
    print ("PD/EFK: ""{:,.2f}".format(pdefk))

    #HBK Hesabi
    hbk = yillikNetKar / (sermaye)
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
    roe = yillikNetKar/ortDefterDegeri
    # print("Yıllık Net Kar: ", "{:,.0f}".format(yillikNetKar).replace(",", "."))
    # print("Özsermaye: ", "{:,.0f}".format(defterDegeri).replace(",", "."))
    # print("Yıllık Ortalama Özsermaye: ", "{:,.0f}".format(ortDefterDegeri).replace(",", "."))
    print("ROE (Özsermaye Karlılığı - Özkaynak Getirisi): ", "{:.2%}".format(roe))

    # Aktif Karlılık Hesabı
    bilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", 0)
    dortOncekiBilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", -4)
    toplamVarliklar = (bilancoDonemiToplamVarliklar + dortOncekiBilancoDonemiToplamVarliklar) / 2
    aktifKarlilik = yillikNetKar / toplamVarliklar
    print("ROA (Aktif Karlılık): ", "{:.2%}".format(aktifKarlilik))

    # Kar Marjı Hesabı
    netKarMarji = yillikNetKar / yillikHasilat
    sonCeyrekNetKarMarji = ceyrekDegeriHesapla(netKarRow, 0) / ceyrekDegeriHesapla(hasilatRow, -0)
    print ("Yıllık Net Kar Marjı: ", "{:.2%}".format(netKarMarji))
    print("Son Çeyrek Net Kar Marjı: ", "{:.2%}".format(sonCeyrekNetKarMarji))

    # Aktif Devir Hızı Hesabı
    aktifDevirHizi = yillikHasilat / toplamVarliklar
    print ("Aktif Devir Hızı: ", "{:.2}".format(aktifDevirHizi))

    # Borç Kaynak Oranı Hesabı
    borcKaynakOrani = getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", 0) / getBilancoDegeri("TOPLAM KAYNAKLAR", 0)
    print("Borç/Kaynak Oranı: ", "{:.2%}".format(borcKaynakOrani))


    # Halka Açıklık Oranını Getir
    halkaAciklikOrani = returnHisseHalkaAciklikOrani(varHisseAdi)
    print("Halka Açıklık Oranı: ", "{:.2%}".format(halkaAciklikOrani))

    # Piyasa Değerini Getir
    print ("Piyasa Değeri: " + "{:,.0f}".format(piyasaDegeri/1000000).replace(",", ".") + " Milyon TL")

    # Sermaye Getir
    print ("Sermaye: " + "{:,.0f}".format(sermaye/1000000).replace(",", ".") + " Milyon TL")

    # Bedelsiz Sermaye Artırım Potansiyeli Hesapla

    def sermayeArtirimPot(varHisseAdi, varBilancoDonemi):
        odenmisSermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)
        ozkaynaklar = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
        netKarRow = getBilancoTitleRow("Net Dönem Karı veya Zararı")

        try:
            netKarYillik = yilliklandirmisDegerHesapla(netKarRow, 0)
            sermayeArtirimPotansiyeli = (ozkaynaklar - odenmisSermaye) / odenmisSermaye
            print("Sermaye Artirim Potansiyeli:" , "{:.0%}".format(sermayeArtirimPotansiyeli))
            return sermayeArtirimPotansiyeli
        except Exception as e:
            print(varHisseAdi, "\t", "HATA")
            return -1

    sermayeArtirimPotansiyeli = sermayeArtirimPot(hisseAdi, bilancoDonemi)



    # ROIC Hesabi

    print ("-------------ROIC HESABI----------")

    # uzunVadeBorcunKısaVadeliKisimlari = getBilancoDegeri("Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", 0)
    # kisaVadeliFinansalBorclar = getBilancoDegeri("Kısa Vadeli Borçlanmalar",0) + uzunVadeBorcunKısaVadeliKisimlari + getBilancoDegeri("Kısa Diğer Finansal Yükümlülükler1",0)
    # uzunVadeliBorclar = getBilancoDegeri("Uzun Vadeli Borçlanmalar",0)
    # uzunDigerFinansalYukumlulukler = getBilancoDegeri("Uzun Diğer Finansal Yükümlülükler",0)
    # kisaDigerFinansalYukumlulukler = getBilancoDegeri("Kısa Diğer Finansal Yükümlülükler",0)
    # toplamKisaVadeliFinansalBorclar = kisaVadeliFinansalBorclar + kisaDigerFinansalYukumlulukler
    # toplamUzunVadeliFinansalBorclar = uzunVadeliBorclar + uzunDigerFinansalYukumlulukler
    # yatirilmisSermaye = toplamOzkaynak + toplamKisaVadeliFinansalBorclar + toplamUzunVadeliFinansalBorclar
    # toplamOzkaynak = getBilancoDegeri("TOPLAM ÖZKAYNAKLAR", 0)

    yillikEsasFaaliyetKari = yilliklandirmisDegerHesapla(esasFaaliyetKariRow,0)
    yillikDonemVergi = yilliklandirmisDegerHesapla(donemVergiGideriRow,0)
    yillikErtelenmisVergiGeliri = yilliklandirmisDegerHesapla(ertelenmisVergiGideriRow, 0)
    nopat = yillikEsasFaaliyetKari - yillikDonemVergi + yillikErtelenmisVergiGeliri

    kontrolGucuOlmayanPaylar = getBilancoDegeri("Kontrol Gücü Olmayan Paylar",0)
    anaOrtakligaAitOzkaynaklar = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
    ozkaynaklar = anaOrtakligaAitOzkaynaklar - nakitVeNakitBenzerleri - finansalYatirimlar
    kvFinansalBorclar = getBilancoDegeri("Kısa Vadeli Borçlanmalar",0)
    uvFinansalBorclar = getBilancoDegeri("Uzun Vadeli Borçlanmalar",0)
    ertelenmisGelirler = getBilancoDegeri("Ertelenmiş Gelirler",0)
    uzunVadeliKarsiliklar = getBilancoDegeri("Uzun Vadeli Karşılıklar",0)

    finansmanYaklasimi = ozkaynaklar + kvFinansalBorclar + uvFinansalBorclar + ertelenmisGelirler + uzunVadeliKarsiliklar + kontrolGucuOlmayanPaylar

    roic = nopat / finansmanYaklasimi
    print("ROIC: ", "{:.2%}".format(roic))



    #Excel'e Rasyolari Yazdir
    #######################################################################################
    def createTopRow():
        bookSheetWrite.write(0, 0, "Hisse Adı")
        bookSheetWrite.write(0, 1, "Tarih")
        bookSheetWrite.write(0, 2, "Hisse Fiyatı")
        bookSheetWrite.write(0, 3, "Net Kar Büyüme Yıllık")
        bookSheetWrite.write(0, 4, "Net Kar Büyüme 4 Önceki Çeyreğe Göre")
        bookSheetWrite.write(0, 5, "F/K")
        bookSheetWrite.write(0, 6, "Nakit/PD")
        bookSheetWrite.write(0, 7, "Nakit/FD")
        bookSheetWrite.write(0, 8, "PD/DD")
        bookSheetWrite.write(0, 9, "PEG")
        bookSheetWrite.write(0, 10, "FD/Satışlar")
        bookSheetWrite.write(0, 11, "FD/FAVÖK")
        bookSheetWrite.write(0, 12, "PD/EFK")
        bookSheetWrite.write(0, 13, "Cari Oran")
        bookSheetWrite.write(0, 14, "ROE (Özsermaye Karlılığı)")
        bookSheetWrite.write(0, 15, "ROA (Aktif Karlılık)")
        bookSheetWrite.write(0, 16, "Yıllık Net Kar Marjı")
        bookSheetWrite.write(0, 17, "Son Çeyrek Net Kar Marjı")
        bookSheetWrite.write(0, 18, "Aktif Devir Hızı")
        bookSheetWrite.write(0, 19, "Borç/Kaynak")
        bookSheetWrite.write(0, 20, "Halka Açıklık Oranı")
        bookSheetWrite.write(0, 21, "Piyasa Değeri Milyon TL")
        bookSheetWrite.write(0, 22, "Sermaye Milyon TL")
        bookSheetWrite.write(0, 23, "Sermaye Artırım Potansiyeli")

    def reportResults(rowNumber):
        bookSheetWrite.write(rowNumber, 0, varHisseAdi)
        bookSheetWrite.write(rowNumber, 1, datetime.today().strftime('%d.%m.%Y'))
        bookSheetWrite.write(rowNumber, 2, hisseFiyati)
        bookSheetWrite.write(rowNumber, 3, netKarBuyumeOraniYillik)
        bookSheetWrite.write(rowNumber, 4, oncekiYilAyniCeyregeGoreNetKarBuyume)
        bookSheetWrite.write(rowNumber, 5, fkOrani)
        bookSheetWrite.write(rowNumber, 6, nakitPd)
        bookSheetWrite.write(rowNumber, 7, nakitFd)
        bookSheetWrite.write(rowNumber, 8, pddd)
        bookSheetWrite.write(rowNumber, 9, pegOrani)
        bookSheetWrite.write(rowNumber, 10, fdSatislar)
        bookSheetWrite.write(rowNumber, 11, fdfavok)
        bookSheetWrite.write(rowNumber, 12, pdefk)
        bookSheetWrite.write(rowNumber, 13, cariOran)
        bookSheetWrite.write(rowNumber, 14, roe)
        bookSheetWrite.write(rowNumber, 15, aktifKarlilik)
        bookSheetWrite.write(rowNumber, 16, netKarMarji)
        bookSheetWrite.write(rowNumber, 17, sonCeyrekNetKarMarji)
        bookSheetWrite.write(rowNumber, 18, aktifDevirHizi)
        bookSheetWrite.write(rowNumber, 19, borcKaynakOrani)
        bookSheetWrite.write(rowNumber, 20, halkaAciklikOrani)
        bookSheetWrite.write(rowNumber, 21, (int) (piyasaDegeri/1000000))
        bookSheetWrite.write(rowNumber, 22, (int) (sermaye / 1000000))
        bookSheetWrite.write(rowNumber, 23, sermayeArtirimPotansiyeli)


    if os.path.isfile(varReportFile):
        print("Rapor dosyası var, güncellenecek:", varReportFile)
        bookRead = xlrd.open_workbook(varReportFile, formatting_info=True)
        sheetRead = bookRead.sheet_by_index(0)
        rowNumber = sheetRead.nrows
        bookWrite = copy(bookRead)
        bookSheetWrite = bookWrite.get_sheet(0)
        reportResults(rowNumber)
        bookWrite.save(varReportFile)

    else:
        print("Rapor dosyası yeni oluşturulacak: ", varReportFile)
        bookWrite = xlwt.Workbook()
        bookSheetWrite = bookWrite.add_sheet(str(bilancoDonemi))
        createTopRow()
        reportResults(1)
        bookWrite.save(varReportFile)



    

hesapla(hisseAdi, bilancoDonemi)

