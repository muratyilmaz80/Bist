import math

from GetGuncelHisseDegeri import returnGuncelHisseDegeri
from GetHisseHalkaAciklikOrani import returnHisseHalkaAciklikOrani
import os
import xlrd
import xlwt
from xlutils.copy import copy
import os.path
from datetime import datetime
import pandas as pd


hisseAdi = "DEVA"
print("Hisse Adı: ", hisseAdi)
bilancoDonemi = 202306
print (f"Bilanço Dönemi: {bilancoDonemi}")
directory = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar"
varReportFile = "//Users//myilmaz//Documents//bist//Report_202306_Rasyolar.xls"

bilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + hisseAdi + ".xlsx"
bd_df = pd.read_excel(bilancoDosyasi, index_col=0)

hesaplamaTarihi = datetime.today().strftime('%d.%m.%Y')

hisseFiyati = returnGuncelHisseDegeri(hisseAdi)
print("Güncel Hisse Fiyatı: ", hisseFiyati)




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
    quarter = donem % 100
    birOncekiBilancoDonemi = birOncekiBilancoDoneminiHesapla(donem)
    if (quarter == 3):
        return bd_df.loc[r][donem]
    else:
        return (bd_df.loc[r][donem] - bd_df.loc[r][birOncekiBilancoDonemi])


def getBilancoDegeri(label, col):
    donem = bilancoDoneminiBul(col)
    try:
        bilancoDegeri = bd_df.loc[label][donem]
        if math.isnan(bilancoDegeri):
            return 0
        else:
            return bilancoDegeri
    except:
        print (f"Bilançoda ilgili alan bulunamadı! Label: {label} Çeyrek: {donem}")
        return -1;

def yilliklandirmisDegerHesapla (row, bd):
    toplam = ceyrekDegeriHesapla(row, bd) + ceyrekDegeriHesapla(row, bd-1) + ceyrekDegeriHesapla(row, bd-2) + ceyrekDegeriHesapla(row, bd-3)
    return toplam



def rasyolariHesapla():

    rasyolariHesapla.netKarBuyumeOraniYillik = -1
    rasyolariHesapla.oncekiYilAyniCeyregeGoreNetKarBuyume = -1
    rasyolariHesapla.yillikEsasFaaliyetKariBuyumeOrani = -1
    rasyolariHesapla.yillikHasilatBuyumeOrani = -1
    rasyolariHesapla.fkOrani = -1
    rasyolariHesapla.nakitPd = -1
    rasyolariHesapla.nakitFd = -1
    rasyolariHesapla.pddd = -1
    rasyolariHesapla.pegOrani = -1
    rasyolariHesapla.fdSatislar = -1
    rasyolariHesapla.favok = -1 #Hesaplarda kullanmak icin
    rasyolariHesapla.favokOncekiYil = -1  # Hesaplarda kullanmak icin
    rasyolariHesapla.yillikFavokArtisOrani = -1
    rasyolariHesapla.yillikEsasFaaliyetKari = -1 #Hesaplarda kullanmak icin
    rasyolariHesapla.yillikHasilat = -1 #Hesaplarda kullanmak icin
    rasyolariHesapla.yillikNetKar = -1 #Hesaplarda kullanmak icin
    rasyolariHesapla.oncekiYilNetKar = -1 #Hesaplarda kullanmak icin
    rasyolariHesapla.netBorc = -1 #Hesaplarda kullanmak icin
    rasyolariHesapla.fdfavok = -1
    rasyolariHesapla.pdefk = -1
    rasyolariHesapla.cariOran = -1
    rasyolariHesapla.likitOrani = -1
    rasyolariHesapla.nakitOrani = -1
    rasyolariHesapla.asitTestOrani = -1
    rasyolariHesapla.roe = -1
    rasyolariHesapla.aktifKarlilik = -1
    rasyolariHesapla.yillikNetKarMarji = -1
    rasyolariHesapla.sonCeyrekNetKarMarji = -1
    rasyolariHesapla.aktifDevirHizi = -1
    rasyolariHesapla.borcKaynakOrani = -1
    rasyolariHesapla.halkaAciklikOrani = -1
    rasyolariHesapla.piyasaDegeri = -1
    rasyolariHesapla.sermaye = -1
    rasyolariHesapla.sermayeArtirimPotansiyeli = -1
    rasyolariHesapla.yillikOzsermayeBuyumesi = -1

    # Ortak Hesaplamalar
    rasyolariHesapla.yillikHasilat = yilliklandirmisDegerHesapla("Hasılat", 0)
    rasyolariHesapla.yillikEsasFaaliyetKari = yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", 0)
    rasyolariHesapla.yillikNetKar = yilliklandirmisDegerHesapla("Net Dönem Karı veya Zararı", 0)
    rasyolariHesapla.oncekiYilNetKar = yilliklandirmisDegerHesapla("Net Dönem Karı veya Zararı", -4)
    rasyolariHesapla.sermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)


    # FİNANSAL ORANLARIN HESABI
    print("")
    print("FİNANSAL ORANLAR:")



    def netKarBuyumeOraniYillikHesapla():
        rasyolariHesapla.netKarBuyumeOraniYillik = (rasyolariHesapla.yillikNetKar / rasyolariHesapla.oncekiYilNetKar - 1)
        ynkPrint = "{:,.0f}".format(rasyolariHesapla.yillikNetKar).replace(",", ".")
        oynkPrint = "{:,.0f}".format(rasyolariHesapla.oncekiYilNetKar).replace(",", ".")
        nkboyPrint = "{:.2%}".format(rasyolariHesapla.netKarBuyumeOraniYillik)
        print(f"Yıllık Net Kar Büyüme: {nkboyPrint} ({ynkPrint}/{oynkPrint})")


    def oncekiYilAyniCeyregeGoreNetKarBuyumeOraniHesapla():
        rasyolariHesapla.oncekiYilAyniCeyregeGoreNetKarBuyume = (ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", 0) / ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -4) - 1)
        scnkPrint = "{:,.0f}".format(ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", 0)).replace(",", ".")
        oyacnkPrint = "{:,.0f}".format(ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -4)).replace(",", ".")
        oncekiYilAyniCeyregeGoreNetKarBuyumePrint = "{:.2%}".format(rasyolariHesapla.oncekiYilAyniCeyregeGoreNetKarBuyume)
        print(f"Önceki Yıl Aynı Çeyreğe Göre Net Kar Büyüme: {oncekiYilAyniCeyregeGoreNetKarBuyumePrint} ({scnkPrint}/{oyacnkPrint})" )

    def esasFaaliyetKariBuyumeOraniHesapla():
        yillikEfk = yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", 0)
        oncekiYilEfk = yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", -4)
        rasyolariHesapla.yillikEsasFaaliyetKariBuyumeOrani = (yillikEfk / oncekiYilEfk -1)
        print (f"Yıllık Esas Faaliyet Karı Artış Oranı: {rasyolariHesapla.yillikEsasFaaliyetKariBuyumeOrani}")

    def hasilatBuyumeOraniHesapla():
        yillikHasilat = yilliklandirmisDegerHesapla("Hasılat", 0)
        oncekiYilHasilat = yilliklandirmisDegerHesapla("Hasılat", -4)
        rasyolariHesapla.yillikHasilatBuyumeOrani = (yillikHasilat / oncekiYilHasilat -1)
        print (f"Yıllık Hasılat Artış Oranı: {rasyolariHesapla.yillikHasilatBuyumeOrani}")


    def fkOraniHesapla():
        anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", 0) / getBilancoDegeri("Net Dönem Karı veya Zararı",0)
        rasyolariHesapla.fkOrani = hisseFiyati / ((rasyolariHesapla.yillikNetKar * anaOrtaklikPayi) / (rasyolariHesapla.sermaye))
        print("F/K Orani: ", "{:,.2f}".format(rasyolariHesapla.fkOrani))


    def piyasaDegeriHesapla():
        sermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)
        rasyolariHesapla.piyasaDegeri = sermaye * hisseFiyati;
        print("Piyasa Değeri (PD): ", "{:,.0f}".format(rasyolariHesapla.piyasaDegeri).replace(",", "."))


    def pdDdOraniHesapla():
        defterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
        rasyolariHesapla.pddd = rasyolariHesapla.piyasaDegeri / defterDegeri
        print("PD/DD: ", "{:,.2f}".format(rasyolariHesapla.pddd))


    def nakitPdOraniHespala():
        nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
        rasyolariHesapla.nakitPd = nakitVeNakitBenzerleri / rasyolariHesapla.piyasaDegeri
        print("Nakit / PD: ", "{:,.2f}".format(rasyolariHesapla.nakitPd))


    def pegOraniHesapla():
        rasyolariHesapla.pegOrani = rasyolariHesapla.fkOrani / (rasyolariHesapla.netKarBuyumeOraniYillik * 100)
        print("PEG Orani: ", "{:,.4f}".format(rasyolariHesapla.pegOrani))

    def netBorcHesapla():
        kisaVadeliFinansalBorclar = getBilancoDegeri("Kısa Vadeli Borçlanmalar", 0) + getBilancoDegeri(
            "Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", 0)
        uzunVadeliFinansalBorclar = getBilancoDegeri("Uzun Vadeli Borçlanmalar", 0)
        finansalBorclar = kisaVadeliFinansalBorclar + uzunVadeliFinansalBorclar
        nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
        finansalYatirimlar = getBilancoDegeri("Duran Finansal Yatırımlar", 0) + getBilancoDegeri(
            "Dönen Finansal Yatırımlar", 0)
        rasyolariHesapla.netBorc = finansalBorclar - nakitVeNakitBenzerleri - finansalYatirimlar
        print("Net Borç: ", "{:,.0f}".format(rasyolariHesapla.netBorc).replace(",", "."))


    def firmaDegeriHesapla():
        rasyolariHesapla.firmaDegeri = rasyolariHesapla.piyasaDegeri + rasyolariHesapla.netBorc
        print("Firma Değeri (FD): ", "{:,.0f}".format(rasyolariHesapla.firmaDegeri).replace(",", "."))


    def nakitFdOraniHesapla():
        nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
        rasyolariHesapla.nakitFd = nakitVeNakitBenzerleri / rasyolariHesapla.firmaDegeri
        print("Nakit / FD: ", "{:,.2f}".format(rasyolariHesapla.nakitFd))

    def fdSatislarOraniHesapla():
        rasyolariHesapla.fdSatislar = rasyolariHesapla.firmaDegeri / rasyolariHesapla.yillikHasilat
        print("FD/Satışlar: ", "{:,.2f}".format(rasyolariHesapla.fdSatislar))


    def favokHesabi():
        yillikBrutKar = yilliklandirmisDegerHesapla("BRÜT KAR (ZARAR)", 0)
        yillikGenelYonetimGiderleri = yilliklandirmisDegerHesapla("Genel Yönetim Giderleri", 0)

        try:
            yillikPazarlamaGiderleri = yilliklandirmisDegerHesapla("Pazarlama Giderleri", 0)
        except Exception as e:
            print("Bilançoda Pazarlama Giderleri Bulunmamaktadır!")
            yillikPazarlamaGiderleri = 0

        try:
            yillikArgeGiderleri = yilliklandirmisDegerHesapla("Araştırma ve Geliştirme Giderleri", 0)
        except Exception as e:
            print("Bilançoda AR-GE Giderleri Bulunmamaktadır!")
            yillikArgeGiderleri = 0

        try:
            yillikAmortisman = yilliklandirmisDegerHesapla("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler", 0)
        except Exception as e:
            print("Bilançoda Amortisman Gideri Bulunmamaktadır!")
            yillikAmortisman = 0

        rasyolariHesapla.favok = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
        print("FAVÖK: ", "{:,.0f}".format(rasyolariHesapla.favok).replace(",", "."))



    def oncekiYilFavokHesabi():
        yillikBrutKar = yilliklandirmisDegerHesapla("BRÜT KAR (ZARAR)", -4)
        yillikGenelYonetimGiderleri = yilliklandirmisDegerHesapla("Genel Yönetim Giderleri", -4)

        try:
            yillikPazarlamaGiderleri = yilliklandirmisDegerHesapla("Pazarlama Giderleri", -4)
        except Exception as e:
            print("Bilançoda Pazarlama Giderleri Bulunmamaktadır!")
            yillikPazarlamaGiderleri = 0

        try:
            yillikArgeGiderleri = yilliklandirmisDegerHesapla("Araştırma ve Geliştirme Giderleri", -4)
        except Exception as e:
            print("Bilançoda AR-GE Giderleri Bulunmamaktadır!")
            yillikArgeGiderleri = 0

        try:
            yillikAmortisman = yilliklandirmisDegerHesapla("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler", -4)
        except Exception as e:
            print("Bilançoda Amortisman Gideri Bulunmamaktadır!")
            yillikAmortisman = 0

        rasyolariHesapla.favokOncekiYil = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
        print("Önceki Yıl FAVÖK: ", "{:,.0f}".format(rasyolariHesapla.favokOncekiYil).replace(",", "."))


    def favokArtisOraniHesapla():
        rasyolariHesapla.yillikFavokArtisOrani = (rasyolariHesapla.favok / rasyolariHesapla.favokOncekiYil -1)
        print("Yıllık FAVÖK Artış Oranı: ", "{:.2%}".format(rasyolariHesapla.yillikFavokArtisOrani))

    def fdFavokOraniHesabi():
        rasyolariHesapla.fdfavok = rasyolariHesapla.firmaDegeri/rasyolariHesapla.favok
        print("FD/FAVÖK: ", "{:,.2f}".format(rasyolariHesapla.fdfavok))


    def pdEfkOraniHesapla():
        rasyolariHesapla.pdefk = rasyolariHesapla.piyasaDegeri / rasyolariHesapla.yillikEsasFaaliyetKari
        print("PD/EFK: ""{:,.2f}".format(rasyolariHesapla.pdefk))

    def hbkOraniHesapla():
        rasyolariHesapla.hbk = rasyolariHesapla.yillikNetKar / (rasyolariHesapla.sermaye)
        print("HBK:", "{:,.2f}".format(rasyolariHesapla.hbk))

    def roeHesabi():
        defterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
        dortOncekiCeyrekDefterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", -4)
        ortDefterDegeri = (defterDegeri + dortOncekiCeyrekDefterDegeri) / 2
        rasyolariHesapla.roe = rasyolariHesapla.yillikNetKar / ortDefterDegeri
        print("ROE (Özsermaye Karlılığı - Özkaynak Getirisi): ", "{:.2%}".format(rasyolariHesapla.roe))

    def aktifKarlilikHesapla():
        bilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", 0)
        dortOncekiBilancoDonemiToplamVarliklar = getBilancoDegeri("TOPLAM VARLIKLAR", -4)
        toplamVarliklar = (bilancoDonemiToplamVarliklar + dortOncekiBilancoDonemiToplamVarliklar) / 2
        rasyolariHesapla.aktifKarlilik = rasyolariHesapla.yillikNetKar / toplamVarliklar
        print("ROA (Aktif Karlılık): ", "{:.2%}".format(rasyolariHesapla.aktifKarlilik))

    def yillikNetKarMarjiHesapla():
        rasyolariHesapla.yillikNetKarMarji = rasyolariHesapla.yillikNetKar / rasyolariHesapla.yillikHasilat
        print("Yıllık Net Kar Marjı: ", "{:.2%}".format(rasyolariHesapla.yillikNetKarMarji))

    def sonCeyrekNetKarMarjiHesapla():
        rasyolariHesapla.sonCeyrekNetKarMarji = ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", 0) / ceyrekDegeriHesapla("Hasılat", -0)
        print("Son Çeyrek Net Kar Marjı: ", "{:.2%}".format(rasyolariHesapla.sonCeyrekNetKarMarji))

    def aktifDevirHiziHesapla():
        rasyolariHesapla.aktifDevirHizi = rasyolariHesapla.yillikHasilat / getBilancoDegeri("TOPLAM VARLIKLAR", 0)
        print ("Aktif Devir Hızı: ", "{:.2}".format(rasyolariHesapla.aktifDevirHizi))

    def borcKaynakOraniHesapla():
        rasyolariHesapla.borcKaynakOrani = getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", 0) / getBilancoDegeri("TOPLAM KAYNAKLAR", 0)
        print("Borç/Kaynak Oranı: ", "{:.2%}".format(rasyolariHesapla.borcKaynakOrani))


    def cariOranHesapla():
        donenVarliklar = getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", 0)
        kisaVadeliYukumlulukler = getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
        rasyolariHesapla.cariOran = donenVarliklar / kisaVadeliYukumlulukler
        print("Cari Oran: ", "{:.3}".format(rasyolariHesapla.cariOran))


    def likitOraniHesapla():
        donenVarliklar = getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", 0)
        stoklar = getBilancoDegeri("Stoklar", 0)
        digerDonenVarliklar = getBilancoDegeri("Diğer Dönen Varlıklar",0)
        kisaVadeliYukumlulukler = getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
        rasyolariHesapla.likitOrani = (donenVarliklar - stoklar - digerDonenVarliklar)/kisaVadeliYukumlulukler
        print("Likit Oranı: ", "{:.3}".format(rasyolariHesapla.likitOrani))


    def nakitOraniHesapla ():
        hazirDegerler = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
        kisaVadeliYukumlulukler = getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
        rasyolariHesapla.nakitOrani = hazirDegerler/kisaVadeliYukumlulukler
        print("Nakit Oranı: ", "{:.3}".format(rasyolariHesapla.nakitOrani))


    def asitTestOraniHesapla():
        donenVarliklar = getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", 0)
        stoklar = getBilancoDegeri("Stoklar", 0)
        kisaVadeliYukumlulukler = getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
        rasyolariHesapla.asitTestOrani = (donenVarliklar - stoklar)/kisaVadeliYukumlulukler
        print("Asit Test Oranı: ", "{:.3}".format(rasyolariHesapla.asitTestOrani))

    def halkaAciklikOraniniGetir():
        rasyolariHesapla.halkaAciklikOrani = returnHisseHalkaAciklikOrani(hisseAdi)
        print("Halka Açıklık Oranı: ", "{:.2%}".format(rasyolariHesapla.halkaAciklikOrani))


    def sermayeArtirimPotansiyeliniHesapla():
            odenmisSermaye = getBilancoDegeri("Ödenmiş Sermaye", 0)
            ozkaynaklar = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
            rasyolariHesapla.sermayeArtirimPotansiyeli = (ozkaynaklar - odenmisSermaye) / odenmisSermaye
            print("Sermaye Artirim Potansiyeli:" , "{:.0%}".format(rasyolariHesapla.sermayeArtirimPotansiyeli))

    # def roicHesapla():
    #
    #     uzunVadeBorcunKısaVadeliKisimlari = getBilancoDegeri("Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", 0)
    #     kisaVadeliBorclanmalar = getBilancoDegeri("Kısa Vadeli Borçlanmalar",0)
    #     digerKisaVadeliYukumlulukler = getBilancoDegeri("Diğer Kısa Vadeli Yükümlülükler",0)
    #     kisaVadeliFinansalBorclar = kisaVadeliBorclanmalar + uzunVadeBorcunKısaVadeliKisimlari + digerKisaVadeliYukumlulukler
    #
    #     uzunVadeliBorclar = getBilancoDegeri("Uzun Vadeli Borçlanmalar",0)
    #     digerUzunVadeliYukumlulukler = getBilancoDegeri("Diğer Uzun Vadeli Yükümlülükler",0)
    #     uzunVadeliFinansalBorclar = uzunVadeliBorclar + digerUzunVadeliYukumlulukler
    #
    #     toplamFinansalBorclar = uzunVadeliFinansalBorclar + kisaVadeliFinansalBorclar
    #
    #     toplamOzkaynaklar = getBilancoDegeri("TOPLAM ÖZKAYNAKLAR", 0)
    #     nakitVeNakitBenzerleri = getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
    #     ertelenmisGelirler = getBilancoDegeri("Kısa Ertelenmiş Gelirler", 0) + getBilancoDegeri("Uzun Ertelenmiş Gelirler", 0)
    #     uzunVadeliKarsiliklar = getBilancoDegeri("Uzun Vadeli Karşılıklar", 0)
    #     kontrolGucuOlmayanPaylar = getBilancoDegeri("Kontrol Gücü Olmayan Paylar", 0)
    #
    #     yatirilanSermaye = toplamOzkaynaklar + toplamFinansalBorclar - nakitVeNakitBenzerleri + ertelenmisGelirler + uzunVadeliKarsiliklar + kontrolGucuOlmayanPaylar
    #
    #     yilliklandirilmisVergi = yilliklandirmisDegerHesapla("Sürdürülen Faaliyetler Vergi (Gideri) Geliri",0)
    #     yillikEsasFaaliyetKari = yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)",0)
    #     ertelenmisVergi = yilliklandirmisDegerHesapla("Ertelenmiş Vergi (Gideri) Geliri", 0)
    #     isletmeKarliligi = yillikEsasFaaliyetKari + yilliklandirilmisVergi + ertelenmisVergi
    #
    #     roic = isletmeKarliligi / yatirilanSermaye
    #     print("ROIC: ", "{:.2%}".format(roic))


    def ozsermayeBuyumesiHesapla():
        defterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
        dortOncekiCeyrekDefterDegeri = getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", -4)
        rasyolariHesapla.yillikOzsermayeBuyumesi = defterDegeri / dortOncekiCeyrekDefterDegeri
        print("Yıllık Özsermaye Büyümesi: ", "{:.2%}".format(rasyolariHesapla.yillikOzsermayeBuyumesi))



    def excelExport():
        def createTopRow():
            bookSheetWrite.write(0, 0, "Hisse Adı")
            bookSheetWrite.write(0, 1, "Tarih")
            bookSheetWrite.write(0, 2, "Hisse Fiyatı")
            bookSheetWrite.write(0, 3, "Net Kar Büyüme Yıllık")
            bookSheetWrite.write(0, 4, "Net Kar Büyüme 4 Önceki Çeyreğe Göre")
            bookSheetWrite.write(0, 5, "Esas Faaliyet Karı Büyüme Yıllık")
            bookSheetWrite.write(0, 6, "Hasılat Büyüme Yıllık")
            bookSheetWrite.write(0, 7, "FAVÖK Büyüme Yıllık")
            bookSheetWrite.write(0, 8, "F/K")
            bookSheetWrite.write(0, 9, "Nakit/PD")
            bookSheetWrite.write(0, 10, "Nakit/FD")
            bookSheetWrite.write(0, 11, "PD/DD")
            bookSheetWrite.write(0, 12, "PEG")
            bookSheetWrite.write(0, 13, "FD/Satışlar")
            bookSheetWrite.write(0, 14, "FD/FAVÖK")
            bookSheetWrite.write(0, 15, "PD/EFK")
            bookSheetWrite.write(0, 16, "Cari Oran")
            bookSheetWrite.write(0, 17, "Likit Oranı")
            bookSheetWrite.write(0, 18, "Nakit Oranı")
            bookSheetWrite.write(0, 19, "Asit Test Oranı")
            bookSheetWrite.write(0, 20, "ROE (Özsermaye Karlılığı)")
            bookSheetWrite.write(0, 21, "ROA (Aktif Karlılık)")
            bookSheetWrite.write(0, 22, "Yıllık Net Kar Marjı")
            bookSheetWrite.write(0, 23, "Son Çeyrek Net Kar Marjı")
            bookSheetWrite.write(0, 24, "Aktif Devir Hızı")
            bookSheetWrite.write(0, 25, "Borç/Kaynak")
            bookSheetWrite.write(0, 26, "Özsermaye Büyümesi")
            bookSheetWrite.write(0, 27, "Halka Açıklık Oranı")
            bookSheetWrite.write(0, 28, "Piyasa Değeri Milyon TL")
            bookSheetWrite.write(0, 29, "Sermaye Milyon TL")
            bookSheetWrite.write(0, 30, "Sermaye Artırım Potansiyeli")


        def reportResults(rowNumber):
            bookSheetWrite.write(rowNumber, 0, hisseAdi)
            bookSheetWrite.write(rowNumber, 1, datetime.today().strftime('%d.%m.%Y'))
            bookSheetWrite.write(rowNumber, 2, hisseFiyati)
            bookSheetWrite.write(rowNumber, 3, rasyolariHesapla.netKarBuyumeOraniYillik)
            bookSheetWrite.write(rowNumber, 4, rasyolariHesapla.oncekiYilAyniCeyregeGoreNetKarBuyume)
            bookSheetWrite.write(rowNumber, 5, rasyolariHesapla.yillikEsasFaaliyetKariBuyumeOrani)
            bookSheetWrite.write(rowNumber, 6, rasyolariHesapla.yillikHasilatBuyumeOrani)
            bookSheetWrite.write(rowNumber, 7, rasyolariHesapla.yillikFavokArtisOrani)
            bookSheetWrite.write(rowNumber, 8, rasyolariHesapla.fkOrani)
            bookSheetWrite.write(rowNumber, 9, rasyolariHesapla.nakitPd)
            bookSheetWrite.write(rowNumber, 10, rasyolariHesapla.nakitFd)
            bookSheetWrite.write(rowNumber, 11, rasyolariHesapla.pddd)
            bookSheetWrite.write(rowNumber, 12, rasyolariHesapla.pegOrani)
            bookSheetWrite.write(rowNumber, 13, rasyolariHesapla.fdSatislar)
            bookSheetWrite.write(rowNumber, 14, rasyolariHesapla.fdfavok)
            bookSheetWrite.write(rowNumber, 15, rasyolariHesapla.pdefk)
            bookSheetWrite.write(rowNumber, 16, rasyolariHesapla.cariOran)
            bookSheetWrite.write(rowNumber, 17, rasyolariHesapla.likitOrani)
            bookSheetWrite.write(rowNumber, 18, rasyolariHesapla.nakitOrani)
            bookSheetWrite.write(rowNumber, 19, rasyolariHesapla.asitTestOrani)
            bookSheetWrite.write(rowNumber, 20, rasyolariHesapla.roe)
            bookSheetWrite.write(rowNumber, 21, rasyolariHesapla.aktifKarlilik)
            bookSheetWrite.write(rowNumber, 22, rasyolariHesapla.yillikNetKarMarji)
            bookSheetWrite.write(rowNumber, 23, rasyolariHesapla.sonCeyrekNetKarMarji)
            bookSheetWrite.write(rowNumber, 24, rasyolariHesapla.aktifDevirHizi)
            bookSheetWrite.write(rowNumber, 25, rasyolariHesapla.borcKaynakOrani)
            bookSheetWrite.write(rowNumber, 26, rasyolariHesapla.yillikOzsermayeBuyumesi)
            bookSheetWrite.write(rowNumber, 27, rasyolariHesapla.halkaAciklikOrani)
            bookSheetWrite.write(rowNumber, 28, (int)(rasyolariHesapla.piyasaDegeri / 1000000))
            bookSheetWrite.write(rowNumber, 29, (int)(rasyolariHesapla.sermaye / 1000000))
            bookSheetWrite.write(rowNumber, 30, rasyolariHesapla.sermayeArtirimPotansiyeli)

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



    netKarBuyumeOraniYillikHesapla()
    oncekiYilAyniCeyregeGoreNetKarBuyumeOraniHesapla()
    esasFaaliyetKariBuyumeOraniHesapla()
    hasilatBuyumeOraniHesapla()
    fkOraniHesapla()
    piyasaDegeriHesapla()
    nakitPdOraniHespala()
    pdDdOraniHesapla()
    pegOraniHesapla()
    netBorcHesapla()
    firmaDegeriHesapla()
    nakitFdOraniHesapla()
    fdSatislarOraniHesapla()
    favokHesabi()
    oncekiYilFavokHesabi()
    favokArtisOraniHesapla()
    fdFavokOraniHesabi()
    pdEfkOraniHesapla()
    hbkOraniHesapla()
    roeHesabi()
    aktifKarlilikHesapla()
    yillikNetKarMarjiHesapla()
    sonCeyrekNetKarMarjiHesapla()
    aktifDevirHiziHesapla()
    borcKaynakOraniHesapla()
    ozsermayeBuyumesiHesapla()
    cariOranHesapla()
    likitOraniHesapla()
    nakitOraniHesapla()
    asitTestOraniHesapla()
    halkaAciklikOraniniGetir()
    sermayeArtirimPotansiyeliniHesapla()
    excelExport()


rasyolariHesapla()
