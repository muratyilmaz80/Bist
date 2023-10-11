from datetime import datetime
from ExcelRowClass import ExcelRowClass
from GetHisseHalkaAciklikOrani import returnHisseHalkaAciklikOrani
from Rapor_Olustur import exportReportExcel
from prettytable import PrettyTable
import logging,sys,math
from BilancoOrtalamaDolarDegeri import ucAylikBilancoDonemiOrtalamaDolarDegeriBul
import pandas as pd
import xlrd, xlwt
from xlutils.copy import copy

class Algoritma():

    def __init__(self, bilancoDosyasi, bilancoDonemi, bondYield, hisseFiyati, reportFile, logPath, logLevel):
        hisseAdiTemp = bilancoDosyasi[64:]
        self.hisseAdi = hisseAdiTemp[:-5]
        self.bilancoDosyasi = bilancoDosyasi
        self.hisseFiyati = hisseFiyati
        self.bondYield = bondYield
        self.my_logger = logging.getLogger()
        self.my_logger.setLevel(logLevel)
        self.bilancoDonemi = bilancoDonemi
        self.reportFile = reportFile
        self.output_file_handler = logging.FileHandler(logPath + self.hisseAdi + ".txt")
        self.output_file_handler.level = logging.INFO
        self.stdout_handler = logging.StreamHandler(sys.stdout)
        self.my_logger.addHandler(self.output_file_handler)
        self.my_logger.addHandler(self.stdout_handler)
        self.bd_df = pd.read_excel(self.bilancoDosyasi, index_col=0)
        self.cok_kullanilan_degerleri_hesapla()
        self.my_logger.info ("-------------------------------- %s ------------------------", self.hisseAdi)

    def birOncekibilancoDoneminiHesapla(self, dnm):
        yil = int(dnm / 100)
        ceyrek = int(dnm % 100)
        if ceyrek == 3:
            return (yil - 1) * 100 + 12
        else:
            return yil * 100 + (ceyrek - 3)


    def bilancoDoneminiBul(self,i):
        if (i > 0):
            print("Hatalı Bilanço Dönemi!")
            return -999
        elif (i == 0):
            return self.bilancoDonemi
        else:
            a = self.bilancoDonemi
            while (i < 0):
                a = self.birOncekibilancoDoneminiHesapla(a)
                i = i + 1
            return a


    def getBilancoDegeri(self, label, col):
        donem = self.bilancoDoneminiBul(col)
        try:
            bilancoDegeri = self.bd_df.loc[label][donem]
            if math.isnan(bilancoDegeri):
                return 0
            else:
                return bilancoDegeri
        except:
            print(f"Bilançoda ilgili alan bulunamadı! Label: {label} Çeyrek: {donem}")
            return -1;


    def ceyrekDegeriHesapla(self, r, col):
        donem = self.bilancoDoneminiBul(col)
        quarter = donem % 100
        birOncekibilancoDonemi = self.birOncekibilancoDoneminiHesapla(donem)
        if (quarter == 3):
            return self.bd_df.loc[r][donem]
        else:
            return (self.bd_df.loc[r][donem] - self.bd_df.loc[r][birOncekibilancoDonemi])


    def yilliklandirmisDegerHesapla(self, row, bd):
        toplam = self.ceyrekDegeriHesapla(row, bd) + self.ceyrekDegeriHesapla(row, bd - 1) + self.ceyrekDegeriHesapla(row,bd - 2) + self.ceyrekDegeriHesapla(row, bd - 3)
        return toplam

    def onceki_yil_ayni_ceyrege_gore_degisimi_hesapla(self, row, donem):
        self.my_logger.debug("fonksiyon: onceki_yil_ayni_ceyrek_degisimi_hesapla")
        ceyrekDegeri = self.ceyrekDegeriHesapla(row, donem)
        self.my_logger.debug(f"Çeyrek Değeri: {ceyrekDegeri}")
        oncekiCeyrekDegeri = self.ceyrekDegeriHesapla(row, donem - 4)
        self.my_logger.debug(f"Önceki Çeyrek Değeri: {oncekiCeyrekDegeri}", )
        degisimSonucu = ceyrekDegeri / oncekiCeyrekDegeri - 1
        return degisimSonucu


    def cok_kullanilan_degerleri_hesapla(self):
        self.hasilat0 = self.ceyrekDegeriHesapla("Hasılat", 0)
        self.hasilat1 = self.ceyrekDegeriHesapla("Hasılat", -1)
        self.hasilat2 = self.ceyrekDegeriHesapla("Hasılat", -2)
        self.hasilat3 = self.ceyrekDegeriHesapla("Hasılat", -3)
        self.hasilat4 = self.ceyrekDegeriHesapla("Hasılat", -4)
        self.hasilat5 = self.ceyrekDegeriHesapla("Hasılat", -5)
        self.hasilat6 = self.ceyrekDegeriHesapla("Hasılat", -6)
        self.hasilat7 = self.ceyrekDegeriHesapla("Hasılat", -7)
        self.yillikHasilat = self.yilliklandirmisDegerHesapla("Hasılat", 0)
        self.oncekiYilHasilat = self.yilliklandirmisDegerHesapla("Hasılat", -4)

        self.faaliyetKari0 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", 0)
        self.faaliyetKari1 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -1)
        self.faaliyetKari2 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -2)
        self.faaliyetKari3 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -3)
        self.faaliyetKari4 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -4)
        self.faaliyetKari5 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -5)
        self.faaliyetKari6 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -6)
        self.faaliyetKari7 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -7)
        self.yillikFaaliyetKari = self.yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", 0)

        self.netKar0 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", 0)
        self.netKar1 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -1)
        self.netKar2 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -2)
        self.netKar3 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -3)
        self.netKar4 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -4)
        self.netKar5 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -5)
        self.netKar6 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -6)
        self.netKar7 = self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", -7)
        self.yillikNetKar = self.yilliklandirmisDegerHesapla("Net Dönem Karı veya Zararı", 0)
        self.oncekiYilNetKar = self.yilliklandirmisDegerHesapla("Net Dönem Karı veya Zararı", -4)

        self.brutKar0 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", 0)
        self.brutKar1 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", -1)
        self.brutKar2 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", -2)
        self.brutKar3 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", -3)
        self.brutKar4 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", -4)
        self.brutKar5 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", -5)
        self.brutKar6 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", -6)
        self.brutKar7 = self.ceyrekDegeriHesapla("BRÜT KAR (ZARAR)", -7)

        self.nakit = self.getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
        self.stoklar = self.getBilancoDegeri("Stoklar", 0)
        self.digerVarliklar = self.getBilancoDegeri("Diğer Dönen Varlıklar", 0)
        self.sermaye = self.getBilancoDegeri("Ödenmiş Sermaye", 0)
        self.defterDegeri = self.getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
        self.anaOrtaklikPayi = self.getBilancoDegeri("Ana Ortaklık Payları", 0) / self.getBilancoDegeri("DÖNEM KARI (ZARARI)", 0)
        self.borclar = int(self.getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", 0))

        self.alacaklar = self.getBilancoDegeri("Dönen Ticari Alacaklar", 0) + \
                    self.getBilancoDegeri("Dönen Diğer Alacaklar", 0) + \
                    self.getBilancoDegeri("Duran Ticari Alacaklar", 0) + \
                    self.getBilancoDegeri("Duran Diğer Alacaklar", 0)

        self.finansalVarliklar = self.getBilancoDegeri("Duran Finansal Yatırımlar", 0) + \
                            self.getBilancoDegeri("Dönen Finansal Yatırımlar", 0) + \
                            self.getBilancoDegeri("Özkaynak Yöntemiyle Değerlenen Yatırımlar", 0)

        self.maddiDuranVarliklar = self.getBilancoDegeri("Maddi Duran Varlıklar", 0)

        self.nakitVeNakitBenzerleri = self.getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)

        self.finansalYatirimlar = self.getBilancoDegeri("Duran Finansal Yatırımlar", 0) + self.getBilancoDegeri("Dönen Finansal Yatırımlar", 0)

        self.kisaVadeliFinansalBorclar = self.getBilancoDegeri("Kısa Vadeli Borçlanmalar", 0) + self.getBilancoDegeri("Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", 0)
        self.uzunVadeliFinansalBorclar = self.getBilancoDegeri("Uzun Vadeli Borçlanmalar", 0)
        self.finansalBorclar = self.kisaVadeliFinansalBorclar + self.uzunVadeliFinansalBorclar

        self.ortalamaDolarKuru0 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(0))
        self.ortalamaDolarKuru1 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(-1))
        self.ortalamaDolarKuru2 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(-2))
        self.ortalamaDolarKuru3 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(-3))
        self.ortalamaDolarKuru4 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(-4))
        self.ortalamaDolarKuru5 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(-5))
        self.ortalamaDolarKuru6 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(-6))
        self.ortalamaDolarKuru7 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(-7))




    def likidasyonDegeriHesapla(self):
        likidasyonDegeri = self.nakit + (self.alacaklar * 0.7) + (self.stoklar * 0.5) + (self.digerVarliklar * 0.7) + (self.finansalVarliklar * 0.7) + (self.maddiDuranVarliklar * 0.2)
        return likidasyonDegeri




    def runAlgoritma(self):

        self.my_logger.debug("Bilanco Donemi: %d", self.bilancoDonemi)



        def hasilat_hesaplari():

            # Bilanço Dönemi Satış(Hasılat) Gelirleri
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------HASILAT(SATIŞ) GELİRLERİ---------------------------")
            self.my_logger.info("")

            hasilat0Print = "{:,.0f}".format(self.hasilat0).replace(",", ".")
            hasilat4Print = "{:,.0f}".format(self.hasilat4).replace(",", ".")
            self.hasilat0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", 0)
            hasilat0DegisimiPrint = "{:.2%}".format(self.hasilat0Degisimi)

            hasilat1Print = "{:,.0f}".format(self.hasilat1).replace(",", ".")
            hasilat5Print = "{:,.0f}".format(self.hasilat5).replace(",", ".")
            self.hasilat1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", -1)
            hasilat1DegisimiPrint = "{:.2%}".format(self.hasilat1Degisimi)

            hasilat2Print = "{:,.0f}".format(self.hasilat2).replace(",", ".")
            hasilat6Print = "{:,.0f}".format(self.hasilat6).replace(",", ".")
            self.hasilat2Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", -2)
            hasilat2DegisimiPrint = "{:.2%}".format(self.hasilat2Degisimi)

            hasilat3Print = "{:,.0f}".format(self.hasilat3).replace(",", ".")
            hasilat7Print = "{:,.0f}".format(self.hasilat7).replace(",", ".")
            self.hasilat3Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", -3)
            hasilat3DegisimiPrint = "{:.2%}".format(self.hasilat3Degisimi)

            satisTablosu = PrettyTable()
            satisTablosu.field_names = ["ÇEYREK", "SATIŞ", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ", "YÜZDE DEĞİŞİM"]
            satisTablosu.align["SATIŞ"] = "r"
            satisTablosu.align["ÖNCEKİ YIL SATIŞ"] = "r"
            satisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            satisTablosu.add_row([self.bilancoDoneminiBul(0), hasilat0Print, self.bilancoDoneminiBul(-4), hasilat4Print, hasilat0DegisimiPrint])
            satisTablosu.add_row([self.bilancoDoneminiBul(-1), hasilat1Print, self.bilancoDoneminiBul(-5), hasilat5Print, hasilat1DegisimiPrint])
            satisTablosu.add_row([self.bilancoDoneminiBul(-2), hasilat2Print, self.bilancoDoneminiBul(-6), hasilat6Print, hasilat2DegisimiPrint])
            satisTablosu.add_row([self.bilancoDoneminiBul(-3), hasilat3Print, self.bilancoDoneminiBul(-7),hasilat7Print, hasilat3DegisimiPrint])
            self.my_logger.info(satisTablosu)

            # Bilanço Dönemi Satış Geliri Artış Kriteri
            self.bilancoDonemiHasilatDegisimiGecmeDurumu = (self.hasilat0Degisimi > 0.1)
            printText = "Bilanço Dönemi Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(self.hasilat0Degisimi) + " >? 10% " + " " + str(self.bilancoDonemiHasilatDegisimiGecmeDurumu)
            self.my_logger.info(printText)

            # Önceki Dönem Hasılat Geliri Artış Kriteri

            if (self.hasilat0Degisimi >= 1):
                self.my_logger.info ("Bilanço Dönemi Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak.")
                self.oncekiDonemHasilatDegisimiGecmeDurumu = True
                self.my_logger.info ("Önceki Dönem Satış Gelir Artışı Geçme Durumu: %s", self.oncekiDonemHasilatDegisimiGecmeDurumu)

            else:
                self.oncekiDonemHasilatDegisimiGecmeDurumu = (self.hasilat1Degisimi < self.hasilat0Degisimi)
                printText = "Önceki Dönem Satış Gelir Artışı Bilanço Döneminden Düşük Mü: " + "{:.2%}".format(self.hasilat1Degisimi) + " <? " + "{:.2%}".format(self.hasilat0Degisimi) + " " + str(self.oncekiDonemHasilatDegisimiGecmeDurumu)
                self.my_logger.info(printText)



        def faaliyet_kari_hesaplari():

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------------FAALİYET KARI---------------------------------")
            self.my_logger.info("")


            faaliyetKari0Print = "{:,.0f}".format(self.faaliyetKari0).replace(",", ".")
            faaliyetKari4Print = "{:,.0f}".format(self.faaliyetKari4).replace(",", ".")
            self.faaliyetKari0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", 0)
            faaliyetKari0DegisimiPrint = "{:.2%}".format(self.faaliyetKari0Degisimi)

            faaliyetKari1Print = "{:,.0f}".format(self.faaliyetKari1).replace(",", ".")
            faaliyetKari5Print = "{:,.0f}".format(self.faaliyetKari5).replace(",", ".")
            self.faaliyetKari1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", -1)
            faaliyetKari1DegisimiPrint = "{:.2%}".format(self.faaliyetKari1Degisimi)

            faaliyetKari2Print = "{:,.0f}".format(self.faaliyetKari2).replace(",", ".")
            faaliyetKari6Print = "{:,.0f}".format(self.faaliyetKari6).replace(",", ".")
            self.faaliyetKari2Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", -2)
            faaliyetKari2DegisimiPrint = "{:.2%}".format(self.faaliyetKari2Degisimi)

            faaliyetKari3Print = "{:,.0f}".format(self.faaliyetKari3).replace(",", ".")
            faaliyetKari7Print = "{:,.0f}".format(self.faaliyetKari7).replace(",", ".")
            self.faaliyetKari3Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", -3)
            faaliyetKari3DegisimiPrint = "{:.2%}".format(self.faaliyetKari3Degisimi)

            faaliyetKariTablosu = PrettyTable()
            faaliyetKariTablosu.field_names = ["ÇEYREK", "FAALİYET KARI", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI", "YÜZDE DEĞİŞİM"]
            faaliyetKariTablosu.align["FAALİYET KARI"] = "r"
            faaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI"] = "r"
            faaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(0), faaliyetKari0Print, self.bilancoDoneminiBul(-4), faaliyetKari4Print, faaliyetKari0DegisimiPrint])
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(-1), faaliyetKari1Print, self.bilancoDoneminiBul(-5), faaliyetKari5Print, faaliyetKari1DegisimiPrint])
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(-2), faaliyetKari2Print, self.bilancoDoneminiBul(-6), faaliyetKari6Print, faaliyetKari2DegisimiPrint])
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(-3), faaliyetKari3Print, self.bilancoDoneminiBul(-7),faaliyetKari7Print, faaliyetKari3DegisimiPrint])
            self.my_logger.info(faaliyetKariTablosu)


            # Bilanço Dönemi Faaliyet Kar Artış Kriteri
            if self.ceyrekDegeriHesapla("Net Dönem Karı veya Zararı", 0) < 0:
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = False
                self.my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Çeyrek Net Kar Negatif", str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu))

            elif self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", 0) < 0:
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = False
                self.my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Ceyrek Faaliyet Kari Negatif", str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu))

            elif ((self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", 0) > 0) and (self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -4)) < 0):
                faaliyetKari0Degisimi = 0
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = True
                self.my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Çeyrek Faaliyet Karı Negatiften Pozitife Geçmiş", str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu))

            else:
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = (self.faaliyetKari0Degisimi > 0.15)
                printText = "Bilanço Dönemi Faaliyet Karı Artışı:" + "{:.2%}".format(self.faaliyetKari0Degisimi) + " >? 15% " + str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)

            # Önceki Dönem Faaliyet Kar Artış Kriteri

            if self.faaliyetKari0Degisimi >= 1:
                self.birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu = True
                printText = "Önceki Dönem Faaliyet Kar Artışı: Bilanço Dönemi Faaliyet Karı Artışı 100%'ün Üzerinde, Karşılaştırma Yapılmayacak: " + "{:.2%}".format(faaliyetKari0Degisimi) + " " + str(self.birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)

            else:
                self.birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu = (self.faaliyetKari1Degisimi < self.faaliyetKari0Degisimi)
                printText = "Önceki Dönem Faaliyet Kar Artışı:" + "{:.2%}".format(self.faaliyetKari1Degisimi) + " < ? " + "{:.2%}".format(self.faaliyetKari0Degisimi) + str(self.birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)




        def net_kar_hesaplari():

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("-------------------NET KAR (DÖNEM KARI/ZARARI)--------------------------")
            self.my_logger.info("")



            netKar0Print = "{:,.0f}".format(self.netKar0).replace(",", ".")
            netKar4Print = "{:,.0f}".format(self.netKar4).replace(",", ".")
            netKar0DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Net Dönem Karı veya Zararı", 0))

            netKar1Print = "{:,.0f}".format(self.netKar1).replace(",", ".")
            netKar5Print = "{:,.0f}".format(self.netKar5).replace(",", ".")
            netKar1DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Net Dönem Karı veya Zararı", -1))

            netKar2Print = "{:,.0f}".format(self.netKar2).replace(",", ".")
            netKar6Print = "{:,.0f}".format(self.netKar6).replace(",", ".")
            netKar2DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Net Dönem Karı veya Zararı", -2))

            netKar3Print = "{:,.0f}".format(self.netKar3).replace(",", ".")
            netKar7Print = "{:,.0f}".format(self.netKar7).replace(",", ".")
            netKar3DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Net Dönem Karı veya Zararı", -3))

            netKarTablosu = PrettyTable()
            netKarTablosu.field_names = ["ÇEYREK", "NET KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL NET KAR", "YÜZDE DEĞİŞİM"]
            netKarTablosu.align["NET KAR"] = "r"
            netKarTablosu.align["ÖNCEKİ YIL NET KAR"] = "r"
            netKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            netKarTablosu.add_row([self.bilancoDoneminiBul(0), netKar0Print, self.bilancoDoneminiBul(-4), netKar4Print, netKar0DegisimiPrint])
            netKarTablosu.add_row([self.bilancoDoneminiBul(-1), netKar1Print, self.bilancoDoneminiBul(-5), netKar5Print, netKar1DegisimiPrint])
            netKarTablosu.add_row([self.bilancoDoneminiBul(-2), netKar2Print, self.bilancoDoneminiBul(-6),netKar6Print, netKar2DegisimiPrint])
            netKarTablosu.add_row([self.bilancoDoneminiBul(-3), netKar3Print, self.bilancoDoneminiBul(-7),netKar7Print, netKar3DegisimiPrint])
            self.my_logger.info(netKarTablosu)


        def brut_kar_hesaplari():
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("-------------------BRÜT KAR (BRÜT KAR/ZARAR)--------------------------")
            self.my_logger.info("")

            brutKar0Print = "{:,.0f}".format(self.brutKar0).replace(",", ".")
            brutKar4Print = "{:,.0f}".format(self.brutKar4).replace(",", ".")
            brutKar0DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", 0))

            brutKar1Print = "{:,.0f}".format(self.brutKar1).replace(",", ".")
            brutKar5Print = "{:,.0f}".format(self.brutKar5).replace(",", ".")
            brutKar1DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", -1))

            brutKar2Print = "{:,.0f}".format(self.brutKar2).replace(",", ".")
            brutKar6Print = "{:,.0f}".format(self.brutKar6).replace(",", ".")
            brutKar2DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", -2))

            brutKar3Print = "{:,.0f}".format(self.brutKar3).replace(",", ".")
            brutKar7Print = "{:,.0f}".format(self.brutKar7).replace(",", ".")
            brutKar3DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", -3))

            brutKarTablosu = PrettyTable()
            brutKarTablosu.field_names = ["ÇEYREK", "BRÜT KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL BRÜT KAR", "YÜZDE DEĞİŞİM"]
            brutKarTablosu.align["BRÜT KAR"] = "r"
            brutKarTablosu.align["ÖNCEKİ YIL BRÜT KAR"] = "r"
            brutKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            brutKarTablosu.add_row([self.bilancoDoneminiBul(0), brutKar0Print, self.bilancoDoneminiBul(-4), brutKar4Print, brutKar0DegisimiPrint])
            brutKarTablosu.add_row([self.bilancoDoneminiBul(-1), brutKar1Print, self.bilancoDoneminiBul(-5), brutKar5Print, brutKar1DegisimiPrint])
            brutKarTablosu.add_row([self.bilancoDoneminiBul(-2), brutKar2Print, self.bilancoDoneminiBul(-6), brutKar6Print, brutKar2DegisimiPrint])
            brutKarTablosu.add_row([self.bilancoDoneminiBul(-3), brutKar3Print, self.bilancoDoneminiBul(-7), brutKar7Print, brutKar3DegisimiPrint])
            self.my_logger.info(brutKarTablosu)



        def gercek_deger_hesabi():
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("----------------GERÇEK DEĞER HESABI--------------------------------------------")

            self.my_logger.info("Sermaye: %s TL", "{:,.0f}".format(self.sermaye).replace(",", "."))
            self.my_logger.info("Ana Ortaklık Payı: %s", "{:.3f}".format(self.anaOrtaklikPayi))

            hasilat0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", 0)
            hasilat1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", -1)
            self.my_logger.info("Son 4 Çeyrek Hasılat Toplamı: %s TL","{:,.0f}".format(self.yillikHasilat).replace(",", "."))
            self.onumuzdekiDortCeyrekHasilatTahmini = ((((hasilat0Degisimi + hasilat1Degisimi) / 2) + 1) * self.yillikHasilat)

            hasilatlarCeyrek = [self.hasilat3, self.hasilat2, self.hasilat1, self.hasilat0]
            maxHasilatCeyrek = max(hasilatlarCeyrek)

            self.my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini: %s TL","{:,.0f}".format(self.onumuzdekiDortCeyrekHasilatTahmini).replace(",", "."))

            if (self.onumuzdekiDortCeyrekHasilatTahmini > 4 * maxHasilatCeyrek):
                self.onumuzdekiDortCeyrekHasilatTahmini = 4 * maxHasilatCeyrek
                self.my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini 4*maxCeyrek olarak duzeltildi:")
                self.my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini: %s TL",
                               "{:,.0f}".format(self.onumuzdekiDortCeyrekHasilatTahmini).replace(",", "."))

            # HASILAT TAHMININI MANUEL DEGISTIRMEK ICIN
            # self.onumuzdekiDortCeyrekHasilatTahmini = 700000000000



            self.onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (self.faaliyetKari1 + self.faaliyetKari0) / (self.hasilat0 + self.hasilat1)
            self.my_logger.info("Önümüzdeki 4 Çeyrek Faaliyet Kar Marjı Tahmini: %s ","{:.2%}".format(self.onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))

            self.faaliyetKariTahmini1 = self.onumuzdekiDortCeyrekHasilatTahmini * self.onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
            self.my_logger.info("Faaliyet Kar Tahmini1: %s TL", "{:,.0f}".format(self.faaliyetKariTahmini1).replace(",", "."))

            self.faaliyetKariTahmini2 = ((self.faaliyetKari1 + self.faaliyetKari0) * 2 * 0.3) + (self.faaliyetKari0 * 4 * 0.5) + ((self.faaliyetKari3 + self.faaliyetKari2 + self.faaliyetKari1 + self.faaliyetKari0) * 0.2)
            self.my_logger.info("Faaliyet Kar Tahmini2: %s TL", "{:,.0f}".format(self.faaliyetKariTahmini2).replace(",", "."))

            self.ortalamaFaaliyetKariTahmini = (self.faaliyetKariTahmini1 + self.faaliyetKariTahmini2) / 2
            self.my_logger.info("Ortalama Faaliyet Kari Tahmini: %s TL","{:,.0f}".format(self.ortalamaFaaliyetKariTahmini).replace(",", "."))

            # print ("----MURAT-----")
            #
            # istiraklerdenGelenKarRow = getBilancoTitleRow("Özkaynak Yöntemiyle Değerlenen Yatırımların Karlarından (Zararlarından) Paylar")
            # istiraklerdenGelenNetKarSonCeyrek = ceyrekDegeriHesapla(istiraklerdenGelenKarRow,self.bilancoDonemiColumn)
            # print ("İştiraklerden Gelen Net Kar Son Çeyrek: ", "{:,.0f}".format(istiraklerdenGelenNetKarSonCeyrek).replace(",","."))
            #
            # istiraklerdenGelenNetKarYillik = ceyrekDegeriHesapla(istiraklerdenGelenKarRow,self.bilancoDonemiColumn) + ceyrekDegeriHesapla(istiraklerdenGelenKarRow,birOncekibilancoDonemiColumn) + ceyrekDegeriHesapla(istiraklerdenGelenKarRow,ikiOncekiself.bilancoDonemiColumn) + ceyrekDegeriHesapla(istiraklerdenGelenKarRow,ucOncekiself.bilancoDonemiColumn)
            # print ("İştiraklerden Gelen Net Kar Yıllık: ", "{:,.0f}".format(istiraklerdenGelenNetKarYillik).replace(",","."))
            #
            # print("----MURAT-----")

            self.hisseBasinaOrtalamaKarTahmini = ((self.ortalamaFaaliyetKariTahmini) * self.anaOrtaklikPayi) / self.sermaye
            self.my_logger.info("Hisse Başına Ortalama Kar Tahmini: %s TL", format(self.hisseBasinaOrtalamaKarTahmini, ".2f"))

            self.likidasyonDegeri = self.likidasyonDegeriHesapla()
            self.my_logger.info("Likidasyon Değeri: %s TL", "{:,.0f}".format(self.likidasyonDegeri).replace(",", "."))

            self.my_logger.info("Borçlar: %s TL", "{:,.0f}".format(self.borclar).replace(",", "."))

            self.bilancoEtkisi = (self.likidasyonDegeri - self.borclar) / self.sermaye * self.anaOrtaklikPayi
            self.my_logger.info("Bilanço Etkisi: %s TL", format(self.bilancoEtkisi, ".2f"))

            self.gercekDeger = (self.hisseBasinaOrtalamaKarTahmini * 7) + self.bilancoEtkisi
            self.my_logger.info("Gerçek Hisse Değeri: %s TL", format(self.gercekDeger, ".2f"))

            self.targetBuy = self.gercekDeger * 0.66
            self.my_logger.info("Target Buy: %s TL", format(self.targetBuy, ".2f"))

            self.my_logger.info("Bilanço Tarihindeki Hisse Fiyatı: %s TL", format(self.hisseFiyati, ".2f"))

            self.gercekFiyataUzaklik = self.hisseFiyati / self.targetBuy
            self.my_logger.info("Gerçek Fiyata Uzaklık Oranı: %s", "{:.2%}".format(self.gercekFiyataUzaklik))

            self.gercekFiyataUzaklikTl = self.hisseFiyati - self.targetBuy
            self.my_logger.info("Gerçek Fiyata Uzaklık %s TL:", format(self.gercekFiyataUzaklikTl, ".2f"))




        def netpro_kriteri_hesabi():
            # Netpro Hesapla
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("----------------NETPRO  ve FORWARD_PE KRİTERİ-----------------------------------------------------------------")

            sonDortDonemFaaliyetKariToplami = self.yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", 0)
            sonDortDonemNetKarToplami = self.yilliklandirmisDegerHesapla("Net Dönem Karı veya Zararı", 0)

            self.my_logger.info("Son 4 Dönem Net Kar Toplamı: %s TL", "{:,.0f}".format(sonDortDonemNetKarToplami).replace(",", "."))
            self.my_logger.info("Son 4 Dönem Faaliyet Karı Toplamı: %s TL", "{:,.0f}".format(sonDortDonemFaaliyetKariToplami).replace(",", "."))

            self.fkOrani = self.hisseFiyati / ((sonDortDonemNetKarToplami * self.anaOrtaklikPayi) / (self.sermaye))
            self.my_logger.info("F/K Oranı: %s", "{:,.2f}".format(self.fkOrani))

            self.hbkOrani = sonDortDonemNetKarToplami / (self.sermaye) * self.anaOrtaklikPayi
            self.my_logger.info("HBK Oranı: %s", "{:,.2f}".format(self.hbkOrani))

            hasilat0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", 0)
            hasilat1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", -1)
            onumuzdekiDortCeyrekHasilatTahmini = ((((hasilat0Degisimi + hasilat1Degisimi) / 2) + 1) * self.yillikHasilat)

            onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (self.faaliyetKari1 + self.faaliyetKari0) / (self.hasilat0 + self.hasilat1)
            self.my_logger.info("Önümüzdeki 4 Çeyrek Faaliyet Kar Marjı Tahmini: %s ","{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))
            faaliyetKariTahmini1 = onumuzdekiDortCeyrekHasilatTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
            faaliyetKariTahmini2 = ((self.faaliyetKari1 + self.faaliyetKari0) * 2 * 0.3) + (self.faaliyetKari0 * 4 * 0.5) + ((self.faaliyetKari3 + self.faaliyetKari2 + self.faaliyetKari1 + self.faaliyetKari0) * 0.2)
            ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1 + faaliyetKariTahmini2) / 2

            netProEstDegeri = ((ortalamaFaaliyetKariTahmini / sonDortDonemFaaliyetKariToplami) * sonDortDonemNetKarToplami) * self.anaOrtaklikPayi
            self.my_logger.info("NetPro Est Değeri: %s TL", "{:,.0f}".format(netProEstDegeri).replace(",", "."))

            likidasyonDegeri = self.likidasyonDegeriHesapla()
            borclar = int(self.getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", 0))
            bilancoEtkisi = (likidasyonDegeri - borclar) / self.sermaye * self.anaOrtaklikPayi
            piyasaDegeri = (bilancoEtkisi * self.sermaye * -1) + (self.hisseFiyati * self.sermaye)

            self.my_logger.info("Piyasa Değeri: %s TL", "{:,.0f}".format(piyasaDegeri).replace(",", "."))
            self.my_logger.info("self.bondYield: %s", "{:.2%}".format(self.bondYield))

            self.netProKriteri = (netProEstDegeri / piyasaDegeri) / self.bondYield
            self.netProKriteriGecmeDurumu = (self.netProKriteri > 2)
            self.my_logger.info("NetPro Kriteri (2'den Büyük Olmalı): %s %s", format(self.netProKriteri, ".2f"), str(self.netProKriteriGecmeDurumu))

            minNetProIcinhisseFiyati = (netProEstDegeri / (1.9 * self.bondYield) - (bilancoEtkisi * self.sermaye * -1)) / self.sermaye
            self.my_logger.info("NetPro 1.9 Olması İçin Olması Gereken Hisse Fiyatı: %s", format(minNetProIcinhisseFiyati, ".2f"))

            self.forwardPeKriteri = (piyasaDegeri) / netProEstDegeri
            self.forwardPeKriteriGecmeDurumu = (self.forwardPeKriteri < 4)
            printText = "Forward PE Kriteri (4'ten Küçük Olmalı): " + format(self.forwardPeKriteri, ".2f") + " " + str(self.forwardPeKriteriGecmeDurumu)
            self.my_logger.info(printText)



        def bilanco_donemi_dolar_hesabi():

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("----------------BİLANÇO DOLAR HESABI-------------------------------------")
            self.my_logger.info("")

            self.my_logger.info ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(0) , "{:,.2f}".format(self.ortalamaDolarKuru0))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(-1) , "{:,.2f}".format(self.ortalamaDolarKuru1))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(-2) ,"{:,.2f}".format(self.ortalamaDolarKuru2))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(-3) ,"{:,.2f}".format(self.ortalamaDolarKuru3))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(-4) ,"{:,.2f}".format(self.ortalamaDolarKuru4))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL", self.bilancoDoneminiBul(-5), "{:,.2f}".format(self.ortalamaDolarKuru5))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(-6), "{:,.2f}".format(self.ortalamaDolarKuru6))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL", self.bilancoDoneminiBul(-7), "{:,.2f}".format(self.ortalamaDolarKuru7))

        def dolar_hasilat_hesaplari():
            # Bilanço Dönemi Satış(Hasılat) Gelirleri (DOLAR)
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------HASILAT(SATIŞ) GELİRLERİ (DOLAR)----------------------")
            self.my_logger.info("")

            self.dolarHasilat0 = self.hasilat0 / self.ortalamaDolarKuru0
            self.dolarHasilat1 = self.hasilat1 / self.ortalamaDolarKuru1
            self.dolarHasilat2 = self.hasilat2 / self.ortalamaDolarKuru2
            self.dolarHasilat3 = self.hasilat3 / self.ortalamaDolarKuru3
            self.dolarHasilat4 = self.hasilat4 / self.ortalamaDolarKuru4
            self.dolarHasilat5 = self.hasilat5 / self.ortalamaDolarKuru5
            self.dolarHasilat6 = self.hasilat6 / self.ortalamaDolarKuru6
            self.dolarHasilat7 = self.hasilat7 / self.ortalamaDolarKuru7

            dolarHasilat0Print = "{:,.0f}".format(self.dolarHasilat0).replace(",", ".")
            dolarHasilat4Print = "{:,.0f}".format(self.dolarHasilat4).replace(",",".")
            self.dolarHasilat0Degisimi = self.dolarHasilat0 / self.dolarHasilat4 - 1
            dolarHasilatDegisimi0Print = "{:.2%}".format(self.dolarHasilat0Degisimi)

            dolarHasilat1Print = "{:,.0f}".format(self.dolarHasilat1).replace(",",".")
            dolarHasilat5Print = "{:,.0f}".format(self.dolarHasilat5).replace(",", ".")
            self.dolarHasilat1Degisimi = self.dolarHasilat1 / self.dolarHasilat5 - 1
            dolarHasilatDegisimi1Print = "{:.2%}".format(self.dolarHasilat1Degisimi)

            dolarHasilat2Print = "{:,.0f}".format(self.dolarHasilat2).replace(",",".")
            dolarHasilat6Print = "{:,.0f}".format(self.dolarHasilat6).replace(",", ".")
            dolarHasilatDegisimi2Print = "{:.2%}".format(self.dolarHasilat2 / self.dolarHasilat6 - 1)

            dolarHasilat3Print = "{:,.0f}".format(self.dolarHasilat3).replace(",",".")
            dolarHasilat7Print = "{:,.0f}".format(self.dolarHasilat7).replace(",", ".")
            dolarHasilatDegisimi3Print = "{:.2%}".format(self.dolarHasilat3 / self.dolarHasilat7 - 1)

            dolarSatisTablosu = PrettyTable()
            dolarSatisTablosu.field_names = ["ÇEYREK", "SATIŞ (USD)", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ (USD)", "YÜZDE DEĞİŞİM"]
            dolarSatisTablosu.align["SATIŞ (USD)"] = "r"
            dolarSatisTablosu.align["ÖNCEKİ YIL SATIŞ (USD)"] = "r"
            dolarSatisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            dolarSatisTablosu.add_row([self.bilancoDoneminiBul(0), dolarHasilat0Print, self.bilancoDoneminiBul(-4),dolarHasilat4Print, dolarHasilatDegisimi0Print])
            dolarSatisTablosu.add_row([self.bilancoDoneminiBul(-1), dolarHasilat1Print, self.bilancoDoneminiBul(-5), dolarHasilat5Print, dolarHasilatDegisimi1Print])
            dolarSatisTablosu.add_row([self.bilancoDoneminiBul(-2), dolarHasilat2Print, self.bilancoDoneminiBul(-6), dolarHasilat6Print, dolarHasilatDegisimi2Print])
            dolarSatisTablosu.add_row([self.bilancoDoneminiBul(-3), dolarHasilat3Print, self.bilancoDoneminiBul(-7),dolarHasilat7Print, dolarHasilatDegisimi3Print])
            self.my_logger.info(dolarSatisTablosu)


            # Bilanço Dönemi (DOLAR) Satış Geliri Artış Kriteri
            self.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (self.dolarHasilat0Degisimi > 0.1)

            printText = "Bilanço Dönemi (DOLAR) Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(self.dolarHasilat0Degisimi) + " >? 10%" + " " + str(self.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
            self.my_logger.info(printText)

            # Önceki Dönem (DOLAR) Hasılat Geliri Artış Kriteri
            #
            if (self.dolarHasilat0Degisimi >= 1):
                printText = "Bilanço Dönemi (DOLAR) Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak."
                self.my_logger.info(printText)
                self.oncekibilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = True
                printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Geçme Durumu: " + str(self.oncekibilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
                self.my_logger.info(printText)

            else:
                self.oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (self.dolarHasilat1Degisimi < self.dolarHasilat0Degisimi)
                printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Bilanço Döneminden Düşük Mü: " + "{:.2%}".format(
                    self.dolarHasilat1Degisimi) + " <? " + "{:.2%}".format(self.dolarHasilat0Degisimi) + " " + str(self.oncekiBilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
                self.my_logger.info(printText)







        def dolar_faaliyet_kari_hesaplari():

            # Bilanço Dönemi Faaliyet Karı Gelirleri (DOLAR)
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------------FAALİYET KARI (DOLAR)-------------------------")
            self.my_logger.info("")

            self.dolarFaaliyetKari0 = self.faaliyetKari0/self.ortalamaDolarKuru0
            self.dolarFaaliyetKari1 = self.faaliyetKari1/self.ortalamaDolarKuru1
            self.dolarFaaliyetKari2 = self.faaliyetKari2/self.ortalamaDolarKuru2
            self.dolarFaaliyetKari3 = self.faaliyetKari3/self.ortalamaDolarKuru3
            self.dolarFaaliyetKari4 = self.faaliyetKari4/self.ortalamaDolarKuru4
            self.dolarFaaliyetKari5 = self.faaliyetKari5/self.ortalamaDolarKuru5
            self.dolarFaaliyetKari6 = self.faaliyetKari6/self.ortalamaDolarKuru6
            self.dolarFaaliyetKari7 = self.faaliyetKari7/self.ortalamaDolarKuru7

            dolarFaaliyetKari0Print = "{:,.0f}".format(self.dolarFaaliyetKari0).replace(",", ".")
            dolarFaaliyetKari4Print = "{:,.0f}".format(self.dolarFaaliyetKari4).replace(",",".")
            self.dolarFaaliyetKari0Degisimi = self.dolarFaaliyetKari0/self.dolarFaaliyetKari4-1
            dolarFaaliyetKari0DegisimiPrint = "{:.2%}".format(self.dolarFaaliyetKari0Degisimi)

            dolarFaaliyetKari1Print = "{:,.0f}".format(self.dolarFaaliyetKari1).replace(",", ".")
            dolarFaaliyetKari5Print = "{:,.0f}".format(self.dolarFaaliyetKari5).replace(",", ".")
            self.dolarFaaliyetKari1Degisimi = self.dolarFaaliyetKari1/self.dolarFaaliyetKari5-1
            dolarFaaliyetKari1DegisimiPrint = "{:.2%}".format(self.dolarFaaliyetKari1Degisimi)

            dolarFaaliyetKari2Print = "{:,.0f}".format(self.dolarFaaliyetKari2).replace(",", ".")
            dolarFaaliyetKari6Print = "{:,.0f}".format(self.dolarFaaliyetKari6).replace(",", ".")
            dolarFaaliyetKari2DegisimiPrint = "{:.2%}".format(self.dolarFaaliyetKari2/self.dolarFaaliyetKari6-1)

            dolarFaaliyetKari3Print = "{:,.0f}".format(self.dolarFaaliyetKari3).replace(",", ".")
            dolarFaaliyetKari7Print = "{:,.0f}".format(self.dolarFaaliyetKari7).replace(",", ".")
            dolarFaaliyetKari3DegisimiPrint = "{:.2%}".format(self.dolarFaaliyetKari3/self.dolarFaaliyetKari7-1)

            dolarFaaliyetKariTablosu = PrettyTable()
            dolarFaaliyetKariTablosu.field_names = ["ÇEYREK", "FAALİYET KARI (DOLAR)", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI (DOLAR)", "YÜZDE DEĞİŞİM"]
            dolarFaaliyetKariTablosu.align["FAALİYET KARI (DOLAR)"] = "r"
            dolarFaaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI (DOLAR)"] = "r"
            dolarFaaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            dolarFaaliyetKariTablosu.add_row([self.bilancoDoneminiBul(0), dolarFaaliyetKari0Print, self.bilancoDoneminiBul(-4), dolarFaaliyetKari4Print, dolarFaaliyetKari0DegisimiPrint])
            dolarFaaliyetKariTablosu.add_row([self.bilancoDoneminiBul(-1), dolarFaaliyetKari1Print, self.bilancoDoneminiBul(-5), dolarFaaliyetKari5Print, dolarFaaliyetKari1DegisimiPrint])
            dolarFaaliyetKariTablosu.add_row([self.bilancoDoneminiBul(-2), dolarFaaliyetKari2Print, self.bilancoDoneminiBul(-6), dolarFaaliyetKari6Print, dolarFaaliyetKari2DegisimiPrint])
            dolarFaaliyetKariTablosu.add_row([self.bilancoDoneminiBul(-3), dolarFaaliyetKari3Print, self.bilancoDoneminiBul(-7), dolarFaaliyetKari7Print, dolarFaaliyetKari3DegisimiPrint])
            self.my_logger.info (dolarFaaliyetKariTablosu)

            # Bilanço Dönem Faaliyet Kar Artış Kriteri (DOLAR)
            if self.netKar0 < 0:
                self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu = False
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu) + " Son Çeyrek Net Kar Negatif"
                self.my_logger.info (printText)

            elif self.dolarFaaliyetKari0 < 0:
                self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu = False
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu) + " Son Ceyrek Dolar Faaliyet Kari Negatif"
                self.my_logger.info (printText)

            elif (self.dolarFaaliyetKari0 > 0) and (self.dolarFaaliyetKari4 < 0):
                self.bilancoDonemiDolarFaaliyetKariArtisi = 0
                self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = True
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu) + " Son Çeyrek Dolar Faaliyet Karı Negatiften Pozitife Geçmiş"
                self.my_logger.info (printText)

            else:
                self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu = (self.dolarFaaliyetKari0Degisimi > 0.15)
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + "{:.2%}".format(self.dolarFaaliyetKari0Degisimi) + " >? 15% " + str(self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)


            # Önceki Dönem Faaliyet Kar Artış Kriteri (DOLAR)

            if self.dolarFaaliyetKari0Degisimi >= 1:
                self.birOncekibilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu = True
                printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: Bilanço Dönemi Artış " + "{:.2%}".format(self.dolarFaaliyetKari0Degisimi) + " > 100%, Karşılaştırma Yapılmayacak: " + str(self.birOncekibilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)


            else:
                self.birOncekibilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu = (self.dolarFaaliyetKari1Degisimi < self.dolarFaaliyetKari0Degisimi)
                printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: " + "{:.2%}".format(self.dolarFaaliyetKari1Degisimi) + \
                            " <? " + "{:.2%}".format(self.dolarFaaliyetKari0Degisimi) + " " + str(self.birOncekibilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)







        # RASYO HESAPLARI
        #TODO: printleri loga çevir


        def netKarBuyumeOraniYillikHesapla():
            self.netKarBuyumeOraniYillik = (self.yillikNetKar / self.oncekiYilNetKar - 1)
            ynkPrint = "{:,.0f}".format(self.yillikNetKar).replace(",", ".")
            oynkPrint = "{:,.0f}".format(self.oncekiYilNetKar).replace(",", ".")
            nkboyPrint = "{:.2%}".format(self.netKarBuyumeOraniYillik)
            print(f"Yıllık Net Kar Büyüme: {nkboyPrint} ({ynkPrint}/{oynkPrint})")

        def oncekiYilAyniCeyregeGoreNetKarBuyumeOraniHesapla():
            self.oncekiYilAyniCeyregeGoreNetKarBuyume = self.netKar0 / self.netKar4 - 1
            scnkPrint = "{:,.0f}".format(self.netKar0).replace(",", ".")
            oyacnkPrint = "{:,.0f}".format(self.netKar4).replace(",", ".")
            oncekiYilAyniCeyregeGoreNetKarBuyumePrint = "{:.2%}".format(self.oncekiYilAyniCeyregeGoreNetKarBuyume)
            print(f"Önceki Yıl Aynı Çeyreğe Göre Net Kar Büyüme: {oncekiYilAyniCeyregeGoreNetKarBuyumePrint} ({scnkPrint}/{oyacnkPrint})")

        def esasFaaliyetKariBuyumeOraniHesapla():
            yillikEfk = self.yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", 0)
            oncekiYilEfk = self.yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", -4)
            self.yillikEsasFaaliyetKariBuyumeOrani = (yillikEfk / oncekiYilEfk - 1)
            print(f"Yıllık Esas Faaliyet Karı Artış Oranı: {self.yillikEsasFaaliyetKariBuyumeOrani}")

        def hasilatBuyumeOraniHesapla():
            self.yillikHasilatBuyumeOrani = (self.yillikHasilat / self.oncekiYilHasilat - 1)
            print(f"Yıllık Hasılat Artış Oranı: {self.yillikHasilatBuyumeOrani}")

        def fkOraniHesapla():
            self.fkOrani = self.hisseFiyati / ((self.yillikNetKar * self.anaOrtaklikPayi) / (self.sermaye))
            print("F/K Orani: ", "{:,.2f}".format(self.fkOrani))

        def piyasaDegeriHesapla():
            self.piyasaDegeri = self.sermaye * self.hisseFiyati;
            print("Piyasa Değeri (PD): ", "{:,.0f}".format(self.piyasaDegeri).replace(",", "."))

        def pdDdOraniHesapla():
            self.pdDd = self.piyasaDegeri / self.defterDegeri
            print("PD/DD: ", "{:,.2f}".format(self.pdDd))

        def nakitPdOraniHespala():
            self.nakitPd = self.nakitVeNakitBenzerleri / self.piyasaDegeri
            print("Nakit / PD: ", "{:,.2f}".format(self.nakitPd))

        def pegOraniHesapla():
            self.pegOrani = self.fkOrani / (self.netKarBuyumeOraniYillik * 100)
            print("PEG Orani: ", "{:,.4f}".format(self.pegOrani))

        def netBorcHesapla():
            self.netBorc = self.finansalBorclar - self.nakitVeNakitBenzerleri - self.finansalYatirimlar
            print("Net Borç: ", "{:,.0f}".format(self.netBorc).replace(",", "."))

        def firmaDegeriHesapla():
            self.firmaDegeri = self.piyasaDegeri + self.netBorc
            print("Firma Değeri (FD): ", "{:,.0f}".format(self.firmaDegeri).replace(",", "."))

        def nakitFdOraniHesapla():
            self.nakitFd = self.nakitVeNakitBenzerleri / self.firmaDegeri
            print("Nakit / FD: ", "{:,.2f}".format(self.nakitFd))

        def fdSatislarOraniHesapla():
            self.fdSatislar = self.firmaDegeri / self.yillikHasilat
            print("FD/Satışlar: ", "{:,.2f}".format(self.fdSatislar))

        def genelFavokHesabi(ceyrek):
            yillikBrutKar = self.yilliklandirmisDegerHesapla("BRÜT KAR (ZARAR)", ceyrek)
            yillikGenelYonetimGiderleri = self.yilliklandirmisDegerHesapla("Genel Yönetim Giderleri", ceyrek)

            try:
                yillikPazarlamaGiderleri = self.yilliklandirmisDegerHesapla("Pazarlama Giderleri", ceyrek)
            except Exception as e:
                print("Bilançoda Pazarlama Giderleri Bulunmamaktadır!")
                yillikPazarlamaGiderleri = 0

            try:
                yillikArgeGiderleri = self.yilliklandirmisDegerHesapla("Araştırma ve Geliştirme Giderleri", ceyrek)
            except Exception as e:
                print("Bilançoda AR-GE Giderleri Bulunmamaktadır!")
                yillikArgeGiderleri = 0

            try:
                yillikAmortisman = self.yilliklandirmisDegerHesapla("Amortisman ve İtfa Gideri İle İlgili Düzeltmeler", ceyrek)
            except Exception as e:
                print("Bilançoda Amortisman Gideri Bulunmamaktadır!")
                yillikAmortisman = 0

            favok = yillikBrutKar + yillikPazarlamaGiderleri + yillikGenelYonetimGiderleri + yillikArgeGiderleri + yillikAmortisman
            print(f"FAVÖK{ceyrek}: ", "{:,.0f}".format(favok).replace(",", "."))
            return favok

        def favokHesapla():
            self.favok = genelFavokHesabi(0)
            self.oncekiYilFavok = genelFavokHesabi(-4)

        def favokArtisOraniHesapla():
            self.yillikFavokArtisOrani = (self.favok / self.oncekiYilFavok - 1)
            print("Yıllık FAVÖK Artış Oranı: ", "{:.2%}".format(self.yillikFavokArtisOrani))

        def fdFavokOraniHesabi():
            self.fdFavok = self.firmaDegeri / self.favok
            print("FD/FAVÖK: ", "{:,.2f}".format(self.fdFavok))

        def pdEfkOraniHesapla():
            self.pdEfk = self.piyasaDegeri / self.yillikFaaliyetKari
            print("PD/EFK: ""{:,.2f}".format(self.pdEfk))

        def hbkOraniHesapla():
            self.hbk = self.yillikNetKar / (self.sermaye)
            print("HBK:", "{:,.2f}".format(self.hbk))

        def roeHesabi():
            dortOncekiCeyrekDefterDegeri = self.getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", -4)
            ortDefterDegeri = (self.defterDegeri + dortOncekiCeyrekDefterDegeri) / 2
            self.roe = self.yillikNetKar / ortDefterDegeri
            print("ROE (Özsermaye Karlılığı - Özkaynak Getirisi): ", "{:.2%}".format(self.roe))

        def aktifKarlilikHesapla():
            bilancoDonemiToplamVarliklar = self.getBilancoDegeri("TOPLAM VARLIKLAR", 0)
            dortOncekiBilancoDonemiToplamVarliklar = self.getBilancoDegeri("TOPLAM VARLIKLAR", -4)
            toplamVarliklar = (bilancoDonemiToplamVarliklar + dortOncekiBilancoDonemiToplamVarliklar) / 2
            self.roa = self.yillikNetKar / toplamVarliklar
            print("ROA (Aktif Karlılık): ", "{:.2%}".format(self.roa))

        def yillikNetKarMarjiHesapla():
            self.yillikNetKarMarji = self.yillikNetKar / self.yillikHasilat
            print("Yıllık Net Kar Marjı: ", "{:.2%}".format(self.yillikNetKarMarji))

        def sonCeyrekNetKarMarjiHesapla():
            self.sonCeyrekNetKarMarji = self.netKar0/self.hasilat0
            print("Son Çeyrek Net Kar Marjı: ", "{:.2%}".format(self.sonCeyrekNetKarMarji))

        def aktifDevirHiziHesapla():
            self.aktifDevirHizi = self.yillikHasilat / self.getBilancoDegeri("TOPLAM VARLIKLAR", 0)
            print("Aktif Devir Hızı: ", "{:.2}".format(self.aktifDevirHizi))

        def borcKaynakOraniHesapla():
            self.borcKaynakOrani = self.getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", 0) / self.getBilancoDegeri("TOPLAM KAYNAKLAR", 0)
            print("Borç/Kaynak Oranı: ", "{:.2%}".format(self.borcKaynakOrani))

        def cariOranHesapla():
            donenVarliklar = self.getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", 0)
            kisaVadeliYukumlulukler = self.getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
            self.cariOran = donenVarliklar / kisaVadeliYukumlulukler
            print("Cari Oran: ", "{:.3}".format(self.cariOran))

        def likitOraniHesapla():
            donenVarliklar = self.getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", 0)
            stoklar = self.getBilancoDegeri("Stoklar", 0)
            digerDonenVarliklar = self.getBilancoDegeri("Diğer Dönen Varlıklar", 0)
            kisaVadeliYukumlulukler = self.getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
            self.likitOrani = (donenVarliklar - stoklar - digerDonenVarliklar) / kisaVadeliYukumlulukler
            print("Likit Oranı: ", "{:.3}".format(self.likitOrani))

        def nakitOraniHesapla():
            hazirDegerler = self.getBilancoDegeri("Nakit ve Nakit Benzerleri", 0)
            kisaVadeliYukumlulukler = self.getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
            self.nakitOrani = hazirDegerler / kisaVadeliYukumlulukler
            print("Nakit Oranı: ", "{:.3}".format(self.nakitOrani))

        def asitTestOraniHesapla():
            donenVarliklar = self.getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", 0)
            stoklar = self.getBilancoDegeri("Stoklar", 0)
            kisaVadeliYukumlulukler = self.getBilancoDegeri("TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER", 0)
            self.asitTestOrani = (donenVarliklar - stoklar) / kisaVadeliYukumlulukler
            print("Asit Test Oranı: ", "{:.3}".format(self.asitTestOrani))

        def halkaAciklikOraniniGetir():
            self.halkaAciklikOrani = returnHisseHalkaAciklikOrani(self.hisseAdi)
            print("Halka Açıklık Oranı: ", "{:.2%}".format(self.halkaAciklikOrani))

        def sermayeArtirimPotansiyeliniHesapla():
            odenmisSermaye = self.getBilancoDegeri("Ödenmiş Sermaye", 0)
            ozkaynaklar = self.getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", 0)
            self.sermayeArtirimPotansiyeli = (ozkaynaklar - odenmisSermaye) / odenmisSermaye
            print("Sermaye Artirim Potansiyeli:", "{:.0%}".format(self.sermayeArtirimPotansiyeli))

        def ozsermayeBuyumesiHesapla():
            dortOncekiCeyrekDefterDegeri = self.getBilancoDegeri("Ana Ortaklığa Ait Özkaynaklar", -4)
            self.yillikOzsermayeBuyumesi = self.defterDegeri / dortOncekiCeyrekDefterDegeri
            print("Yıllık Özsermaye Büyümesi: ", "{:.2%}".format(self.yillikOzsermayeBuyumesi))


        def rapor_olustur_excel():
            self.my_logger.debug("")
            self.my_logger.debug("")
            self.my_logger.debug("----------------RAPOR DOSYASI OLUŞTURMA/GÜNCELLEME-------------------------------------")
            self.my_logger.debug (self.hisseAdi)

            excelRow = ExcelRowClass()

            excelRow.bilancoDonemiHasilat = self.hasilat0
            excelRow.oncekiYilAyniCeyrekHasilat = self.hasilat4
            excelRow.bilancoDonemiHasilatDegisimi = self.hasilat0Degisimi
            excelRow.oncekiBilancoDonemiHasilat = self.hasilat1
            excelRow.besOncekiBilancoDonemiHasilat = self.hasilat5
            excelRow.oncekiBilancoDonemiHasilatDegisimi = self.hasilat1Degisimi
            excelRow.bilancoDonemiHasilatDegisimiGecmeDurumu = self.bilancoDonemiHasilatDegisimiGecmeDurumu
            excelRow.oncekiBilancoDonemiHasilatDegisimiGecmeDurumu = self.oncekiDonemHasilatDegisimiGecmeDurumu
            excelRow.bilancoDonemiFaaliyetKari = self.faaliyetKari0
            excelRow.oncekiYilAyniCeyrekFaaliyetKari = self.faaliyetKari4
            excelRow.bilancoDonemiFaaliyetKariDegisimi = self.faaliyetKari0Degisimi
            excelRow.oncekiBilancoDonemiFaaliyetKari = self.faaliyetKari1
            excelRow.besOncekiBilancoDonemiFaaliyetKari = self.faaliyetKari5
            excelRow.oncekiBilancoDonemiFaaliyetKariDegisimi = self.faaliyetKari1Degisimi
            excelRow.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu
            excelRow.oncekiBilancoDonemiFaaliyetKarDegisimiGecmeDurumu = self.birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu
            excelRow.bilancoDonemiOrtalamaDolarKuru = self.ortalamaDolarKuru0
            excelRow.bilancoDonemiDolarHasilat = self.dolarHasilat0
            excelRow.oncekiYilAyniCeyrekDolarHasilat = self.dolarHasilat4
            excelRow.bilancoDonemiDolarHasilatDegisimi = self.dolarHasilat0Degisimi
            excelRow.oncekiBilancoDonemiDolarHasilatDegisimi = self.dolarHasilat1Degisimi
            excelRow.bilancoDonemiDolarHasilatDegisimiGecmeDurumu = self.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu
            excelRow.oncekiBilancoDonemiDolarHasilatDegisimiGecmeDurumu = self.birOncekibilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu
            excelRow.bilancoDonemiDolarFaaliyetKari = self.dolarFaaliyetKari0
            excelRow.dortOncekiBilancoDonemiDolarFaaliyetKari = self.dolarFaaliyetKari4
            excelRow.bilancoDonemiDolarFaaliyetKariDegisimi = self.dolarFaaliyetKari0Degisimi
            excelRow.oncekiBilancoDonemiDolarFaaliyetKariDegisimi = self.dolarFaaliyetKari1Degisimi
            excelRow.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu = self.bilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu
            excelRow.oncekiBilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu = self.birOncekibilancoDonemiDolarFaaliyetKariDegisimiGecmeDurumu
            excelRow.sermaye = self.sermaye
            excelRow.anaOrtaklikPayi = self.anaOrtaklikPayi
            excelRow.sonDortBilancoDonemiHasilatToplami = self.yillikHasilat
            excelRow.onumuzdekiDortBilancoDonemiHasilatTahmini = self. onumuzdekiDortCeyrekHasilatTahmini
            excelRow.onumuzdekiDortBilancoDonemiFaaliyetKarMarjiTahmini = self.onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
            excelRow.faaliyetKariTahmini1 = self.faaliyetKariTahmini1
            excelRow.faaliyetKariTahmini2 = self.faaliyetKariTahmini2
            excelRow.ortalamaFaaliyetKariTahmini = self.ortalamaFaaliyetKariTahmini
            excelRow.hisseBasinaOrtalamaKarTahmini = self.hisseBasinaOrtalamaKarTahmini
            excelRow.bilancoEtkisi = self.bilancoEtkisi
            excelRow.bilancoTarihiHisseFiyati = self.hisseFiyati
            excelRow.gercekHisseDegeri = self.gercekDeger
            excelRow.targetBuy = self.targetBuy
            excelRow.gercekFiyataUzaklik = self.gercekFiyataUzaklik
            excelRow.netProKriteri = self.netProKriteri
            excelRow.forwardPeKriteri = self.forwardPeKriteri

            excelRow.tarih = datetime.today().strftime('%d.%m.%Y')
            excelRow.netKarBuyumeYillik = self.netKarBuyumeOraniYillik
            excelRow.netKarBuyume4OncekiCeyregeGore = self.oncekiYilAyniCeyregeGoreNetKarBuyume
            excelRow.esasFaaliyetKariBuyumeYillik = self.yillikEsasFaaliyetKariBuyumeOrani
            excelRow.hasilatBuyumeYillik = self.yillikHasilatBuyumeOrani
            excelRow.favokBuyumeYillik = self.yillikFavokArtisOrani
            excelRow.fkOrani = self.fkOrani
            excelRow.nakitPd = self.nakitPd
            excelRow.nakitFd = self.nakitFd
            excelRow.pdDd = self.pdDd
            excelRow.pegOrani = self.pegOrani
            excelRow.fdSatislar = self.fdSatislar
            excelRow.fdFavok = self.fdFavok
            excelRow.pdEfk = self.pdEfk
            excelRow.cariOran = self.cariOran
            excelRow.likitOrani = self.likitOrani
            excelRow.nakitOrani = self.nakitOrani
            excelRow.asitTestOrani = self.asitTestOrani
            excelRow.roe = self.roe
            excelRow.roa = self.roa
            excelRow.yillikNetKarMarji = self.yillikNetKarMarji
            excelRow.sonCeyrekNetKarMarji = self.sonCeyrekNetKarMarji
            excelRow.aktifDevirHizi = self.aktifDevirHizi
            excelRow.borcKaynak = self.borcKaynakOrani
            excelRow.ozsermayeBuyumesi = self.yillikOzsermayeBuyumesi
            excelRow.halkaAciklikOrani = self.halkaAciklikOrani
            excelRow.piyasaDegeri = (int)(self.piyasaDegeri/1000000)
            excelRow.sermaye = (int)(self.sermaye/1000000)
            excelRow.sermayeArtirimPotansiyeli = self.sermayeArtirimPotansiyeli

            exportReportExcel(self.hisseAdi, self.reportFile, self.bilancoDonemi, excelRow)

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info ("-------------------------------- %s ------------------------", self.hisseAdi)

            self.my_logger.removeHandler(self.output_file_handler)
            self.my_logger.removeHandler(self.stdout_handler)



        hasilat_hesaplari()
        faaliyet_kari_hesaplari()
        net_kar_hesaplari()
        brut_kar_hesaplari()
        gercek_deger_hesabi()
        netpro_kriteri_hesabi()
        bilanco_donemi_dolar_hesabi()
        dolar_hasilat_hesaplari()
        dolar_faaliyet_kari_hesaplari()

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
        favokHesapla()
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
        cariOranHesapla()
        likitOraniHesapla()
        nakitOraniHesapla()
        asitTestOraniHesapla()
        halkaAciklikOraniniGetir()
        sermayeArtirimPotansiyeliniHesapla()
        ozsermayeBuyumesiHesapla()
        rapor_olustur_excel()


