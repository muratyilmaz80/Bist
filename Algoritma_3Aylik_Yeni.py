from ExcelRowClass import ExcelRowClass
from Rapor_Olustur import exportReportExcel
from prettytable import PrettyTable
import logging,sys,math, xlrd
from BilancoOrtalamaDolarDegeri import ucAylikBilancoDonemiOrtalamaDolarDegeriBul
import pandas as pd

class Algoritma():

    def __init__(self, bilancoDosyasi, bilancoDonemi, bondYield, hisseFiyati, reportFile, logPath, logLevel):
        hisseAdiTemp = bilancoDosyasi[64:]
        hisseAdi = hisseAdiTemp[:-5]
        self.bilancoDosyasi = bilancoDosyasi
        self.hisseFiyati = hisseFiyati
        self.bondYield = bondYield
        self.my_logger = logging.getLogger()
        self.my_logger.setLevel(logLevel)
        self.bilancoDonemi = bilancoDonemi
        output_file_handler = logging.FileHandler(logPath + hisseAdi + ".txt")
        output_file_handler.level = logging.INFO
        stdout_handler = logging.StreamHandler(sys.stdout)
        self.my_logger.addHandler(output_file_handler)
        self.my_logger.addHandler(stdout_handler)
        self.bd_df = pd.read_excel(self.bilancoDosyasi, index_col=0)
        self.cok_kullanilan_degerleri_hesapla()
        self.my_logger.info ("-------------------------------- %s ------------------------", hisseAdi)

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
            return -999;
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


    def cok_kullanilan_degerleri_hesapla(self):
        self.hasilat0 = self.ceyrekDegeriHesapla("Hasılat", 0)
        self.hasilat1 = self.ceyrekDegeriHesapla("Hasılat", -1)
        self.hasilat2 = self.ceyrekDegeriHesapla("Hasılat", -2)
        self.hasilat3 = self.ceyrekDegeriHesapla("Hasılat", -3)
        self.hasilat4 = self.ceyrekDegeriHesapla("Hasılat", -4)
        self.hasilat5 = self.ceyrekDegeriHesapla("Hasılat", -5)
        self.hasilat6 = self.ceyrekDegeriHesapla("Hasılat", -6)
        self.hasilat7 = self.ceyrekDegeriHesapla("Hasılat", -7)

        self.faaliyetKari0 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", 0)
        self.faaliyetKari1 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -1)
        self.faaliyetKari2 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -2)
        self.faaliyetKari3 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -3)
        self.faaliyetKari4 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -4)
        self.faaliyetKari5 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -5)
        self.faaliyetKari6 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -6)
        self.faaliyetKari7 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", -7)

        self.netKar0 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", 0)
        self.netKar1 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", -1)
        self.netKar2 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", -2)
        self.netKar3 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", -3)
        self.netKar4 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", -4)
        self.netKar5 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", -5)
        self.netKar6 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", -6)
        self.netKar7 = self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", -7)

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

        self.alacaklar = self.getBilancoDegeri("Dönen Ticari Alacaklar", 0) + \
                    self.getBilancoDegeri("Dönen Diğer Alacaklar", 0) + \
                    self.getBilancoDegeri("Duran Ticari Alacaklar", 0) + \
                    self.getBilancoDegeri("Duran Diğer Alacaklar", 0)

        self.finansalVarliklar = self.getBilancoDegeri("Duran Finansal Yatırımlar", 0) + \
                            self.getBilancoDegeri("Dönen Finansal Yatırımlar", 0) + \
                            self.getBilancoDegeri("Özkaynak Yöntemiyle Değerlenen Yatırımlar", 0)

        self.maddiDuranVarliklar = self.getBilancoDegeri("Maddi Duran Varlıklar", 0)


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


    def likidasyonDegeriHesapla(self):
        likidasyonDegeri = self.nakit + (self.alacaklar * 0.7) + (self.stoklar * 0.5) + (self.digerVarliklar * 0.7) + (self.finansalVarliklar * 0.7) + (self.maddiDuranVarliklar * 0.2)
        return likidasyonDegeri




    def runAlgoritma(self):

        self.my_logger.debug("Bilanco Donemi: %d", self.bilancoDonemi)





        def hasilat_hesaplari(ceyrek):

            # Bilanço Dönemi Satış(Hasılat) Gelirleri
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------HASILAT(SATIŞ) GELİRLERİ---------------------------")
            self.my_logger.info("")

            hasilat0Print = "{:,.0f}".format(self.hasilat0).replace(",", ".")
            hasilat4Print = "{:,.0f}".format(self.hasilat4).replace(",", ".")
            hasilat0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek)
            hasilat0DegisimiPrint = "{:.2%}".format(hasilat0Degisimi)

            hasilat1Print = "{:,.0f}".format(self.hasilat1).replace(",", ".")
            hasilat5Print = "{:,.0f}".format(self.hasilat5).replace(",", ".")
            hasilat1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek-1)
            hasilat1DegisimiPrint = "{:.2%}".format(hasilat1Degisimi)

            hasilat2Print = "{:,.0f}".format(self.hasilat2).replace(",", ".")
            hasilat6Print = "{:,.0f}".format(self.hasilat6).replace(",", ".")
            hasilat2Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek-2)
            hasilat2DegisimiPrint = "{:.2%}".format(hasilat2Degisimi)

            hasilat3Print = "{:,.0f}".format(self.hasilat3).replace(",", ".")
            hasilat7Print = "{:,.0f}".format(self.hasilat7).replace(",", ".")
            hasilat3Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek-3)
            hasilat3DegisimiPrint = "{:.2%}".format(hasilat3Degisimi)

            satisTablosu = PrettyTable()
            satisTablosu.field_names = ["ÇEYREK", "SATIŞ", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ", "YÜZDE DEĞİŞİM"]
            satisTablosu.align["SATIŞ"] = "r"
            satisTablosu.align["ÖNCEKİ YIL SATIŞ"] = "r"
            satisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            satisTablosu.add_row([self.bilancoDoneminiBul(ceyrek), hasilat0Print, self.bilancoDoneminiBul(ceyrek-4), hasilat4Print, hasilat0DegisimiPrint])
            satisTablosu.add_row([self.bilancoDoneminiBul(ceyrek-1), hasilat1Print, self.bilancoDoneminiBul(ceyrek-5), hasilat5Print, hasilat1DegisimiPrint])
            satisTablosu.add_row([self.bilancoDoneminiBul(ceyrek-2), hasilat2Print, self.bilancoDoneminiBul(ceyrek-6), hasilat6Print, hasilat2DegisimiPrint])
            satisTablosu.add_row([self.bilancoDoneminiBul(ceyrek-3), hasilat3Print, self.bilancoDoneminiBul(ceyrek-7),hasilat7Print, hasilat3DegisimiPrint])
            self.my_logger.info(satisTablosu)

            # Bilanço Dönemi Satış Geliri Artış Kriteri
            self.bilancoDonemiHasilatGelirArtisiGecmeDurumu = (hasilat0Degisimi > 0.1)
            printText = "Bilanço Dönemi Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(hasilat0Degisimi) + " >? 10% " + " " + str(self.bilancoDonemiHasilatGelirArtisiGecmeDurumu)
            self.my_logger.info(printText)

            # Önceki Dönem Hasılat Geliri Artış Kriteri

            if (hasilat0Degisimi >= 1):
                self.my_logger.info ("Bilanço Dönemi Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak.")
                oncekiDonemHasilatGelirArtisiGecmeDurumu = True
                self.my_logger.info ("Önceki Dönem Satış Gelir Artışı Geçme Durumu: %s", oncekiDonemHasilatGelirArtisiGecmeDurumu)

            else:
                oncekiDonemHasilatGelirArtisiGecmeDurumu = (hasilat1Degisimi < hasilat0Degisimi)
                printText = "Önceki Dönem Satış Gelir Artışı Bilanço Döneminden Düşük Mü: " + "{:.2%}".format(hasilat1Degisimi) + " <? " + "{:.2%}".format(hasilat0Degisimi) + " " + str(oncekiDonemHasilatGelirArtisiGecmeDurumu)
                self.my_logger.info(printText)



        def faaliyet_kari_hesaplari(ceyrek):

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------------FAALİYET KARI---------------------------------")
            self.my_logger.info("")


            faaliyetKari0Print = "{:,.0f}".format(self.faaliyetKari0).replace(",", ".")
            faaliyetKari4Print = "{:,.0f}".format(self.faaliyetKari4).replace(",", ".")
            faaliyetKari0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek)
            faaliyetKari0DegisimiPrint = "{:.2%}".format(faaliyetKari0Degisimi)

            faaliyetKari1Print = "{:,.0f}".format(self.faaliyetKari1).replace(",", ".")
            faaliyetKari5Print = "{:,.0f}".format(self.faaliyetKari5).replace(",", ".")
            faaliyetKari1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek-1)
            faaliyetKari1DegisimiPrint = "{:.2%}".format(faaliyetKari1Degisimi)

            faaliyetKari2Print = "{:,.0f}".format(self.faaliyetKari2).replace(",", ".")
            faaliyetKari6Print = "{:,.0f}".format(self.faaliyetKari6).replace(",", ".")
            faaliyetKari2Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek-2)
            faaliyetKari2DegisimiPrint = "{:.2%}".format(faaliyetKari2Degisimi)

            faaliyetKari3Print = "{:,.0f}".format(self.faaliyetKari3).replace(",", ".")
            faaliyetKari7Print = "{:,.0f}".format(self.faaliyetKari7).replace(",", ".")
            faaliyetKari3Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek-3)
            faaliyetKari3DegisimiPrint = "{:.2%}".format(faaliyetKari3Degisimi)

            faaliyetKariTablosu = PrettyTable()
            faaliyetKariTablosu.field_names = ["ÇEYREK", "FAALİYET KARI", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI", "YÜZDE DEĞİŞİM"]
            faaliyetKariTablosu.align["FAALİYET KARI"] = "r"
            faaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI"] = "r"
            faaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(ceyrek), faaliyetKari0Print, self.bilancoDoneminiBul(ceyrek-4), faaliyetKari4Print, faaliyetKari0DegisimiPrint])
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(ceyrek-1), faaliyetKari1Print, self.bilancoDoneminiBul(ceyrek-5), faaliyetKari5Print, faaliyetKari1DegisimiPrint])
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(ceyrek-2), faaliyetKari2Print, self.bilancoDoneminiBul(ceyrek-6), faaliyetKari6Print, faaliyetKari2DegisimiPrint])
            faaliyetKariTablosu.add_row([self.bilancoDoneminiBul(ceyrek-3), faaliyetKari3Print, self.bilancoDoneminiBul(ceyrek-7),faaliyetKari7Print, faaliyetKari3DegisimiPrint])
            self.my_logger.info(faaliyetKariTablosu)


            # Bilanço Dönemi Faaliyet Kar Artış Kriteri
            if self.ceyrekDegeriHesapla("DÖNEM KARI (ZARARI)", 0) < 0:
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = False
                self.my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Çeyrek Net Kar Negatif", str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu))

            elif self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek) < 0:
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = False
                self.my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Ceyrek Faaliyet Kari Negatif", str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu))

            elif ((self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek) > 0) and (self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek-4)) < 0):
                faaliyetKari0Degisimi = 0
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = True
                self.my_logger.info("Bilanço Dönemi Faaliyet Karı Artışı: %s Son Çeyrek Faaliyet Karı Negatiften Pozitife Geçmiş", str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu))

            else:
                self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu = (faaliyetKari0Degisimi > 0.15)
                printText = "Bilanço Dönemi Faaliyet Karı Artışı:" + "{:.2%}".format(faaliyetKari0Degisimi) + " >? 15% " + str(self.bilancoDonemiFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)

            # Önceki Dönem Faaliyet Kar Artış Kriteri

            if faaliyetKari0Degisimi >= 1:
                birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu = True
                printText = "Önceki Dönem Faaliyet Kar Artışı: Bilanço Dönemi Faaliyet Karı Artışı 100%'ün Üzerinde, Karşılaştırma Yapılmayacak: " + "{:.2%}".format(faaliyetKari0Degisimi) + " " + str(birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)

            else:
                birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu = (faaliyetKari1Degisimi < faaliyetKari0Degisimi)
                printText = "Önceki Dönem Faaliyet Kar Artışı:" + "{:.2%}".format(faaliyetKari1Degisimi) + " < ? " + "{:.2%}".format(faaliyetKari0Degisimi) + str(birOncekibilancoDonemiFaaliyetKariDegisimiGecmeDurumu)
                self.my_logger.info(printText)




        def net_kar_hesaplari(ceyrek):

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("-------------------NET KAR (DÖNEM KARI/ZARARI)--------------------------")
            self.my_logger.info("")



            netKar0Print = "{:,.0f}".format(self.netKar0).replace(",", ".")
            netKar4Print = "{:,.0f}".format(self.netKar4).replace(",", ".")
            netKar0DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("DÖNEM KARI (ZARARI)", ceyrek))

            netKar1Print = "{:,.0f}".format(self.netKar1).replace(",", ".")
            netKar5Print = "{:,.0f}".format(self.netKar5).replace(",", ".")
            netKar1DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("DÖNEM KARI (ZARARI)", ceyrek-1))

            netKar2Print = "{:,.0f}".format(self.netKar2).replace(",", ".")
            netKar6Print = "{:,.0f}".format(self.netKar6).replace(",", ".")
            netKar2DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("DÖNEM KARI (ZARARI)", ceyrek-2))

            netKar3Print = "{:,.0f}".format(self.netKar3).replace(",", ".")
            netKar7Print = "{:,.0f}".format(self.netKar7).replace(",", ".")
            netKar3DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("DÖNEM KARI (ZARARI)", ceyrek-3))

            netKarTablosu = PrettyTable()
            netKarTablosu.field_names = ["ÇEYREK", "NET KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL NET KAR", "YÜZDE DEĞİŞİM"]
            netKarTablosu.align["NET KAR"] = "r"
            netKarTablosu.align["ÖNCEKİ YIL NET KAR"] = "r"
            netKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            netKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek), netKar0Print, self.bilancoDoneminiBul(ceyrek-4), netKar4Print, netKar0DegisimiPrint])
            netKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek-1), netKar1Print, self.bilancoDoneminiBul(ceyrek-5), netKar5Print, netKar1DegisimiPrint])
            netKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek-2), netKar2Print, self.bilancoDoneminiBul(ceyrek-6),netKar6Print, netKar2DegisimiPrint])
            netKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek-3), netKar3Print, self.bilancoDoneminiBul(ceyrek-7),netKar7Print, netKar3DegisimiPrint])
            self.my_logger.info(netKarTablosu)


        def brut_kar_hesaplari(ceyrek):
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("-------------------BRÜT KAR (BRÜT KAR/ZARAR)--------------------------")
            self.my_logger.info("")

            brutKar0Print = "{:,.0f}".format(self.brutKar0).replace(",", ".")
            brutKar4Print = "{:,.0f}".format(self.brutKar4).replace(",", ".")
            brutKar0DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", ceyrek))

            brutKar1Print = "{:,.0f}".format(self.brutKar1).replace(",", ".")
            brutKar5Print = "{:,.0f}".format(self.brutKar5).replace(",", ".")
            brutKar1DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", ceyrek-1))

            brutKar2Print = "{:,.0f}".format(self.brutKar2).replace(",", ".")
            brutKar6Print = "{:,.0f}".format(self.brutKar6).replace(",", ".")
            brutKar2DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", ceyrek-2))

            brutKar3Print = "{:,.0f}".format(self.brutKar3).replace(",", ".")
            brutKar7Print = "{:,.0f}".format(self.brutKar7).replace(",", ".")
            brutKar3DegisimiPrint = "{:.2%}".format(self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("BRÜT KAR (ZARAR)", ceyrek-3))

            brutKarTablosu = PrettyTable()
            brutKarTablosu.field_names = ["ÇEYREK", "BRÜT KAR", "ÖNCEKİ YIL", "ÖNCEKİ YIL BRÜT KAR", "YÜZDE DEĞİŞİM"]
            brutKarTablosu.align["BRÜT KAR"] = "r"
            brutKarTablosu.align["ÖNCEKİ YIL BRÜT KAR"] = "r"
            brutKarTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            brutKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek), brutKar0Print, self.bilancoDoneminiBul(ceyrek-4), brutKar4Print, brutKar0DegisimiPrint])
            brutKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek-1), brutKar1Print, self.bilancoDoneminiBul(ceyrek-5), brutKar5Print, brutKar1DegisimiPrint])
            brutKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek-2), brutKar2Print, self.bilancoDoneminiBul(ceyrek-6), brutKar6Print, brutKar2DegisimiPrint])
            brutKarTablosu.add_row([self.bilancoDoneminiBul(ceyrek-3), brutKar3Print, self.bilancoDoneminiBul(ceyrek-7), brutKar7Print, brutKar3DegisimiPrint])
            self.my_logger.info(brutKarTablosu)



        def gercek_deger_hesabi(ceyrek):
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("----------------GERÇEK DEĞER HESABI--------------------------------------------")

            sermaye = self.getBilancoDegeri("Ödenmiş Sermaye", ceyrek)
            self.my_logger.info("Sermaye: %s TL", "{:,.0f}".format(sermaye).replace(",", "."))

            anaOrtaklikPayi = self.getBilancoDegeri("Ana Ortaklık Payları", ceyrek) / self.getBilancoDegeri("DÖNEM KARI (ZARARI)", ceyrek)
            self.my_logger.info("Ana Ortaklık Payı: %s", "{:.3f}".format(anaOrtaklikPayi))

            hasilat0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek)
            hasilat1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek-1)

            sonDortCeyrekHasilatToplami = self.yilliklandirmisDegerHesapla("Hasılat", ceyrek)
            self.my_logger.info("Son 4 Çeyrek Hasılat Toplamı: %s TL","{:,.0f}".format(sonDortCeyrekHasilatToplami).replace(",", "."))

            onumuzdekiDortCeyrekHasilatTahmini = ((((hasilat0Degisimi + hasilat1Degisimi) / 2) + 1) * sonDortCeyrekHasilatToplami)

            hasilatlarCeyrek = [self.hasilat3, self.hasilat2, self.hasilat1, self.hasilat0]
            maxHasilatCeyrek = max(hasilatlarCeyrek)

            self.my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini: %s TL","{:,.0f}".format(onumuzdekiDortCeyrekHasilatTahmini).replace(",", "."))

            if (onumuzdekiDortCeyrekHasilatTahmini > 4 * maxHasilatCeyrek):
                onumuzdekiDortCeyrekHasilatTahmini = 4 * maxHasilatCeyrek
                self.my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini 4*maxCeyrek olarak duzeltildi:")
                self.my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini: %s TL",
                               "{:,.0f}".format(onumuzdekiDortCeyrekHasilatTahmini).replace(",", "."))

            # HASILAT TAHMININI MANUEL DEGISTIRMEK ICIN
            # onumuzdekiDortCeyrekHasilatTahmini = 700000000000

            faaliyetKari0 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek)
            faaliyetKari1 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek-1)
            faaliyetKari2 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek-2)
            faaliyetKari3 = self.ceyrekDegeriHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek-3)

            onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (faaliyetKari1 + faaliyetKari0) / (self.hasilat0 + self.hasilat1)
            self.my_logger.info("Önümüzdeki 4 Çeyrek Faaliyet Kar Marjı Tahmini: %s ","{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))

            faaliyetKariTahmini1 = onumuzdekiDortCeyrekHasilatTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
            self.my_logger.info("Faaliyet Kar Tahmini1: %s TL", "{:,.0f}".format(faaliyetKariTahmini1).replace(",", "."))

            faaliyetKariTahmini2 = ((faaliyetKari1 + faaliyetKari0) * 2 * 0.3) + (faaliyetKari0 * 4 * 0.5) + ((faaliyetKari3 + faaliyetKari2 + faaliyetKari1 + faaliyetKari0) * 0.2)
            self.my_logger.info("Faaliyet Kar Tahmini2: %s TL", "{:,.0f}".format(faaliyetKariTahmini2).replace(",", "."))

            ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1 + faaliyetKariTahmini2) / 2
            self.my_logger.info("Ortalama Faaliyet Kari Tahmini: %s TL","{:,.0f}".format(ortalamaFaaliyetKariTahmini).replace(",", "."))

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

            hisseBasinaOrtalamaKarTahmini = ((ortalamaFaaliyetKariTahmini) * anaOrtaklikPayi) / sermaye
            self.my_logger.info("Hisse Başına Ortalama Kar Tahmini: %s TL", format(hisseBasinaOrtalamaKarTahmini, ".2f"))

            likidasyonDegeri = self.likidasyonDegeriHesapla()
            self.my_logger.info("Likidasyon Değeri: %s TL", "{:,.0f}".format(likidasyonDegeri).replace(",", "."))

            borclar = int(self.getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", ceyrek))
            self.my_logger.info("Borçlar: %s TL", "{:,.0f}".format(borclar).replace(",", "."))

            bilancoEtkisi = (likidasyonDegeri - borclar) / sermaye * anaOrtaklikPayi
            self.my_logger.info("Bilanço Etkisi: %s TL", format(bilancoEtkisi, ".2f"))

            gercekDeger = (hisseBasinaOrtalamaKarTahmini * 7) + bilancoEtkisi
            self.my_logger.info("Gerçek Hisse Değeri: %s TL", format(gercekDeger, ".2f"))

            targetBuy = gercekDeger * 0.66
            self.my_logger.info("Target Buy: %s TL", format(targetBuy, ".2f"))

            self.my_logger.info("Bilanço Tarihindeki Hisse Fiyatı: %s TL", format(self.hisseFiyati, ".2f"))

            gercekFiyataUzaklik = self.hisseFiyati / targetBuy
            self.my_logger.info("Gerçek Fiyata Uzaklık Oranı: %s", "{:.2%}".format(gercekFiyataUzaklik))

            gercekFiyataUzaklikTl = self.hisseFiyati - targetBuy
            self.my_logger.info("Gerçek Fiyata Uzaklık %s TL:", format(gercekFiyataUzaklikTl, ".2f"))




        def netpro_kriteri_hesabi(ceyrek):
            # Netpro Hesapla
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("----------------NETPRO  ve FORWARD_PE KRİTERİ-----------------------------------------------------------------")

            sonDortDonemFaaliyetKariToplami = self.yilliklandirmisDegerHesapla("ESAS FAALİYET KARI (ZARARI)", ceyrek)
            sonDortDonemNetKarToplami = self.yilliklandirmisDegerHesapla("DÖNEM KARI (ZARARI)", ceyrek)

            self.my_logger.info("Son 4 Dönem Net Kar Toplamı: %s TL", "{:,.0f}".format(sonDortDonemNetKarToplami).replace(",", "."))
            self.my_logger.info("Son 4 Dönem Faaliyet Karı Toplamı: %s TL", "{:,.0f}".format(sonDortDonemFaaliyetKariToplami).replace(",", "."))

            anaOrtaklikPayi = self.getBilancoDegeri("Ana Ortaklık Payları", ceyrek) / self.getBilancoDegeri("DÖNEM KARI (ZARARI)", ceyrek)
            sermaye = self.getBilancoDegeri("Ödenmiş Sermaye", ceyrek)

            fkOrani = self.hisseFiyati / ((sonDortDonemNetKarToplami * anaOrtaklikPayi) / (sermaye))
            self.my_logger.info("F/K Oranı: %s", "{:,.2f}".format(fkOrani))

            hbkOrani = sonDortDonemNetKarToplami / (sermaye) * anaOrtaklikPayi
            self.my_logger.info("HBK Oranı: %s", "{:,.2f}".format(hbkOrani))

            hasilat0Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek)
            hasilat1Degisimi = self.onceki_yil_ayni_ceyrege_gore_degisimi_hesapla("Hasılat", ceyrek-1)
            sonDortCeyrekHasilatToplami = self.yilliklandirmisDegerHesapla("Hasılat", ceyrek)
            onumuzdekiDortCeyrekHasilatTahmini = ((((hasilat0Degisimi + hasilat1Degisimi) / 2) + 1) * sonDortCeyrekHasilatToplami)

            onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (self.faaliyetKari1 + self.faaliyetKari0) / (self.hasilat0 + self.hasilat1)
            self.my_logger.info("Önümüzdeki 4 Çeyrek Faaliyet Kar Marjı Tahmini: %s ","{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))
            faaliyetKariTahmini1 = onumuzdekiDortCeyrekHasilatTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
            faaliyetKariTahmini2 = ((self.faaliyetKari1 + self.faaliyetKari0) * 2 * 0.3) + (self.faaliyetKari0 * 4 * 0.5) + ((self.faaliyetKari3 + self.faaliyetKari2 + self.faaliyetKari1 + self.faaliyetKari0) * 0.2)
            ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1 + faaliyetKariTahmini2) / 2

            netProEstDegeri = ((ortalamaFaaliyetKariTahmini / sonDortDonemFaaliyetKariToplami) * sonDortDonemNetKarToplami) * anaOrtaklikPayi
            self.my_logger.info("NetPro Est Değeri: %s TL", "{:,.0f}".format(netProEstDegeri).replace(",", "."))

            likidasyonDegeri = self.likidasyonDegeriHesapla()
            borclar = int(self.getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", ceyrek))
            bilancoEtkisi = (likidasyonDegeri - borclar) / sermaye * anaOrtaklikPayi
            piyasaDegeri = (bilancoEtkisi * sermaye * -1) + (self.hisseFiyati * sermaye)

            self.my_logger.info("Piyasa Değeri: %s TL", "{:,.0f}".format(piyasaDegeri).replace(",", "."))
            self.my_logger.info("self.bondYield: %s", "{:.2%}".format(self.bondYield))

            netProKriteri = (netProEstDegeri / piyasaDegeri) / self.bondYield
            netProKriteriGecmeDurumu = (netProKriteri > 2)
            self.my_logger.info("NetPro Kriteri (2'den Büyük Olmalı): %s %s", format(netProKriteri, ".2f"), str(netProKriteriGecmeDurumu))

            minNetProIcinhisseFiyati = (netProEstDegeri / (1.9 * self.bondYield) - (bilancoEtkisi * sermaye * -1)) / sermaye
            self.my_logger.info("NetPro 1.9 Olması İçin Olması Gereken Hisse Fiyatı: %s", format(minNetProIcinhisseFiyati, ".2f"))

            forwardPeKriteri = (piyasaDegeri) / netProEstDegeri
            forwardPeKriteriGecmeDurumu = (forwardPeKriteri < 4)
            printText = "Forward PE Kriteri (4'ten Küçük Olmalı): " + format(forwardPeKriteri, ".2f") + " " + str(forwardPeKriteriGecmeDurumu)
            self.my_logger.info(printText)



        def bilanco_donemi_dolar_hesabi(donem):

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("----------------BİLANÇO DOLAR HESABI-------------------------------------")
            self.my_logger.info("")

            self.ortalamaDolarKuru0 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem))
            self.my_logger.info ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(donem) , "{:,.2f}".format(self.ortalamaDolarKuru0))

            self.ortalamaDolarKuru1 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem-1))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(donem-1) , "{:,.2f}".format(self.ortalamaDolarKuru1))

            self.ortalamaDolarKuru2 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem-2))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(donem-2) ,"{:,.2f}".format(self.ortalamaDolarKuru2))

            self.ortalamaDolarKuru3 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem-3))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(donem-3) ,"{:,.2f}".format(self.ortalamaDolarKuru3))

            self.ortalamaDolarKuru4 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem-4))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(donem-4) ,"{:,.2f}".format(self.ortalamaDolarKuru4))

            self.ortalamaDolarKuru5 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem-5))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL", self.bilancoDoneminiBul(donem-5), "{:,.2f}".format(self.ortalamaDolarKuru5))

            self.ortalamaDolarKuru6 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem-6))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL" , self.bilancoDoneminiBul(donem-6), "{:,.2f}".format(self.ortalamaDolarKuru6))

            self.ortalamaDolarKuru7 = ucAylikBilancoDonemiOrtalamaDolarDegeriBul(self.bilancoDoneminiBul(donem-7))
            self.my_logger.debug ("%s Bilanço Dönemi Ortalama Dolar Kuru: %s TL", self.bilancoDoneminiBul(donem-7), "{:,.2f}".format(self.ortalamaDolarKuru7))






        def diger_hesaplar():

            # Bilanço Dönemi Satış(Hasılat) Gelirleri (DOLAR)
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------HASILAT(SATIŞ) GELİRLERİ (DOLAR)----------------------")
            self.my_logger.info("")

            self.bilancoDonemiDolarHasilat = self.bilancoDonemiHasilat/self.ortalamaDolarKuru0
            birOncekibilancoDonemiDolarHasilat = birOncekibilancoDonemiHasilat/birOncekibilancoDonemiOrtalamaDolarKuru
            ikiOncekiself.bilancoDonemiDolarHasilat = ikiOncekiself.bilancoDonemiHasilat/ikiOncekiself.bilancoDonemiOrtalamaDolarKuru
            ucOncekiself.bilancoDonemiDolarHasilat = ucOncekiself.bilancoDonemiHasilat/ucOncekiself.bilancoDonemiOrtalamaDolarKuru
            oncekiYilAyniCeyrekDolarHasilat = dortOncekiself.bilancoDonemiHasilat/dortOncekiself.bilancoDonemiOrtalamaDolarKuru
            besOncekiself.bilancoDonemiDolarHasilat = besOncekiself.bilancoDonemiHasilat/besOncekiself.bilancoDonemiOrtalamaDolarKuru
            altiOncekiself.bilancoDonemiDolarHasilat = altiOncekiself.bilancoDonemiHasilat/altiOncekiself.bilancoDonemiOrtalamaDolarKuru
            yediOncekiself.bilancoDonemiDolarHasilat = yediOncekiself.bilancoDonemiHasilat/yediOncekiself.bilancoDonemiOrtalamaDolarKuru

            self.bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(self.bilancoDonemiDolarHasilat).replace(",", ".")
            dortOncekiself.bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(oncekiYilAyniCeyrekDolarHasilat).replace(",", ".")
            self.bilancoDonemiDolarHasilatDegisimi = self.bilancoDonemiDolarHasilat/oncekiYilAyniCeyrekDolarHasilat-1
            self.bilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(self.bilancoDonemiDolarHasilatDegisimi)

            birOncekibilancoDonemiDolarHasilatPrint = "{:,.0f}".format(birOncekibilancoDonemiDolarHasilat).replace(",", ".")
            besOncekiself.bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(besOncekiself.bilancoDonemiDolarHasilat).replace(",", ".")
            birOncekibilancoDonemiDolarHasilatDegisimi = birOncekibilancoDonemiDolarHasilat/besOncekiself.bilancoDonemiDolarHasilat-1
            birOncekibilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(birOncekibilancoDonemiDolarHasilatDegisimi)

            ikiOncekiself.bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(ikiOncekiself.bilancoDonemiDolarHasilat).replace(",", ".")
            altiOncekiself.bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(altiOncekiself.bilancoDonemiDolarHasilat).replace(",", ".")
            ikiOncekiself.bilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(ikiOncekiself.bilancoDonemiDolarHasilat/altiOncekiself.bilancoDonemiDolarHasilat-1)

            ucOncekiself.bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(ucOncekiself.bilancoDonemiDolarHasilat).replace(",", ".")
            yediOncekiself.bilancoDonemiDolarHasilatPrint = "{:,.0f}".format(yediOncekiself.bilancoDonemiDolarHasilat).replace(",", ".")
            ucOncekiself.bilancoDonemiDolarHasilatDegisimiPrint = "{:.2%}".format(ucOncekiself.bilancoDonemiDolarHasilat/yediOncekiself.bilancoDonemiDolarHasilat-1)

            dolarSatisTablosu = PrettyTable()
            dolarSatisTablosu.field_names = ["ÇEYREK", "SATIŞ (USD)", "ÖNCEKİ YIL", "ÖNCEKİ YIL SATIŞ (USD)", "YÜZDE DEĞİŞİM"]
            dolarSatisTablosu.align["SATIŞ (USD)"] = "r"
            dolarSatisTablosu.align["ÖNCEKİ YIL SATIŞ (USD)"] = "r"
            dolarSatisTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            dolarSatisTablosu.add_row([self.bilancoDonemi, self.bilancoDonemiDolarHasilatPrint, dortOncekiself.bilancoDonemi, dortOncekiself.bilancoDonemiDolarHasilatPrint, self.bilancoDonemiDolarHasilatDegisimiPrint])
            dolarSatisTablosu.add_row([birOncekibilancoDonemi, birOncekibilancoDonemiDolarHasilatPrint, besOncekiself.bilancoDonemi, besOncekiself.bilancoDonemiDolarHasilatPrint, birOncekibilancoDonemiDolarHasilatDegisimiPrint])
            dolarSatisTablosu.add_row([ikiOncekiself.bilancoDonemi, ikiOncekiself.bilancoDonemiDolarHasilatPrint, altiOncekiself.bilancoDonemi, altiOncekiself.bilancoDonemiDolarHasilatPrint, ikiOncekiself.bilancoDonemiDolarHasilatDegisimiPrint])
            dolarSatisTablosu.add_row([ucOncekiself.bilancoDonemi, ucOncekiself.bilancoDonemiDolarHasilatPrint, yediOncekiself.bilancoDonemi,yediOncekiself.bilancoDonemiDolarHasilatPrint, ucOncekiself.bilancoDonemiDolarHasilatDegisimiPrint])
            self.my_logger.info (dolarSatisTablosu)

            # Bilanço Dönemi (DOLAR) Satış Geliri Artış Kriteri
            self.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (self.bilancoDonemiDolarHasilatDegisimi > 0.1)

            printText = "Bilanço Dönemi (DOLAR) Satış Geliri Artışı 10%'dan Büyük Mü: " + "{:.2%}".format(self.bilancoDonemiDolarHasilatDegisimi) + " >? 10%" + " " + str(self.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
            self.my_logger.info(printText)



            # Önceki Dönem (DOLAR) Hasılat Geliri Artış Kriteri
            #
            if (self.bilancoDonemiDolarHasilatDegisimi >= 1):
                printText = "Bilanço Dönemi (DOLAR) Satış Gelir Artışı %100 Üzerinde, Karşılaştırma Yapılmayacak."
                self.my_logger.info (printText)
                oncekiself.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = True
                printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Geçme Durumu: " + str(oncekiself.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
                self.my_logger.info (printText)

            else:
                oncekiself.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = (birOncekibilancoDonemiDolarHasilatDegisimi<self.bilancoDonemiDolarHasilatDegisimi)
                printText = "Önceki Dönem (DOLAR) Satış Gelir Artışı Bilanço Döneminden Düşük Mü: " + "{:.2%}".format(birOncekibilancoDonemiDolarHasilatDegisimi) + \
                            " <? " + "{:.2%}".format(self.bilancoDonemiDolarHasilatDegisimi) + " " + str(oncekiself.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu)
                self.my_logger.info(printText)



            # Bilanço Dönemi Faaliyet Karı Gelirleri (DOLAR)
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("--------------------------FAALİYET KARI (DOLAR)-------------------------")
            self.my_logger.info("")

            self.bilancoDonemiDolarFaaliyetKari = self.bilancoDonemiFaaliyetKari/self.bilancoDonemiOrtalamaDolarKuru
            birOncekibilancoDonemiDolarFaaliyetKari = birOncekibilancoDonemiFaaliyetKari/birOncekibilancoDonemiOrtalamaDolarKuru
            ikiOncekiself.bilancoDonemiDolarFaaliyetKari = ikiOncekiself.bilancoDonemiFaaliyetKari/ikiOncekiself.bilancoDonemiOrtalamaDolarKuru
            ucOncekiself.bilancoDonemiDolarFaaliyetKari = ucOncekiself.bilancoDonemiFaaliyetKari/ucOncekiself.bilancoDonemiOrtalamaDolarKuru
            dortOncekiself.bilancoDonemiDolarFaaliyetKari = dortOncekiself.bilancoDonemiFaaliyetKari/dortOncekiself.bilancoDonemiOrtalamaDolarKuru
            besOncekiself.bilancoDonemiDolarFaaliyetKari = besOncekiself.bilancoDonemiFaaliyetKari/besOncekiself.bilancoDonemiOrtalamaDolarKuru
            altiOncekiself.bilancoDonemiDolarFaaliyetKari = altiOncekiself.bilancoDonemiFaaliyetKari/altiOncekiself.bilancoDonemiOrtalamaDolarKuru
            yediOncekiself.bilancoDonemiDolarFaaliyetKari = yediOncekiself.bilancoDonemiFaaliyetKari/yediOncekiself.bilancoDonemiOrtalamaDolarKuru

            self.bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(self.bilancoDonemiDolarFaaliyetKari).replace(",", ".")
            dortOncekiself.bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(dortOncekiself.bilancoDonemiDolarFaaliyetKari).replace(",",".")
            self.bilancoDonemiDolarFaaliyetKariDegisimi = self.bilancoDonemiDolarFaaliyetKari/dortOncekiself.bilancoDonemiDolarFaaliyetKari-1
            self.bilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(self.bilancoDonemiDolarFaaliyetKariDegisimi)

            birOncekibilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(birOncekibilancoDonemiDolarFaaliyetKari).replace(",", ".")
            besOncekiself.bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(besOncekiself.bilancoDonemiDolarFaaliyetKari).replace(",", ".")
            birOncekibilancoDonemiDolarFaaliyetKariDegisimi = birOncekibilancoDonemiDolarFaaliyetKari/besOncekiself.bilancoDonemiDolarFaaliyetKari-1
            birOncekibilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(birOncekibilancoDonemiDolarFaaliyetKariDegisimi)

            ikiOncekiself.bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(ikiOncekiself.bilancoDonemiDolarFaaliyetKari).replace(",", ".")
            altiOncekiself.bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(altiOncekiself.bilancoDonemiDolarFaaliyetKari).replace(",", ".")
            ikiOncekiself.bilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(ikiOncekiself.bilancoDonemiDolarFaaliyetKari/altiOncekiself.bilancoDonemiDolarFaaliyetKari-1)

            ucOncekiself.bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(ucOncekiself.bilancoDonemiDolarFaaliyetKari).replace(",", ".")
            yediOncekiself.bilancoDonemiDolarFaaliyetKariPrint = "{:,.0f}".format(yediOncekiself.bilancoDonemiDolarFaaliyetKari).replace(",", ".")
            ucOncekiself.bilancoDonemiDolarFaaliyetKariDegisimiPrint = "{:.2%}".format(ucOncekiself.bilancoDonemiDolarFaaliyetKari/yediOncekiself.bilancoDonemiDolarFaaliyetKari-1)

            dolarFaaliyetKariTablosu = PrettyTable()
            dolarFaaliyetKariTablosu.field_names = ["ÇEYREK", "FAALİYET KARI (DOLAR)", "ÖNCEKİ YIL", "ÖNCEKİ YIL FAALİYET KARI (DOLAR)", "YÜZDE DEĞİŞİM"]
            dolarFaaliyetKariTablosu.align["FAALİYET KARI (DOLAR)"] = "r"
            dolarFaaliyetKariTablosu.align["ÖNCEKİ YIL FAALİYET KARI (DOLAR)"] = "r"
            dolarFaaliyetKariTablosu.align["YÜZDE DEĞİŞİM"] = "r"
            dolarFaaliyetKariTablosu.add_row([self.bilancoDonemi, self.bilancoDonemiDolarFaaliyetKariPrint, dortOncekiself.bilancoDonemi, dortOncekiself.bilancoDonemiDolarFaaliyetKariPrint, self.bilancoDonemiDolarFaaliyetKariDegisimiPrint])
            dolarFaaliyetKariTablosu.add_row([birOncekibilancoDonemi, birOncekibilancoDonemiDolarFaaliyetKariPrint, besOncekiself.bilancoDonemi,besOncekiself.bilancoDonemiDolarFaaliyetKariPrint, birOncekibilancoDonemiDolarFaaliyetKariDegisimiPrint])
            dolarFaaliyetKariTablosu.add_row([ikiOncekiself.bilancoDonemi, ikiOncekiself.bilancoDonemiDolarFaaliyetKariPrint, altiOncekiself.bilancoDonemi, altiOncekiself.bilancoDonemiDolarFaaliyetKariPrint, ikiOncekiself.bilancoDonemiDolarFaaliyetKariDegisimiPrint])
            dolarFaaliyetKariTablosu.add_row([ucOncekiself.bilancoDonemi, ucOncekiself.bilancoDonemiDolarFaaliyetKariPrint, yediOncekiself.bilancoDonemi, yediOncekiself.bilancoDonemiDolarFaaliyetKariPrint, ucOncekiself.bilancoDonemiDolarFaaliyetKariDegisimiPrint])
            self.my_logger.info (dolarFaaliyetKariTablosu)

            # Bilanço Dönem Faaliyet Kar Artış Kriteri (DOLAR)
            if ceyrekDegeriHesapla(netKarRow, self.bilancoDonemiColumn) < 0:
                self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Çeyrek Net Kar Negatif"
                self.my_logger.info (printText)

            elif (self.bilancoDonemiDolarFaaliyetKari < 0):
                self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = False
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Ceyrek Dolar Faaliyet Kari Negatif"
                self.my_logger.info (printText)

            elif (self.bilancoDonemiDolarFaaliyetKari > 0) and (dortOncekiself.bilancoDonemiDolarFaaliyetKari < 0):
                self.bilancoDonemiDolarFaaliyetKariArtisi = 0
                self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = True
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + str(self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu) + " Son Çeyrek Dolar Faaliyet Karı Negatiften Pozitife Geçmiş"
                self.my_logger.info (printText)

            else:
                self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = (self.bilancoDonemiDolarFaaliyetKariDegisimi > 0.15)
                printText = "Bilanço Dönemi (DOLAR) Faaliyet Karı Artışı: " + "{:.2%}".format(self.bilancoDonemiDolarFaaliyetKariDegisimi) + " >? 15% " + str(self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu)
                self.my_logger.info(printText)

            # Önceki Dönem Faaliyet Kar Artış Kriteri (DOLAR)

            if self.bilancoDonemiDolarFaaliyetKariDegisimi >= 1:
                birOncekibilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = True
                printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: Bilanço Dönemi Artış " + "{:.2%}".format(self.bilancoDonemiDolarFaaliyetKariDegisimi) + \
                            " > 100%, Karşılaştırma Yapılmayacak: " + str(birOncekibilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)
                self.my_logger.info(printText)


            else:
                birOncekibilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = (birOncekibilancoDonemiDolarFaaliyetKariDegisimi < self.bilancoDonemiDolarFaaliyetKariDegisimi)
                printText = "Önceki Bilanço Dönemi (DOLAR) Faaliyet Kar Artışı: " + "{:.2%}".format(birOncekibilancoDonemiDolarFaaliyetKariDegisimi) + \
                            " <? " + "{:.2%}".format(self.bilancoDonemiDolarFaaliyetKariDegisimi) + " " + str(birOncekibilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu)
                self.my_logger.info(printText)


            self.my_logger.debug("")
            self.my_logger.debug("")
            self.my_logger.debug("----------------RAPOR DOSYASI OLUŞTURMA/GÜNCELLEME-------------------------------------")

            self.my_logger.debug (hisseAdi)

            excelRow = ExcelRowClass()

            excelRow.self.bilancoDonemiHasilat = self.bilancoDonemiHasilat
            excelRow.oncekiYilAyniCeyrekHasilat = dortOncekiself.bilancoDonemiHasilat
            excelRow.self.bilancoDonemiHasilatDegisimi = self.bilancoDonemiHasilatDegisimi
            excelRow.birOncekibilancoDonemiHasilat = birOncekibilancoDonemiHasilat
            excelRow.besOncekiself.bilancoDonemiHasilat = besOncekiself.bilancoDonemiHasilat
            excelRow.birOncekibilancoDonemiHasilatDegisimi = birOncekibilancoDonemiHasilatDegisimi
            excelRow.self.bilancoDonemiHasilatGelirArtisiGecmeDurumu = self.bilancoDonemiHasilatGelirArtisiGecmeDurumu
            excelRow.oncekiself.bilancoDonemiHasilatGelirArtisiGecmeDurumu = oncekiDonemHasilatGelirArtisiGecmeDurumu
            excelRow.self.bilancoDonemiFaaliyetKari = self.bilancoDonemiFaaliyetKari
            excelRow.oncekiYilAyniCeyrekFaaliyetKari = dortOncekiself.bilancoDonemiFaaliyetKari
            excelRow.self.bilancoDonemiFaaliyetKariDegisimi = self.bilancoDonemiFaaliyetKariDegisimi
            excelRow.birOncekibilancoDonemiFaaliyetKari = birOncekibilancoDonemiFaaliyetKari
            excelRow.besOncekiself.bilancoDonemiFaaliyetKari = besOncekiself.bilancoDonemiFaaliyetKari
            excelRow.oncekiself.bilancoDonemiFaaliyetKariDegisimi = birOncekibilancoDonemiFaaliyetKariDegisimi
            excelRow.self.bilancoDonemiFaaliyetKariArtisiGecmeDurumu = self.bilancoDonemiFaaliyetKariArtisiGecmeDurumu
            excelRow.oncekiself.bilancoDonemiFaaliyetKarArtisiGecmeDurumu = oncekiCeyrekFaaliyetKarArtisiGecmeDurumu

            excelRow.self.bilancoDonemiOrtalamaDolarKuru = self.bilancoDonemiOrtalamaDolarKuru
            excelRow.self.bilancoDonemiDolarHasilat = self.bilancoDonemiDolarHasilat
            excelRow.oncekiYilAyniCeyrekDolarHasilat = oncekiYilAyniCeyrekDolarHasilat
            excelRow.self.bilancoDonemiDolarHasilatDegisimi = self.bilancoDonemiDolarHasilatDegisimi
            excelRow.birOncekibilancoDonemiDolarHasilatDegisimi = birOncekibilancoDonemiDolarHasilatDegisimi
            excelRow.self.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = self.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu
            excelRow.oncekiself.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu = oncekiself.bilancoDonemiDolarHasilatGelirArtisiGecmeDurumu
            excelRow.self.bilancoDonemiDolarFaaliyetKari = self.bilancoDonemiDolarFaaliyetKari
            excelRow.dortOncekiself.bilancoDonemiDolarFaaliyetKari = dortOncekiself.bilancoDonemiDolarFaaliyetKari
            excelRow.self.bilancoDonemiDolarFaaliyetKariDegisimi = self.bilancoDonemiDolarFaaliyetKariDegisimi
            excelRow.birOncekibilancoDonemiDolarFaaliyetKariDegisimi = birOncekibilancoDonemiDolarFaaliyetKariDegisimi
            excelRow.self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu = self.bilancoDonemiDolarFaaliyetKariArtisiGecmeDurumu
            excelRow.oncekiself.bilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu = birOncekibilancoDonemiDolarFaaliyetKarArtisiGecmeDurumu

            excelRow.sermaye = sermaye
            excelRow.anaOrtaklikPayi = anaOrtaklikPayi
            excelRow.sonDortself.bilancoDonemiHasilatToplami = sonDortCeyrekHasilatToplami
            excelRow.onumuzdekiDortself.bilancoDonemiHasilatTahmini = onumuzdekiDortCeyrekHasilatTahmini
            excelRow.onumuzdekiDortself.bilancoDonemiFaaliyetKarMarjiTahmini = onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
            excelRow.faaliyetKariTahmini1 = faaliyetKariTahmini1
            excelRow.faaliyetKariTahmini2 = faaliyetKariTahmini2
            excelRow.ortalamaFaaliyetKariTahmini = ortalamaFaaliyetKariTahmini
            excelRow.hisseBasinaOrtalamaKarTahmini = hisseBasinaOrtalamaKarTahmini
            excelRow.bilancoEtkisi = bilancoEtkisi
            excelRow.bilancoTarihiself.hisseFiyati = self.hisseFiyati
            excelRow.gercekHisseDegeri = gercekDeger
            excelRow.targetBuy = targetBuy
            excelRow.gercekFiyataUzaklik = gercekFiyataUzaklik
            excelRow.fkOrani = fkOrani
            excelRow.hbkOrani = hbkOrani

            excelRow.netProKriteri = netProKriteri
            excelRow.forwardPeKriteri = forwardPeKriteri

            exportReportExcel(hisseAdi, reportFile, self.bilancoDonemi, excelRow)

            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info("")
            self.my_logger.info ("-------------------------------- %s ------------------------", hisseAdi)

            self.my_logger.removeHandler(output_file_handler)
            self.my_logger.removeHandler(stdout_handler)

        hasilat_hesaplari(0)
        faaliyet_kari_hesaplari(0)
        net_kar_hesaplari(0)
        brut_kar_hesaplari(0)
        gercek_deger_hesabi(0)
        netpro_kriteri_hesabi(0)
        bilanco_donemi_dolar_hesabi(0)
