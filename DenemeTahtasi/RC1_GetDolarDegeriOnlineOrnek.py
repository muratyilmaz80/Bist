#!/usr/bin/python

from GetDolarDegeriOnline import DovizKurlari
# ornek isimli nesne yaratiliyor.
ornek = DovizKurlari()
#print ("Bugunun Kurlar : ")
#print ("EURO DEGERI="+ornek.DegerSor("EUR","ForexBuying"))
# deger sor fonksiyonu ile USD'nin degeri sorgulaniyor.
# Dolar_Deger = ornek.DegerSor("USD","ForexBuying")
#print ("DOLAR DEGERI="+Dolar_Deger)

# Arsiv konusu
#print ("\nArsiv'deki bir degere bakalim. 02.02.2015")
#print ("EURO DEGERI="+ornek.Arsiv(2,2,2015,"EUR","ForexBuying"))
# 02.02.2015 tarihindeki USD'in degeri sorgulaniyor.
#Dolar_Deger = ornek.Arsiv(2,2,2015,"USD","ForexBuying")
#print ("DOLAR DEGERI="+Dolar_Deger)

#print ("\nArsiv'deki bir degere bakalim. 19.02.2018")
#print ("EURO DEGERI="+ornek.Arsiv_tarih("19.02.2018","USD","ForexBuying"))
# Dolar_Deger = ornek.Arsiv_tarih("01.12.2004","USD","ForexBuying")


print ("01 ->" + ornek.Arsiv_tarih("04.01.2022","USD","ForexBuying"))
print ("02 ->" + ornek.Arsiv_tarih("01.02.2021","USD","ForexBuying"))
print ("03 ->" + ornek.Arsiv_tarih("01.03.2021","USD","ForexBuying"))
print ("04 ->" + ornek.Arsiv_tarih("01.04.2021","USD","ForexBuying"))
print ("05 ->" + ornek.Arsiv_tarih("03.05.2021","USD","ForexBuying"))
print ("06 ->" + ornek.Arsiv_tarih("01.06.2021","USD","ForexBuying"))
print ("07 ->" + ornek.Arsiv_tarih("01.07.2021","USD","ForexBuying"))
print ("08 ->" + ornek.Arsiv_tarih("03.08.2021","USD","ForexBuying"))
print ("09 ->" + ornek.Arsiv_tarih("01.09.2021","USD","ForexBuying"))
print ("10 ->" + ornek.Arsiv_tarih("01.10.2021","USD","ForexBuying"))
print ("11 ->" + ornek.Arsiv_tarih("01.11.2021","USD","ForexBuying"))
print ("12 ->" + ornek.Arsiv_tarih("01.12.2021","USD","ForexBuying"))




#Kendi kodunuzda kullanmak için DovizKurlari.py dosyasını kendi projenizin klasörünüze taşıyın.
#Sonrasında DovizKurları Nesnesi yaratarak, DegerSor fonksiyonu ile istediğiniz değeri sistemden çekebilirsiniz.
#"DegerSor" fonksiyonu iki parametre alır.
#DegerSor (Parametre1, Parametre2)
#* Parametre verilmezse JSON olarak tüm veriler döner
#Parametre1 = USD, EUR, AUD gibi para cinsinin resmi kısaltmaları
#Parametre2 = Almak istediğiniz değer ;

#"BanknoteBuying"    : Alış Değeri
#"BanknoteSelling"   : Satış Değeri
#"CrossRateUSD"      : USD ile çapraz kur
#"CurrencyName"      : Resmi Adı
#"ForexBuying"       : Forex Alış
#"ForexSelling"      : Forex Satış
#"Kod"               : Kodu
#"Unit"              : 1
#"isim"              : Türkçe Adı

#Arşivden veri çekmek
#Eski bir tarihteki kur'u ögrenmek için Arsiv veya Arsiv_Tarih fonksiyonlarını kullanabilirsiniz.
#Arsiv ( Gun, Ay, Yil,Parametre1, Parametre2)
#Gun, Ay, Yil = integer veya string olabilir.
#Arsiv_Tarih (Tarih,Parametre1, Parametre2)
#Tarih = "01.02.2015" Şeklinde bir string veri olmalıdır.