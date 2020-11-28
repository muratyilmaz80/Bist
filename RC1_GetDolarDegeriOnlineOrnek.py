#!/usr/bin/python

from RC1_GetDolarDegeriOnline import DovizKurlari
# ornek isimli nesne yaratiliyor.
ornek = DovizKurlari()
print ("Bugunun Kurlar : ")
print ("EURO DEGERI="+ornek.DegerSor("EUR","ForexBuying"))
# deger sor fonksiyonu ile USD'nin degeri sorgulaniyor.
Dolar_Deger = ornek.DegerSor("USD","ForexBuying")
print ("DOLAR DEGERI="+Dolar_Deger)

# Arsiv konusu
print ("\nArsiv'deki bir degere bakalim. 02.02.2015")
print ("EURO DEGERI="+ornek.Arsiv(2,2,2015,"EUR","ForexBuying"))
# 02.02.2015 tarihindeki USD'in degeri sorgulaniyor.
Dolar_Deger = ornek.Arsiv(2,2,2015,"USD","ForexBuying")
print ("DOLAR DEGERI="+Dolar_Deger)

print ("\nArsiv'deki bir degere bakalim. 19.02.2018")
print ("EURO DEGERI="+ornek.Arsiv_tarih("19.02.2018","USD","ForexBuying"))
Dolar_Deger = ornek.Arsiv_tarih("19.02.2018","USD","ForexBuying")
print ("DOLAR DEGERI="+Dolar_Deger)




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