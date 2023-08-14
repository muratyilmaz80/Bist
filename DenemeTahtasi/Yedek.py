# Gerçek Deger Hesaplama 2
my_logger.info("")
my_logger.info("")
my_logger.info("----------------GERÇEK DEĞER HESABI NFK--------------------------------------------")

sermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)
my_logger.info("Sermaye: %s TL", "{:,.0f}".format(sermaye).replace(",", "."))

anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları", bilancoDonemiColumn) / getBilancoDegeri(
    "DÖNEM KARI (ZARARI)", bilancoDonemiColumn)
my_logger.info("Ana Ortaklık Payı: %s", "{:.3f}".format(anaOrtaklikPayi))

sonCeyrekNfk = ceyrekDegeriHesapla(brutKarRow, bilancoDonemiColumn) + ceyrekDegeriHesapla(genelYonetimGiderleriRow,
                                                                                          bilancoDonemiColumn) + ceyrekDegeriHesapla(
    pazarlamaGiderleriRow, bilancoDonemiColumn) + ceyrekDegeriHesapla(argeGiderleriRow, bilancoDonemiColumn)
my_logger.info("Son Çeyrek NFK: %s TL", "{:,.0f}".format(sonCeyrekNfk).replace(",", "."))

sonCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, bilancoDonemi)
birOncekiCeyrekSatisArtisYuzdesi = oncekiYilAyniCeyrekDegisimiHesapla(hasilatRow, birOncekiBilancoDonemi)

sonDortCeyrekHasilatToplami = ucOncekiBilancoDonemiHasilat + ikiOncekiBilancoDonemiHasilat + birOncekiBilancoDonemiHasilat + bilancoDonemiHasilat

my_logger.info("Son 4 Çeyrek Hasılat Toplamı: %s TL",
               "{:,.0f}".format(sonDortCeyrekHasilatToplami).replace(",", "."))

onumuzdekiDortCeyrekHasilatTahmini = (
        (((sonCeyrekSatisArtisYuzdesi + birOncekiCeyrekSatisArtisYuzdesi) / 2) + 1) * sonDortCeyrekHasilatToplami)

# HASILAT TAHMININI MANUEL DEGISTIRMEK ICIN
# onumuzdekiDortCeyrekHasilatTahmini = 85000000000

my_logger.info("Önümüzdeki 4 Çeyrek Hasılat Tahmini: %s TL",
               "{:,.0f}".format(onumuzdekiDortCeyrekHasilatTahmini).replace(",", "."))

ucOncekibilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ucOncekibilancoDonemiColumn)
ikiOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, ikiOncekibilancoDonemiColumn)
birOncekiBilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, birOncekibilancoDonemiColumn)
bilancoDonemiFaaliyetKari = ceyrekDegeriHesapla(faaliyetKariRow, bilancoDonemiColumn)

onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) / (
        bilancoDonemiHasilat + birOncekiBilancoDonemiHasilat)
my_logger.info("Önümüzdeki 4 Çeyrek Faaliyet Kar Marjı Tahmini: %s ",
               "{:.2%}".format(onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini))

faaliyetKariTahmini1 = onumuzdekiDortCeyrekHasilatTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
my_logger.info("Faaliyet Kar Tahmini1: %s TL", "{:,.0f}".format(faaliyetKariTahmini1).replace(",", "."))

faaliyetKariTahmini2 = ((birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 2 * 0.3) + (
        bilancoDonemiFaaliyetKari * 4 * 0.5) + \
                       ((
                                ucOncekibilancoDonemiFaaliyetKari + ikiOncekiBilancoDonemiFaaliyetKari + birOncekiBilancoDonemiFaaliyetKari + bilancoDonemiFaaliyetKari) * 0.2)
my_logger.info("Faaliyet Kar Tahmini2: %s TL", "{:,.0f}".format(faaliyetKariTahmini2).replace(",", "."))

ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1 + faaliyetKariTahmini2) / 2
my_logger.info("Ortalama Faaliyet Kari Tahmini: %s TL",
               "{:,.0f}".format(ortalamaFaaliyetKariTahmini).replace(",", "."))

# print ("----MURAT-----")
#
# istiraklerdenGelenKarRow = getBilancoTitleRow("Özkaynak Yöntemiyle Değerlenen Yatırımların Karlarından (Zararlarından) Paylar")
# istiraklerdenGelenNetKarSonCeyrek = ceyrekDegeriHesapla(istiraklerdenGelenKarRow,bilancoDonemiColumn)
# print ("İştiraklerden Gelen Net Kar Son Çeyrek: ", "{:,.0f}".format(istiraklerdenGelenNetKarSonCeyrek).replace(",","."))
#
# istiraklerdenGelenNetKarYillik = ceyrekDegeriHesapla(istiraklerdenGelenKarRow,bilancoDonemiColumn) + ceyrekDegeriHesapla(istiraklerdenGelenKarRow,birOncekibilancoDonemiColumn) + ceyrekDegeriHesapla(istiraklerdenGelenKarRow,ikiOncekibilancoDonemiColumn) + ceyrekDegeriHesapla(istiraklerdenGelenKarRow,ucOncekibilancoDonemiColumn)
# print ("İştiraklerden Gelen Net Kar Yıllık: ", "{:,.0f}".format(istiraklerdenGelenNetKarYillik).replace(",","."))
#
# print("----MURAT-----")

hisseBasinaOrtalamaKarTahmini = ((ortalamaFaaliyetKariTahmini) * anaOrtaklikPayi) / sermaye
my_logger.info("Hisse Başına Ortalama Kar Tahmini: %s TL", format(hisseBasinaOrtalamaKarTahmini, ".2f"))

likidasyonDegeri = likidasyonDegeriHesapla(bilancoDonemi)
my_logger.info("Likidasyon Değeri: %s TL", "{:,.0f}".format(likidasyonDegeri).replace(",", "."))

borclar = int(getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", bilancoDonemiColumn))
my_logger.info("Borçlar: %s TL", "{:,.0f}".format(borclar).replace(",", "."))

bilancoEtkisi = (likidasyonDegeri - borclar) / sermaye * anaOrtaklikPayi
my_logger.info("Bilanço Etkisi: %s TL", format(bilancoEtkisi, ".2f"))

gercekDeger = (hisseBasinaOrtalamaKarTahmini * 7) + bilancoEtkisi
my_logger.info("Gerçek Hisse Değeri: %s TL", format(gercekDeger, ".2f"))

targetBuy = gercekDeger * 0.66
my_logger.info("Target Buy: %s TL", format(targetBuy, ".2f"))

my_logger.info("Bilanço Tarihindeki Hisse Fiyatı: %s TL", format(hisseFiyati, ".2f"))

gercekFiyataUzaklik = hisseFiyati / targetBuy
my_logger.info("Gerçek Fiyata Uzaklık Oranı: %s", "{:.2%}".format(gercekFiyataUzaklik))

gercekFiyataUzaklikTl = hisseFiyati - targetBuy
my_logger.info("Gerçek Fiyata Uzaklık %s TL:", format(gercekFiyataUzaklikTl, ".2f"))







