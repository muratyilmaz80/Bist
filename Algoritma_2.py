import xlrd

bilancoDosyasi = ("D:\\bist\\bilancolar\\DESPC.xlsx")
hesaplanacakCeyrek = 202003

wb = xlrd.open_workbook(bilancoDosyasi)
sheet = wb.sheet_by_index(0)

def ceyrekColumnFind(col):
    global hesaplanacakCeyrekColumn
    for columni in range(sheet.ncols):
        cell = sheet.cell(0, columni)
        if cell.value == col:
            # print ("Uygun Ceyrek Var, Column: ", columni)
            return columni
    print ("Uygun Ceyrek Bulunamadi!!!")
    return -1

hesaplanacakCeyrekColumn = ceyrekColumnFind(hesaplanacakCeyrek)


def getBilancoDegeri(label, ceyrek):
    ceyrekColumn = ceyrekColumnFind(ceyrek)
    for rowi in range(sheet.nrows):
        cell = sheet.cell(rowi, 0)
        if cell.value == label:
            return sheet.cell_value(rowi, ceyrekColumn)
    print ("Uygun bilanco degeri bulunamadi!")
    return -1


def getBilancoTitleRow (title):
    for rowi in range(sheet.nrows):
        cell = sheet.cell(rowi, 0)
        if cell.value == title:
            return rowi
    print ("Uygun baslik bulunamadi!")
    return -1


hasilatRow = getBilancoTitleRow("Hasılat")
faaliyetKariRow = getBilancoTitleRow("ESAS FAALİYET KARI (ZARARI)");


def ceyrekDegeriHesapla (r,c):
    lastDigits = abs(sheet.cell_value(0,c)) % (10 ** 2)
    if (lastDigits == 3):
        return sheet.cell_value(r,c)
    else:
        return (sheet.cell_value(r,c) - sheet.cell_value(r,(c-1)))


def ceyrekDegisimiHesapla (row, column):
    ceyrekDegeri = ceyrekDegeriHesapla(row, column)
    oncekiCeyrekDegeri = ceyrekDegeriHesapla(row,(column-4))
    degisimSonucu = ceyrekDegeri/oncekiCeyrekDegeri -1
    print(sheet.cell_value(0, column), sheet.cell_value(row, 0), ceyrekDegeri)
    print (sheet.cell_value(0, (column-4)) , sheet.cell_value(row, 0), oncekiCeyrekDegeri)
    return degisimSonucu

def convertYearQuarter (a):
    y = int (a/100)
    q = int (a % 100)
    return (y, q)
#a,b = convertYearQuarter(sheet.cell_value(0,4))


def likidasyonDegeriHesapla(ceyrek):
    nakit = getBilancoDegeri("Nakit ve Nakit Benzerleri", ceyrek)
    alacaklar = getBilancoDegeri("Ticari Alacaklar", ceyrek)
    stoklar = getBilancoDegeri("Stoklar", ceyrek)
    digerVarliklar = getBilancoDegeri("Diğer Dönen Varlıklar", ceyrek)
    finansalVarliklar = getBilancoDegeri("Finansal Yatırımlar", ceyrek)
    maddiDuranVarliklar = getBilancoDegeri("Maddi Duran Varlıklar", ceyrek)
    likidasyonDegeri = nakit + (alacaklar*0.7)+(stoklar*0.5)+(digerVarliklar*0.7)+(finansalVarliklar*0.7)+(maddiDuranVarliklar*0.2)
    return likidasyonDegeri


# 1.kriter hesabi
print("---------------------------------------------------------------------------------")
print ("1.Kriter: Son ceyrek satisi onceki yil ayni ceyrege göre en az %10 fazla olacak")

kriter1SatisGelirArtisi = ceyrekDegisimiHesapla(hasilatRow, hesaplanacakCeyrekColumn)
kriter1GecmeDurumu = (kriter1SatisGelirArtisi > 0.1)
print ("Kriter1: Satis Geliri Artisi:", "{:.2%}".format(kriter1SatisGelirArtisi), kriter1GecmeDurumu)


# 2.kriter hesabi
print("---------------------------------------------------------------------------------")
print ("2.Kriter: Son ceyrek faaliyet kari onceki yil ayni ceyrege göre en az %15 fazla olacak")

kriter2FaaliyetKariArtisi = ceyrekDegisimiHesapla(faaliyetKariRow, (hesaplanacakCeyrekColumn))
kriter2GecmeDurumu = (kriter2FaaliyetKariArtisi > 0.15)
print ("Kriter2: Faaliyet Kari Artisi:", "{:.2%}".format(kriter2FaaliyetKariArtisi), kriter2GecmeDurumu)


# 3.kriter hesabı
print("---------------------------------------------------------------------------------")
print ("3.Kriter: Bir Onceki Ceyrek Satis Artis Yuzdesi Cari Donemden Dusuk Olmali")

kriter3OncekiCeyrekArtisi = ceyrekDegisimiHesapla(hasilatRow, (hesaplanacakCeyrekColumn - 1))
kriter3GecmeDurumu = (kriter3OncekiCeyrekArtisi < kriter1SatisGelirArtisi)
print ("Kriter3: Onceki Ceyrek Satis Geliri Artisi:", "{:.2%}".format(kriter3OncekiCeyrekArtisi), kriter3GecmeDurumu)


# 4.kriter hesabi
print("---------------------------------------------------------------------------------")
print ("4.Kriter: Bir Onceki Ceyrek Faaliyet Kar Artis Yuzdesi Cari Donemden Dusuk Olmali")
kriter4OncekiCeyrekFaaliyetKariArtisi = ceyrekDegisimiHesapla(faaliyetKariRow, (hesaplanacakCeyrekColumn - 1))
kriter4GecmeDurumu = (kriter4OncekiCeyrekFaaliyetKariArtisi < kriter2FaaliyetKariArtisi)
print ("Kriter4: Onceki Yila Gore Faaliyet Kari Artisi:", "{:.2%}".format(kriter4OncekiCeyrekFaaliyetKariArtisi), kriter4GecmeDurumu)


# Gercek Deger Hesapla
print("---------------------------------------------------------------------------------")



sermaye = getBilancoDegeri("Ödenmiş Sermaye", hesaplanacakCeyrek)
print ("Sermaye:", sermaye)

anaOrtaklikPayi = getBilancoDegeri("Ana Ortaklık Payları" , hesaplanacakCeyrek) / getBilancoDegeri("DÖNEM KARI (ZARARI)" , hesaplanacakCeyrek)
print ("Ana Ortaklık Payı:", anaOrtaklikPayi)

sonCeyrekSatisArtisYuzdesi = ceyrekDegisimiHesapla(hasilatRow, hesaplanacakCeyrekColumn)
birOncekiCeyrekSatisArtisYuzdesi = ceyrekDegisimiHesapla(hasilatRow, (hesaplanacakCeyrekColumn-1))


ceyrek1Satis = ceyrekDegeriHesapla (hasilatRow, (hesaplanacakCeyrekColumn-3))
ceyrek2Satis = ceyrekDegeriHesapla (hasilatRow, (hesaplanacakCeyrekColumn-2))
ceyrek3Satis = ceyrekDegeriHesapla (hasilatRow, (hesaplanacakCeyrekColumn-1))
ceyrek4Satis = ceyrekDegeriHesapla (hasilatRow, hesaplanacakCeyrekColumn)

sonDortCeyrekSatisToplami = ceyrek1Satis + ceyrek2Satis + ceyrek3Satis + ceyrek4Satis
print ("Son 4 ceyrek satış toplamı:", sonDortCeyrekSatisToplami)

onumuzdekiDortCeyrekSatisTahmini = ((((sonCeyrekSatisArtisYuzdesi + birOncekiCeyrekSatisArtisYuzdesi)/2)+1)*sonDortCeyrekSatisToplami)
print ("Önümüzdeki 4 çeyrek satış tahmini:", onumuzdekiDortCeyrekSatisTahmini)

ceyrek1FaaliyetKari = ceyrekDegeriHesapla (faaliyetKariRow, (hesaplanacakCeyrekColumn-3))
ceyrek2FaaliyetKari = ceyrekDegeriHesapla (faaliyetKariRow, (hesaplanacakCeyrekColumn-2))
ceyrek3FaaliyetKari = ceyrekDegeriHesapla (faaliyetKariRow, (hesaplanacakCeyrekColumn-1))
ceyrek4FaaliyetKari = ceyrekDegeriHesapla (faaliyetKariRow, hesaplanacakCeyrekColumn)

onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini = (ceyrek3FaaliyetKari + ceyrek4FaaliyetKari) / (ceyrek4Satis + ceyrek3Satis)
print ("Önümüzdeki 4 çeyrek faaliyet kar marjı tahmini:", onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini)

faaliyetKariTahmini1 = onumuzdekiDortCeyrekSatisTahmini * onumuzdekiDortCeyrekFaaliyetKarMarjiTahmini
print ("Faaliyet Kar Tahmini1:", faaliyetKariTahmini1)

faaliyetKariTahmini2 = ((ceyrek3FaaliyetKari+ceyrek4FaaliyetKari)*2*0.3) + (ceyrek4FaaliyetKari*4*0.5) + \
                       ((ceyrek1FaaliyetKari+ceyrek2FaaliyetKari+ceyrek3FaaliyetKari+ceyrek4FaaliyetKari)*0.2)
print ("Faaliyet Kar Tahmini2:", faaliyetKariTahmini2)

ortalamaFaaliyetKariTahmini = (faaliyetKariTahmini1+faaliyetKariTahmini2)/2
print ("Ortalama Faaliyet Kari Tahmini:", ortalamaFaaliyetKariTahmini)

hisseBasinaOrtalamaKarTahmini = (ortalamaFaaliyetKariTahmini*anaOrtaklikPayi)/sermaye
print ("Hisse başına ortalama kar tahmini:", hisseBasinaOrtalamaKarTahmini)


likidasyonDegeri = likidasyonDegeriHesapla(hesaplanacakCeyrek)
print ("Likidasyon değeri:", likidasyonDegeri)

borclar = getBilancoDegeri("TOPLAM YÜKÜMLÜLÜKLER", hesaplanacakCeyrek)
print ("Borçlar:", borclar)

bilancoEtkisi = (likidasyonDegeri-borclar)/sermaye * anaOrtaklikPayi
print ("Bilanço Etkisi:", bilancoEtkisi)

gercekDeger = (hisseBasinaOrtalamaKarTahmini*7) + bilancoEtkisi
print("Gerçek hisse değeri:", gercekDeger)

targetBuy = gercekDeger*0.66
print ("Target buy:", targetBuy)
