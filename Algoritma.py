import xlrd

loc = ("D:\\bist\\bilancolar\\DESPC.xlsx")

ceyrek = 202003

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

hesaplanacakCeyrekColumn = 0;

for columni in range (sheet.ncols):
    cell = sheet.cell(0,columni)
    if cell.value == ceyrek:
        hesaplanacakCeyrekColumn = columni


hasilatRow = 0;
for rowi in range (sheet.nrows):
    cell = sheet.cell(rowi,0)
    if cell.value == "Hasılat":
        hasilatRow = rowi


faaliyetKariRow = 0;
for rowi in range (sheet.nrows):
    cell = sheet.cell(rowi,0)
    if cell.value == "ESAS FAALİYET KARI (ZARARI)":
        faaliyetKariRow = rowi



def ceyrekDegisimiHesapla (row, column):
    ceyrekDegeri = sheet.cell_value(row, column)
    oncekiCeyrekDegeri = sheet.cell_value(row,(column-4))
    degisimSonucu = ceyrekDegeri/oncekiCeyrekDegeri -1
    print(sheet.cell_value(0, column), sheet.cell_value(row, 0), ceyrekDegeri)
    print (sheet.cell_value(0, (column-4)) , sheet.cell_value(row, 0), oncekiCeyrekDegeri)
    print ("Degisim Miktarı: ", "{:.2%}".format(degisimSonucu))
    return degisimSonucu



# 1.kriter hesabi
print("---------------------------------------------------------------------------------")
print ("1.Kriter: Son ceyrek satisi onceki yil ayni ceyrege göre en az %10 fazla olacak")

sonCeyrekSatisi = sheet.cell_value(hasilatRow, hesaplanacakCeyrekColumn)
print ("Son Ceyrek Satisi:", "{:,}".format(sonCeyrekSatisi).replace(',','.'))

oncekiYilAyniCeyrekSatisi = sheet.cell_value(hasilatRow, (hesaplanacakCeyrekColumn - 4))
print ("Onceki Yil Ayni Ceyrek Satisi:", oncekiYilAyniCeyrekSatisi)

kriter1SatisGelirArtisi = sonCeyrekSatisi/oncekiYilAyniCeyrekSatisi - 1
kriter1GecmeDurumu = (kriter1SatisGelirArtisi > 0.1)
print ("Kriter1: Satis Geliri Artisi:", "{:.2%}".format(kriter1SatisGelirArtisi), kriter1GecmeDurumu)

print ("Fonksiyon sonucu: ", ceyrekDegisimiHesapla(hasilatRow, hesaplanacakCeyrekColumn))














print("---------------------------------------------------------------------------------")
# 2.kriter hesabi

print ("2.Kriter: Son ceyrek faaliyet kari onceki yil ayni ceyrege göre en az %15 fazla olacak")

sonCeyrekFaaliyetKari = sheet.cell_value(faaliyetKariRow, hesaplanacakCeyrekColumn)
print ("Son Ceyrek Faaliyet Kari:", sonCeyrekFaaliyetKari)

oncekiYilAyniCeyrekFaaliyetKari = sheet.cell_value(faaliyetKariRow, (hesaplanacakCeyrekColumn - 4))
print ("Onceki Yil Ayni Ceyrek Faaliyet Kari:", oncekiYilAyniCeyrekFaaliyetKari)

kriter2FaaliyetKariArtisi = sonCeyrekFaaliyetKari/oncekiYilAyniCeyrekFaaliyetKari - 1
kriter2GecmeDurumu = (kriter2FaaliyetKariArtisi > 0.15)
print ("Kriter2: Faaliyet Kari Artisi:", "{:.2%}".format(kriter2FaaliyetKariArtisi), kriter2GecmeDurumu)






print("---------------------------------------------------------------------------------")
print ("3.Kriter: Bir Onceki Ceyrek Satis Artis Yuzdesi Cari Donemden Dusuk Olmali")

birOncekiCeyrekSatisi = sheet.cell_value(hasilatRow, (hesaplanacakCeyrekColumn - 1)) - sheet.cell_value(hasilatRow, (hesaplanacakCeyrekColumn - 2))
print ("Bir Onceki Ceyrek Satisi:", birOncekiCeyrekSatisi)

ikiOncekiCeyrekSatisi = sheet.cell_value(hasilatRow, (hesaplanacakCeyrekColumn - 5)) - sheet.cell_value(hasilatRow, (hesaplanacakCeyrekColumn - 6))
print ("2 Onceki Ceyrek Satisi:", ikiOncekiCeyrekSatisi)

kriter3OncekiCeyrekArtisi = birOncekiCeyrekSatisi/ikiOncekiCeyrekSatisi - 1
kriter3GecmeDurumu = (kriter3OncekiCeyrekArtisi < kriter1SatisGelirArtisi)
print ("Kriter3: Onceki Ceyrek Satis Geliri Artisi:", "{:.2%}".format(kriter3OncekiCeyrekArtisi), kriter3GecmeDurumu)











# 4.kriter hesabi
print("---------------------------------------------------------------------------------")

print ("4.Kriter: Onceki Ceyrek Faaliyet Kari Artis Yuzdesi Cari Donemden Dusuk Olmali")

oncekiCeyrekFaaliyetKari = sheet.cell_value(faaliyetKariRow, (hesaplanacakCeyrekColumn - 1)) - sheet.cell_value(faaliyetKariRow, (hesaplanacakCeyrekColumn - 2))
print ("Onceki Ceyrek Faaliyet Kari:", oncekiCeyrekFaaliyetKari)

ikiOncekiCeyrekFaaliyetKari = sheet.cell_value(faaliyetKariRow, (hesaplanacakCeyrekColumn - 5)) - sheet.cell_value(faaliyetKariRow, (hesaplanacakCeyrekColumn - 6))
print ("2 Onceki Ceyrek Faaliyet Kari:", ikiOncekiCeyrekFaaliyetKari)

kriter4OncekiCeyrekFaaliyetKariArtisi = oncekiCeyrekFaaliyetKari/ikiOncekiCeyrekFaaliyetKari - 1
kriter4GecmeDurumu = (kriter4OncekiCeyrekFaaliyetKariArtisi < kriter2FaaliyetKariArtisi)
print ("Kriter4: Onceki Yila Gore Faaliyet Kari Artisi:", "{:.2%}".format(kriter4OncekiCeyrekFaaliyetKariArtisi), kriter4GecmeDurumu)
