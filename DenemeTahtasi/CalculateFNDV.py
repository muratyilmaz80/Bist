import xlrd
from GetGuncelHisseDegeri import returnGuncelHisseDegeri
import os

hisseAdi = "GSDHO"
bilancoDonemi = 202206
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

    bilancoDonemiColumn = donemColumnFind(varBilancoDonemi)

    def getBilancoDegeri(label, column):
        for rowi in range(sheet.nrows):
            cell = sheet.cell(rowi, 0)
            if cell.value == label:
                if sheet.cell_value(rowi, column)=="":
                    # print (label + " :Bilanço alanı boş!")
                    return 0
                else:
                    return sheet.cell_value(rowi, column)
        # print ("Uygun bilanco degeri bulunamadi: %s", label)
        return 0


    toplamDonenVarliklar = getBilancoDegeri("TOPLAM DÖNEN VARLIKLAR", bilancoDonemiColumn)
    kisaVadeliBorclanmalar = getBilancoDegeri("Kısa Vadeli Borçlanmalar", bilancoDonemiColumn)
    uzunVadeliBorcKisaVade = getBilancoDegeri("Uzun Vadeli Borçlanmaların Kısa Vadeli Kısımları", bilancoDonemiColumn)
    sermaye = getBilancoDegeri("Ödenmiş Sermaye", bilancoDonemiColumn)

    hisseBasinaNetDonenVarlik = (toplamDonenVarliklar - kisaVadeliBorclanmalar - uzunVadeliBorcKisaVade)/sermaye
    varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
    fndv = varHisseFiyati / hisseBasinaNetDonenVarlik

    if (fndv < 1):
        print (varHisseAdi, " F/NDV Oranı 1'in Altında: %s", "{:,.2f}".format(fndv))

# hesapla(hisseAdi, bilancoDonemi)

directory = "//Users//myilmaz//Documents//bist//bilancolar"

files_no_ext = [".".join(f.split(".")[:-1]) for f in os.listdir(directory)]
sorted_files = sorted(files_no_ext)
sorted_files.remove("")

for x in sorted_files:
    hesapla(x, bilancoDonemi)
