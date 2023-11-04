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

puanlama_dosyasi = "//Users//myilmaz//Documents//bist//puanlama_listesi.xls"
puanlama_listesi_df = pd.read_excel(puanlama_dosyasi, index_col=0)

for index, row in puanlama_listesi_df.iterrows():
    puanlama_listesi_df.at[index, "PUAN"] = 0
    puanlama_listesi_df.at[index, "Weighted PUAN"] = 0

#tl_puanlama_listesi_df = puanlama_listesi_df.query("tr1  == True & tr2 == True & tr3 == True & tr4 == True")
#usd_puanlama_listesi_df = puanlama_listesi_df.query("dlr1  == True & dlr2 == True & dlr3 == True & dlr4 == True")

#print (f"Tum Liste Boyutu: {len(puanlama_listesi_df)}")
#print (f"TL'den Geçen Liste Boyutu:  {len(tl_puanlama_listesi_df)}")
#print (f"Dolar'dan Geçen Liste Boyutu {len(usd_puanlama_listesi_df)}")


def sort_and_grade_column(column_title, type):
    global puanlama_listesi_df

    if type == "Ascending":
        puanlama_listesi_df = puanlama_listesi_df.sort_values(by=[column_title], ascending=True)
    elif type == "Descending":
        puanlama_listesi_df = puanlama_listesi_df.sort_values(by=[column_title], ascending=False)

    puan = 0
    for index, row in puanlama_listesi_df.iterrows():
        puan = puan + 1
        mevcut_puan = puanlama_listesi_df.at[index, "PUAN"]
        puanlama_listesi_df.at[index, "PUAN"] = mevcut_puan + puan


sort_and_grade_column("Net Kar Büyüme Yıllık", "Ascending")
sort_and_grade_column("Net Kar Büyüme 4 Önceki Çeyreğe Göre", "Ascending")
sort_and_grade_column("Esas Faaliyet Karı Büyüme Yıllık", "Ascending")
sort_and_grade_column("Hasılat Büyüme Yıllık", "Ascending")
sort_and_grade_column("FAVÖK Büyüme Yıllık", "Ascending")
sort_and_grade_column("F/K", "Descending")
sort_and_grade_column("Nakit/PD", "Ascending")
sort_and_grade_column("Nakit/FD", "Ascending")
sort_and_grade_column("PD/DD", "Descending")
sort_and_grade_column("PEG", "Descending")
sort_and_grade_column("FD/Satışlar", "Descending")
sort_and_grade_column("FD/FAVÖK", "Descending")
sort_and_grade_column("PD/EFK", "Descending")
sort_and_grade_column("Cari Oran", "Ascending")
sort_and_grade_column("Likit Oranı", "Ascending")
sort_and_grade_column("Nakit Oranı", "Ascending")
sort_and_grade_column("Asit Test Oranı", "Ascending")
sort_and_grade_column("ROE (Özsermaye Karlılığı)", "Ascending")
sort_and_grade_column("ROA (Aktif Karlılık)", "Ascending")
sort_and_grade_column("Yıllık Net Kar Marjı", "Ascending")
sort_and_grade_column("Son Çeyrek Net Kar Marjı", "Ascending")
sort_and_grade_column("Aktif Devir Hızı", "Ascending")
sort_and_grade_column("Borç/Kaynak", "Descending")
sort_and_grade_column("Özsermaye Büyümesi", "Ascending")



for index, row in puanlama_listesi_df.iterrows():
    gercek_fiyata_uzaklik = puanlama_listesi_df.at[index, "Gerçek Fiyata Uzaklık"]
    gercek_fiyata_uzaklik_nfk = puanlama_listesi_df.at[index, "Gerçek Fiyata Uzaklık NFK"]
    gercek_fiyata_uzaklik_ort = (gercek_fiyata_uzaklik + gercek_fiyata_uzaklik_nfk) / 2
    puan = puanlama_listesi_df.at[index, "PUAN"]
    weighted_puan = puan / gercek_fiyata_uzaklik_ort
    puanlama_listesi_df.at[index, "Weighted PUAN"] = weighted_puan

# for index, row in puanlama_listesi_df.iterrows():
#     print (row["Weighted PUAN"])


puanlama_listesi_df = puanlama_listesi_df.sort_values(by=["Weighted PUAN"], ascending=False)

puanlama_listesi_df.to_excel("//Users//myilmaz//Documents//bist//puanlama_listesi_out.xls")
