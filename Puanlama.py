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
    print (row["PUAN"])

# puanlama_listesi_df.to_excel("//Users//myilmaz//Documents//bist//puanlama_listesi_out.xls")