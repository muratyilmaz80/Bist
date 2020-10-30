import xlrd
from ExcelRowClass import ExcelRowClass
from Rapor_Olustur import exportReportExcel
from RC1_Algoritma_3Aylik_Tekli import runAlgoritma

varBilancoDosyasi = ("//Users//myilmaz//Documents//bist//bilancolar//ARCLK.xlsx")

varBilancoDonemi = 202009
varBondYield = 0.1448
varHisseFiyati = 25.74
varReportFile = "//Users//myilmaz//Documents//bist//RC1_Report_Deneme.xls"

runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile)