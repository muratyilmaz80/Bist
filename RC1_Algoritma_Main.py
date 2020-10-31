import xlrd
from ExcelRowClass import ExcelRowClass
from Rapor_Olustur import exportReportExcel
from RC1_Algoritma_3Aylik_Tekli import runAlgoritma
from RC1_GetBondYield import returnBondYield
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri

varHisseAdi = "ARCLK"

varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"

varBilancoDonemi = 202009
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//RC1_Report_Deneme.xls"

runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile)