import logging

from RC1_Algoritma_3Aylik import runAlgoritma
from RC1_GetBondYield import returnBondYield
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri


varHisseAdi = "ASELS"

varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"

varBilancoDonemi = 202009
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//RC1_Report_Yeni_Deneme.xls"
varLogLevel = logging.DEBUG

runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogLevel)



