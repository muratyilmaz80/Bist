import logging

from RC1_Algoritma_3Aylik import runAlgoritma
from RC1_Algoritma_6Aylik import runAlgoritma6Aylik
from RC1_GetBondYield import returnBondYield
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri


varHisseAdi = "MEGAP"

varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"

varBilancoDonemi = 202006
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//RC1_Report_Yeni_Deneme_6Ay.xls"
varLogLevel = logging.DEBUG
varLogPath = "//Users//myilmaz//Documents//bist//log//6aylik//"

# runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)

runAlgoritma6Aylik(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)


