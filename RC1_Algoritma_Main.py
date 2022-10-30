import logging

from RC1_Algoritma_3Aylik import runAlgoritma
from RC1_Algoritma_6Aylik import runAlgoritma6Aylik
from RC1_GetBondYield import returnBondYield
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri


varHisseAdi ="BRSAN"

varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"

varBilancoDonemi = 202209
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//Report_202209_3Aylik.xls"
varReportFile6Aylik = "//Users//myilmaz//Documents//bist//Report_202206_6Aylik.xls"
varLogLevel = logging.DEBUG
varLogPath = "//Users//myilmaz//Documents//bist//log//2022_09//"



def runAlgoritmaMultiple(string):
    content = string
    content = content.strip()
    contentList = content.split("-")
    contentList = [x for x in contentList]
    print(contentList)

    for varHisseAdi in contentList:

        varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"
        varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
        runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)



runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)

# runAlgoritma6Aylik(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile6Aylik, varLogPath, varLogLevel)

# runAlgoritmaMultiple("DOAS-EMNIS-ARCLK-KRDMD-VESBE-GOODY-PETKM-EREGL-TIRE-TTRAK-ISDMR-BUCIM-ENKAI-BRKSN")


# 6 Aylık İçin
# runAlgoritmaMultiple("VANGD-UZERB-KSTUR-ORMA-SODSN-SUMAS-TKURU-YBTAS-MERIT-OZRDN")
# runAlgoritmaMultiple("MEGAP-SEYKM-SAFKR-ISBIR-AYES-IZFAS-IZTAR-RODRG-YONGA-BASCM-BALAT")