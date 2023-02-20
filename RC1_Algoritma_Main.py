import logging

from RC1_Algoritma_3Aylik import runAlgoritma
from RC1_Algoritma_6Aylik import runAlgoritma6Aylik
from RC1_GetBondYield import returnBondYield
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri


varHisseAdi ="HLGYO"

varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + varHisseAdi + ".xlsx"

varBilancoDonemi = 202212
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//Report_202212_3Aylik.xls"
varReportFile6Aylik = "//Users//myilmaz//Documents//bist//Report_202212_6Aylik.xls"
varLogLevel = logging.DEBUG
varLogPath = "//Users//myilmaz//Documents//bist//log//2022_12//"



def runAlgoritmaMultiple(string):
    content = string
    content = content.strip()
    contentList = content.split("-")
    contentList = [x for x in contentList]
    print(contentList)

    for varHisseAdi in contentList:

        varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + varHisseAdi + ".xlsx"
        varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
        runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)



runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)

# runAlgoritma6Aylik(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile6Aylik, varLogPath, varLogLevel)

# runAlgoritmaMultiple("OZGYO-ISGSY")


# 6 Aylık İçin
# runAlgoritmaMultiple("VANGD-UZERB-KSTUR-ORMA-SODSN-SUMAS-TKURU-YBTAS-MERIT-OZRDN")
