import logging
from Algoritma_3Aylik_Yeni import Algoritma
from GetBondYield import returnBondYield
from GetGuncelHisseDegeri import returnGuncelHisseDegeri

varHisseAdi ="SISE"

varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + varHisseAdi + ".xlsx"
varBilancoDonemi = 202306
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//Report_202306_3Aylik.xls"
varLogLevel = logging.DEBUG
varLogPath = "//Users//myilmaz//Documents//bist//log//2023_06//"

algoritma = Algoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)


def runAlgoritmaMultiple(string):
    content = string
    content = content.strip()
    contentList = content.split("-")
    contentList = [x for x in contentList]
    print(contentList)

    for varHisseAdi in contentList:

        varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + varHisseAdi + ".xlsx"
        varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
        algoritma.runAlgoritma()


algoritma.runAlgoritma()

# runAlgoritma6Aylik(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile6Aylik, varLogPath, varLogLevel)

# runAlgoritmaMultiple("OZGYO-ISGSY")

# 6 Aylık İçin
# runAlgoritmaMultiple("VANGD-UZERB-KSTUR-ORMA-SODSN-SUMAS-TKURU-YBTAS-MERIT-OZRDN")

