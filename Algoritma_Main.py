import logging
from Algoritma_3Aylik_Yeni import Algoritma
from GetBondYield import returnBondYield
from GetGuncelHisseDegeri import returnGuncelHisseDegeri

varHisseAdi =("BOSSA")


varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + varHisseAdi + ".xlsx"
varBilancoDonemi = 202503
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//Report_202503.xls"
varLogLevel = logging.INFO
varLogPath = "//Users//myilmaz//Documents//bist//log//2025_03//"


def runAlgoritmaSingle():
    algoritma = Algoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile, varLogPath, varLogLevel)
    algoritma.runAlgoritma()


def runAlgoritmaMultiple(string):
    content = string
    content = content.strip()
    contentList = content.split("-")
    contentList = [x for x in contentList]
    print(contentList)

    for varHisseAdi in contentList:
        varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + varHisseAdi + ".xlsx"
        varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
        try:
            algoritma = Algoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile,
                                  varLogPath, varLogLevel)
            algoritma.runAlgoritma()
        except:
            print ("HATA")

runAlgoritmaSingle()

#algoritma.runAlgoritma()

# runAlgoritmaMultiple("PLTUR")

