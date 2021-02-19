import logging

from RC1_Algoritma_3Aylik import runAlgoritma
from RC1_Algoritma_6Aylik import runAlgoritma6Aylik
from RC1_GetBondYield import returnBondYield
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri


varHisseAdi = "BRISA"

varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"

varBilancoDonemi = 202012
varBondYield = returnBondYield()
varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
varReportFile = "//Users//myilmaz//Documents//bist//RC1_Report_202012_3Aylik.xls"
varReportFile6Aylik = "//Users//myilmaz//Documents//bist//RC1_Report_202012_6Aylik.xls"
varLogLevel = logging.DEBUG
varLogPath = "//Users//myilmaz//Documents//bist//log//2020_12//"



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



# runAlgoritmaMultiple("DOAS-ARSAN-ECILC-EGSER-RTALB-OYAKC-VESBE-GEDZA-DESPC-INDES")
# runAlgoritmaMultiple("NUHCM-DEVA-CCOLA-BNTAS-GOODY-ULKER-CEMTS-IEYHO-EPLAS-GENTS")
# runAlgoritmaMultiple("YATAS-ARCLK-BRKSN-BUCIM-SANKO-TATGD-CEMAS-FORMT-SISE-TOASO")
# runAlgoritmaMultiple("TCELL-KAREL-TTRAK-DGATE-ARENA-EREGL-JANTS-AKSA-FROTO-AEFES")
# runAlgoritmaMultiple("ENJSA-TMPOL-CIMSA-DYOBY-ALCAR-BRISA-AKCNS-HEKTS-PNSUT-TUKAS")
# runAlgoritmaMultiple("KRONT-KNFRT-KUTPO-EGPRO-KLMSN-KARTN-KRSTL-DMSAS-KRTEK-PRKAB")
# runAlgoritmaMultiple("SANFM-KAPLM-MPARK-SONME-BFREN-MRSHL-GEREL-BURVA-SASA-TKNSA")
# runAlgoritmaMultiple("ZOREN-DITAS-SILVR-PENGD-DGKLB")

# 6 Aylık İçin
# runAlgoritmaMultiple("VANGD-UZERB-KSTUR-ORMA-SODSN-SUMAS-TKURU-YBTAS-MERIT-OZRDN")
# runAlgoritmaMultiple("MEGAP-SEYKM-SAFKR-ISBIR-AYES-IZFAS-IZTAR-RODRG-YONGA-BASCM-BALAT")