import os

import xlrd
from ExcelRowClass import ExcelRowClass
from Rapor_Olustur import exportReportExcel
from RC1_Algoritma_3Aylik import runAlgoritma
from RC1_GetBondYield import returnBondYield
from RC1_GetGuncelHisseDegeri import returnGuncelHisseDegeri


varBilancoDonemi = 202009
varBondYield = returnBondYield()
varReportFile = "//Users//myilmaz//Documents//bist//RC1_Report_Deneme_Multiple.xls"


def runAlgoritmaMultiple(string):
    content = string
    content = content.strip()
    contentList = content.split("-")
    contentList = [x for x in contentList]
    print(contentList)

    for varHisseAdi in contentList:

        varBilancoDosyasi = "//Users//myilmaz//Documents//bist//bilancolar//" + varHisseAdi + ".xlsx"
        varHisseFiyati = returnGuncelHisseDegeri(varHisseAdi)
        runAlgoritma(varBilancoDosyasi, varBilancoDonemi, varBondYield, varHisseFiyati, varReportFile)

runAlgoritmaMultiple("DOAS-ARSAN-ECILC-EGSER-RTALB-OYAKC-VESBE-GEDZA-DESPC-INDES")
runAlgoritmaMultiple("NUHCM-DEVA-CCOLA-BNTAS-GOODY-ULKER-CEMTS-IEYHO-EPLAS-GENTS")
runAlgoritmaMultiple("YATAS-ARCLK-BRKSN-BUCIM-SANKO-TATGD-CEMAS-FORMT-SISE-TOASO")
runAlgoritmaMultiple("TCELL-KAREL-TTRAK-DGATE-ARENA-EREGL-JANTS-AKSA-FROTO-AEFES")
runAlgoritmaMultiple("ENJSA-TMPOL-CIMSA-DYOBY-ALCAR-BRISA-AKCNS-HEKTS-PNSUT-TUKAS")
runAlgoritmaMultiple("KRONT-KNFRT-KUTPO-EGPRO-KLMSN-KARTN-KRSTL-DMSAS-KRTEK-PRKAB")
runAlgoritmaMultiple("SANFM-KAPLM-MPARK-SONME-BFREN-MRSHL-GEREL-BURVA-SASA-TKNSA")
runAlgoritmaMultiple("ZOREN-DITAS-SILVR-PENGD-DGKLB")