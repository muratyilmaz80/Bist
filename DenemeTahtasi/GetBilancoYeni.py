import urllib

from bs4 import BeautifulSoup
import requests
import regex as re
import pandas as pd
from tabulate import tabulate
import os.path
import time
import json
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

def fnc(pg, bildiriNumber):
    try:
        currPage = requests.get(pg)
    except requests.ConnectionError:
        print("ConnectionError")
        return
    except requests.exceptions.Timeout:
        print("Timeout")
        return
    except requests.exceptions.TooManyRedirects:
        print("TooManyRedirects")
        return
    except requests.exceptions.RequestException as e:
        print("RequestException")
        return

    #currPage = requests.get(pg)
    soup = BeautifulSoup(currPage.text, 'html.parser')

    for part in soup.find_all('h1'):
        if re.search("Finansal Rapor.*", part.text):
            reportType = "Finansal Rapor"

            stockName = soup.find('div', {"class": "type-medium type-bold bi-sky-black"})
            stockCode = soup.find('div', {"class": "type-medium bi-dim-gray"})
            stockCode = stockCode.text.split(",")[0]
            year = ""
            period = ""
            gonderimTarihi = ""
            duzeltilmemisFiyat = 0
            sunumParaBirimiInt = 1

            print("Hisse Adı: ", stockName.get_text())
            # stockCode = stockCode[0]
            print("Hisse Kodu: ", stockCode)
            print("Bildirim Tipi: ", reportType)


            if (len(stockCode) >= 4):
                print("Stok kodu 3 haneden büyük, hisse senedi olabilir")
            else:
                print("Stok kodu 4 haneden küçük, hisse senedi değil. Sıradaki bildiri no'ya geçiliyor.")
                return

            for p in soup.find_all('div', {"class": "w-col w-col-3 modal-briefsumcol"}):
                for y in p.find_all('div', {"type-small bi-lightgray"}):
                    if y.text == "Yıl":
                        year = y.find_next('div').text
                        print("year: ", year)
                    elif y.text == "Periyot":
                        period = y.find_next('div').text
                        if period == "Yıllık":
                            period = "12"
                        elif period == "9 Aylık":
                            period = "09"
                        elif period == "6 Aylık":
                            period = "06"
                        elif period == "3 Aylık":
                            period = "03"
                        print("period: ", period)
                    elif y.text == "Gönderim Tarihi":
                        gt = y.find_next('div').text
                        # print("gonderimTarihi: ", gt)
                        gtArray = gt.split()
                        gonderimTarihi = gtArray[0]
                        gonderimTarihi = gonderimTarihi.replace('.', '-')
                        print("gonderimTarihi: ", gonderimTarihi)

            colName = year + period
            try:
                colName = int(colName)
            except ValueError:
                print("colname= ",colName)
                print("ValueError: invalid literal for int hatası")
                print()
                return

            print("Dönem: ", colName)
            cols = [colName]

            sunumParaBirimiLabel = soup.find('td', {"class": "financial-header-title"})  # TODO
            if (sunumParaBirimiLabel):
                sunumParaBirimi = sunumParaBirimiLabel.find_next_sibling().text
                sunumParaBirimi = sunumParaBirimi.strip(' \n\t')
                print("sunumParaBirimi: ", sunumParaBirimi)
                sunumParaBirimi = sunumParaBirimi.replace('.', '')
                sunumParaBirimiSR = re.search(r'\d+', sunumParaBirimi)
                if (sunumParaBirimiSR):
                    sunumParaBirimiInt = int(sunumParaBirimiSR.group())
                print("sunumParaBirimiInt: ", sunumParaBirimiInt)

            # Duzeltilmemis fiyat kismi
            url = 'https://www.isyatirim.com.tr/_layouts/15/Isyatirim.Website/Common/Data.aspx/HisseTekil?hisse=%s&startdate=%s&enddate=%s' % (
                stockCode, gonderimTarihi, gonderimTarihi)
            print("url: ", url)
            session = requests.Session()
            retry = Retry(connect=5, backoff_factor=1)
            adapter = HTTPAdapter(max_retries=retry)
            session.mount('http://', adapter)
            session.mount('https://', adapter)

            response = session.get(url)  # requests.get(url)

            if response.status_code != 204:  # serverdan boş response dönüp dönmediğini kontrol et
                fiyatBilgiJSON = response.json()
                # print("fiyatBilgiJSON: ", fiyatBilgiJSON)
                fiyatBilgiData = fiyatBilgiJSON['value']
                if len(fiyatBilgiData) != 0:
                    fiyatBilgiData = fiyatBilgiData[0]
                    duzeltilmemisFiyat = fiyatBilgiData['HG_KAPANIS']
                    # print("fiyatBilgiData:",fiyatBilgiData)
                    print("düzeltilmemiş fiyat at", gonderimTarihi, ": ", duzeltilmemisFiyat)
                else:
                    print("düzeltilmemiş fiyat bilgisine ulaşılamadı, 0'a set edildi")
                    duzeltilmemisFiyat = 0


            # str1 = '.*general_role_.*data-input-row.*presentation-enabled'
            # str2 = '.*holding_role_.*data-input-row.*presentation-enabled'
            str3 = '.*_role_.*data-input-row.*presentation-enabled'
            trTagClass = re.compile(str3)

            labelClass = "gwt-Label multi-language-content content-tr"
            currDataClass = re.compile("taxonomy-context-value.*")

            df = pd.DataFrame(columns=cols)


            i = 0
            lst = set()
            isHasılatLabelFound=False
            varliklarUstLabel = ""
            yukumluluklerUstLabel = ""
            for EachPart in soup.find_all('tr', {"class": trTagClass}):
                for ep in EachPart.find_all(True, {"class": labelClass}):

                    label = ep.get_text()
                    label = label.strip(' \n\t')

                    if label == "TOPLAM DÖNEN VARLIKLAR":
                        varliklarUstLabel= "TOPLAM DÖNEN VARLIKLAR";

                    elif label == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                        yukumluluklerUstLabel= "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER";

                    # Varliklar

                    elif label == "Ticari Alacaklar":
                        if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                            label = "Duran Ticari Alacaklar"
                        else:
                            label = "Dönen Ticari Alacaklar"

                    elif label == "Finansal Yatırımlar":
                        if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                            label = "Duran Finansal Yatırımlar"
                        else:
                            label = "Dönen Finansal Yatırımlar"

                    elif label == "Diğer Alacaklar":
                        if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                            label = "Duran Diğer Alacaklar"
                        else:
                            label = "Dönen Diğer Alacaklar"

                    elif label == "Peşin Ödenmiş Giderler":
                        if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                            label = "Duran Peşin Ödenmiş Giderler"
                        else:
                            label = "Dönen Peşin Ödenmiş Giderler"

                    elif label == "Türev Araçlar":
                        if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                            label = "Duran Türev Araçlar"
                        else:
                            label = "Dönen Türev Araçlar"

                    elif label == "Finans Sektörü Faaliyetlerinden Alacaklar":
                        if varliklarUstLabel == "TOPLAM DÖNEN VARLIKLAR":
                            label = "Duran Finans Sektörü Faaliyetlerinden Alacaklar"
                        else:
                            label = "Dönen Finans Sektörü Faaliyetlerinden Alacaklar"

                    # Yukumlulukler

                    elif label == "Diğer Finansal Yükümlülükler":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Diğer Finansal Yükümlülükler1"
                        else:
                            label = "Kısa Diğer Finansal Yükümlülükler1"

                    elif label == "Ertelenmiş Gelirler":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Ertelenmiş Gelirler"
                        else:
                            label = "Kısa Ertelenmiş Gelirler"

                    elif label == "Diğer Borçlar":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Diğer Borçlar"
                        else:
                            label = "Kısa Diğer Borçlar"

                    elif label == "Ticari Borçlar":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Ticari Borçlar"
                        else:
                            label = "Kısa Ticari Borçlar"

                    elif label == "Türev Araçlar":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Türev Araçlar"
                        else:
                            label = "Kısa Türev Araçlar"

                    elif label == "Finans Sektörü Faaliyetlerinden Borçlar":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Finans Sektörü Faaliyetlerinden Borçlar"
                        else:
                            label = "Kısa Finans Sektörü Faaliyetlerinden Borçlar"

                    elif label == "Müşteri Sözleşmelerinden Doğan Yükümlülükler":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Müşteri Sözleşmelerinden Doğan Yükümlülükler"
                        else:
                            label = "Kısa Müşteri Sözleşmelerinden Doğan Yükümlülükler"

                    elif label == "Ertelenmiş Gelirler (Müşteri Sözleşmelerinden Doğan Yükümlülüklerin Dışında Kalanlar)":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Ertelenmiş Gelirler (Müşteri Sözleşmelerinden Doğan Yükümlülüklerin Dışında Kalanlar)"
                        else:
                            label = "Kısa Ertelenmiş Gelirler (Müşteri Sözleşmelerinden Doğan Yükümlülüklerin Dışında Kalanlar)"
                    elif label == "Çalışanlara Sağlanan Faydalar Kapsamında Borçlar":
                        if yukumluluklerUstLabel == "TOPLAM KISA VADELİ YÜKÜMLÜLÜKLER":
                            label = "Uzun Çalışanlara Sağlanan Faydalar Kapsamında Borçlar"
                        else:
                            label = "Kısa Çalışanlara Sağlanan Faydalar Kapsamında Borçlar"
                    elif label == "Hasılat":
                        isHasılatLabelFound=True


                    df.rename(index={i: label}, inplace=True)

                    res = EachPart.find('td', {"class": currDataClass})
                    value = res.text
                    value = value.strip(' \n\t')
                    if not lst.__contains__(label):
                        lst.add(label)
                        if value:
                            df.loc[label, colName] = float(value.replace('.', '').replace(',', '.'))
                        else:
                            df.loc[label, colName] = value
                        i = i + 1




            #if(isHasılatLabelFound==True):
            #    print("Hasılat labeli bulundu, hisse senedi olabilir")
            #else:
            #    print("Hasılat labeli bulunamadı, hisse senedi değil. Sıradaki bildiri no'ya geçiliyor.")
            #    return



            df[colName] = df[colName] * sunumParaBirimiInt  # sunumParaBirimi ile çarpım

            df.rename(index={i: "Düzeltilmemiş Fiyat"}, inplace=True)
            if not lst.__contains__("Düzeltilmemiş Fiyat"):
                lst.add("Düzeltilmemiş Fiyat")
                if duzeltilmemisFiyat:
                    df.loc["Düzeltilmemiş Fiyat", colName] = duzeltilmemisFiyat

            print(tabulate(df, headers=[colName]))

            print("=========================")

            fileName = os.path.join(dirname, "//Users//myilmaz//Documents//bist//bilancolar_yeni//bilancolar//" + stockCode + ".xlsx")
            os.makedirs(os.path.dirname(fileName), exist_ok=True)

            if os.path.isfile(fileName):
                archiveDf = pd.read_excel(fileName, index_col=0)
                df = df.loc[~df.index.duplicated(keep='first')]
                archiveDf = archiveDf.loc[~archiveDf.index.duplicated(keep='first')]
                new_df = pd.concat([archiveDf, df], axis=1, sort=False)
                new_df = new_df.loc[:, ~new_df.columns.duplicated(keep='last')]  # duplike kolonları sil
                new_df = new_df.replace('', 0)
                new_df.to_excel(fileName)
            else:
                df = df.replace('', 0)
                try:
                    df.to_excel(fileName)
                except ValueError:
                    print("fileName= ", fileName)
                    print("ValueError: No engine for filetype: '' hatası")
                    print()
                    return

            f = open("//Users//myilmaz//Documents//bist//bilancolar_yeni//RaporBildiriNolar.txt", "a")
            f.write(bildiriNumber + "-")
            f.close()

            f = open("//Users//myilmaz//Documents//bist//bilancolar_yeni//BilancosuYeniGelenHisseler.txt", "a")
            f.write(stockCode + "-")
            f.close()

    f2 = open("//Users//myilmaz//Documents//bist//bilancolar_yeni//SonBildiriNo.txt", "r+")
    f2.write(bildiriNumber)
    f2.close()

def fncMultiple(s):
    content = s
    content = content.strip()
    contentList = content.split("-")
    contentList = ["https://www.kap.org.tr/tr/Bildirim/" + x for x in contentList]
    print(contentList)

    for link in contentList:
        fnc(link)

#fnc("https://www.kap.org.tr/tr/Bildirim/846388")
#fncMultiple("873010-873193-873196-873216")

dirname = os.path.dirname(__file__)
#filename = os.path.join(dirname, "kap_linkler.txt")


file = open("//Users//myilmaz//Documents//bist//bilancolar_yeni//SonBildiriNo.txt", "r+")
content_=file.readline()
#content_=content_.strip()
#contentList_ =content_.split("-")
sonBildiriNo= content_
print("sonBildiriNo= ",sonBildiriNo)

print()
file.close()

for i in range(999999999):
    bildiriNo = str(int(sonBildiriNo) + 1 + i)
    link = "https://www.kap.org.tr/tr/Bildirim/" + bildiriNo
    print("link= ", link)
    print(" bildiri numarası: ", bildiriNo)
    fnc(link, bildiriNo)
    time.sleep(0.51)