from bs4 import BeautifulSoup
import requests
import re

def returnHisseHalkaAciklikOrani(hisseAdi):
    url1 = "https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/sirket-karti.aspx?hisse="
    url2 = hisseAdi
    url = url1 + url2
    print (url)
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    # Halka açıklık oranını çekme
    halka_aciklik_th = soup.find('th', string='Halka Açıklık Oranı (%)')
    halka_aciklik_td = halka_aciklik_th.find_next('td')
    halka_aciklik_orani_text = halka_aciklik_td.get_text()
    halka_aciklik_orani_text2 = halka_aciklik_orani_text.replace(",", ".")
    halka_aciklik_orani = float(halka_aciklik_orani_text2)/100
    # print (hisseAdi + " Halka Açıklık Oranı: " + "{:.2%}".format(halka_aciklik_orani))
    return (halka_aciklik_orani)
