from bs4 import BeautifulSoup
import requests
import re

def returnGuncelHisseDegeri(hisseAdi):
    url = "https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/default.aspx"
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    a = soup.find("a", href=re.compile("hisse=" + hisseAdi)).parent.parent


    b = a.find("td", attrs={"class": "text-right"})

    hisseDegeriText = b.text
    hisseDegeriText2 = hisseDegeriText.replace(".", "")
    hisseDegeriText3 = hisseDegeriText2.replace(",", ".")
    hisseDegeriFloat = float(hisseDegeriText3)
    return (hisseDegeriFloat)
