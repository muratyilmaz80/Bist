from bs4 import BeautifulSoup
import requests

def returnBondYield():
    url = "https://www.bloomberght.com/"
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')
    list = soup.find("small", attrs={"data-secid": "TAHVIL2Y"})

    faizOraniText = list.text
    faizOraniText2 = faizOraniText.replace(",", ".")
    faizOraniFloat = float(faizOraniText2)
    return (faizOraniFloat)
