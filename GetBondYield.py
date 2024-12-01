from bs4 import BeautifulSoup
import requests

# def returnBondYield():
#     url = "https://www.bloomberght.com/"
#     page = requests.get(url)
#     soup = BeautifulSoup(page.text, 'html.parser')
#     list = soup.find("small", attrs={"data-secid": "TAHVIL2Y"})
#
#     faizOraniText = list.text
#     faizOraniText2 = faizOraniText.replace(",", ".")
#     faizOraniFloat = float(faizOraniText2)
#
#     return (faizOraniFloat/100)
#
#

def returnBondYield():

    url = "https://www.bloomberght.com/faiz-bono"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    tahvil_row = soup.find('tr', class_='security-faiz')
    if tahvil_row:
        faizOraniText = tahvil_row.find_all('td')[1].text.strip()
        faizOraniText2 = faizOraniText.replace(",", ".")
        faizOraniFloat = float(faizOraniText2)
        return (faizOraniFloat/100)

    else:
        print("TR 2 YILLIK TAHVİL değeri bulunamadı.")
