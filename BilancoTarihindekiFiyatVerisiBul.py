import xlrd

varHisseAdi = "BFREN"
varBilancoDosyasi = ("//Users//myilmaz//Documents//bist//KAP//Finansal_Rapor_Ilan_Tarihleri//202003aylıkfinansalraportarihleri.xlsx")


wb = xlrd.open_workbook(varBilancoDosyasi)
sheet = wb.sheet_by_index(0)

def bilancoAciklanmaTarihiBul():
    for rowi in range(sheet.nrows):
        cell = sheet.cell(rowi, 10)
        if cell.value == varHisseAdi:
            datetime_date = xlrd.xldate_as_datetime(sheet.cell_value(rowi,12), 0)
            date_object = datetime_date.date()
            string_date = date_object.isoformat()
            return string_date
    print("Uygun Ceyrek Bulunamadi!!!")
    return -1

print (bilancoAciklanmaTarihiBul())