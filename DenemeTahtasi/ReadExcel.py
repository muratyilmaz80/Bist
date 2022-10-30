import xlrd

loc = ("D:\\aa.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(1)

satis_onceki_4_ceyrek = sheet.cell_value(9,3)
print (satis_onceki_4_ceyrek)

satis_onceki_3_ceyrek = sheet.cell_value(10,3)
print (satis_onceki_3_ceyrek)