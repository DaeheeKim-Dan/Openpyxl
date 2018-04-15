import openpyxl as ox


def deleteSC(address):
    return address.replace('$','')


wb=ox.load_workbook('o002_NamedCellControl.xlsx')
sht=wb[list(wb.defined_names['test'].destinations)[0][0]]
address=deleteSC(list(wb.defined_names['test'].destinations)[0][1])

sht[address][0][0].value='MergedCell'
print(sht[address][0][0].value)

sht[address][0][0].value='NamedCell'
print(sht[address][0][0].value)

wb.save('o002_NamedCellControl.xlsx')
