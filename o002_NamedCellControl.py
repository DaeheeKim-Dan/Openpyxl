import openpyxl as ox


def deleteSC(address):
    return address.replace('$','')


def get_addr(defName):
    dn=list(defName.destinations)
    return deleteSC(dn[0][1])

wb=ox.load_workbook('o002_NamedCellControl.xlsx')
sht=wb.active

# Named Range List를 추출하고, 그 주소를 함께 출력
defNames=wb.defined_names.definedName
for item in defNames:
    print(item.name, ' ---> ', item.attr_text)
    

# 전체 named range의 값을 named range의 이름으로 변경
for i in range(len(defNames)):
    print(get_addr(defNames[i]))

for i in range(len(defNames)):
    sht[get_addr(defNames[i]).split(':')[0]].value=''

for i in range(len(defNames)):
    sht[get_addr(defNames[i]).split(':')[0]].value=defNames[i].name

    
##sht=wb[list(wb.defined_names['test'].destinations)[0][0]]
##address=deleteSC(list(wb.defined_names['test'].destinations)[0][1])


wb.save('o002_NamedCellControl.xlsx')
