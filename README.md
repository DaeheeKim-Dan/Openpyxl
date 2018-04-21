# Openpyxl
Openypyxl Study

openpyxl study room.

Contents List
o001 : MergedCellStyle.py
        In this file, described how to change border, fill and font style of the merged cell.
o002 : NamedCellControl.py
		master	: This file open an excel file which has named cell and read and write a string 
				  from/into the cell with the named cell.
				     list(wb.defined_names['defName'].destinations)
		Test	: 엘셀 시트에 존재하는 이름상자들을 모두 추출하고, 그 주소값을 가지고 각 이름 상자의 내용을 편집하기.
				     defNames=wb.defined_names.definedName
				     defNames[i].name
				     list(defNames[i].destinations)
o003 : TransactionSheet.py
        master  : 거래명세서 양식서 로드하여 거래날짜 정보 입력.
