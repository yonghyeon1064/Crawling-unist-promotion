from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os

output = Workbook()
sheetO = output.active

#받은 경로안의 파일들 리스트로 만듬
print("xls 지원 안함")
path = input("excel파일 경로를 입력하세요(현재 폴더를 원하면 그냥 엔터): ")
file_list = os.listdir("./"+path)

#output의 어느 row에서 입력을 시작할지
startrow = 1

#input에서 값을 읽어올때 사용하는 변수들
row = 2

for file in file_list:
    if file[-5:] == ".xlsx": #모든 파일중 .xlsx 형식만 거름
        wb = load_workbook(filename = file)
        sheetI = wb.active

        sheetO.cell(row = startrow, column = 1, value= file )

        #date, 보도, 제목, 출처
        while (sheetI.cell(row, 2).value != None):
            #date
            sheetO.cell(startrow + row - 1, 4).value = int((sheetI.cell(row, 2).value).replace(".",""))
            sheetO.cell(startrow + row - 1, 4).alignment = Alignment(horizontal='center', vertical='center')
            #보도
            sheetO.cell(row = startrow + row - 1, column = 5, value= "보도" )
            sheetO.cell(startrow + row - 1, 5).value = "보도"
            sheetO.cell(startrow + row - 1, 5).alignment = Alignment(horizontal='center', vertical='center')
            #title
            sheetO.cell(startrow + row - 1, 6).value = sheetI.cell(row, 3).value
            sheetO.cell(startrow + row - 1, 6).hyperlink = sheetI.cell(row, 6).value
            #sheetO.cell(startrow + row - 1, 6).alignment = Alignment(horizontal='center', vertical='center')
            sheetO.cell(startrow + row - 1, 6).style = "Hyperlink"
            #출처
            sheetO.cell(startrow + row - 1, 7).value = sheetI.cell(row, 4).value
            sheetO.cell(startrow + row - 1, 7).alignment = Alignment(horizontal='left', vertical='bottom')

            row = row+1
        startrow = row+1
        row=2


output.save(filename = "result.xlsx")