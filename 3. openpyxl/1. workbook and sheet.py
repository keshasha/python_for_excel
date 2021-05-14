from openpyxl import Workbook

wb = Workbook()
ws1 = wb.create_sheet("new sheet1")     # 새로운 sheet 생성
ws2 = wb.create_sheet("new sheet2", 2)  # 2번째에 sheet 생성

ws = wb["new sheet1"]           # dict형태로 sheet 접근 가능

target = wb.copy_worksheet(ws)  # sheet 복사


wb.save("sample.xlsx")
wb.close()