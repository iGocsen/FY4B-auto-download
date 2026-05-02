import openpyxl
wb = openpyxl.load_workbook(r'E:\云图\FY4B\待导入日期\链接生成-有效期24小时.xlsx', data_only=True)
ws = wb.active
# 检查所有列是否有值
for row in range(1, 163):
    b_val = ws[f"B{row}"].value
    if b_val is not None and not str(b_val).startswith('='):
        print(f'Row {row}: B={b_val}')
