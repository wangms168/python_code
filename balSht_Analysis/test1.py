import xlrd

path_data = r"E:\调自有汇总表"
path = 'E:/调自有汇总邮件附件/'
# myList = os.listdir(path)

xlsx = xlrd.open_workbook(r"E:\python\余额表-对比\对比分析.xls")
table = xlsx.sheet_by_name(xlsx.sheet_names()[0])

f21 = table.cell(21, 6).value

print(f21)
