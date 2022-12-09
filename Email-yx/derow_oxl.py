# from openpyxl import load_workbook
#
import time

xlfile = r'D:\清算文件\08-IRS客户日结单\2022利率IRS互换日结单\202212利率IRS互换日结单\202212利率IRS互换日结单-浦发银行\IRS客户日结单_20221201_0003321.xlsx'
#
# xlapp = ''
# book = load_workbook(xlfile)
# # 选择一个sheet
# # sheet = book.active
# # sheet1 = book.worksheets[0]           # 通过索引选择
# sheet1 = book['日终盯市市值明细']  # 通过表名选择
#
# # sht = self.xlBook.Worksheets(sheet)
#
# bottom = sheet1.max_row
# last_col = sheet1.max_column
# print(f'bottom:{bottom},last_col:{last_col}')
#
# last_row = len(sheet1["C"])  # 获取C列为标准的最后一行
# print(f'last_row:{last_row}')
#
# rows = sheet1.rows
# print('len', type(rows))
#
# i = 0
# for row in rows:
#     i += 1
#     print(i)
#     # for cell in row:
#
# # sheet1.append(['测试测试测试测试测试测试测试测试', 2, 3])
# # book.save(filename=xlfile)

import xlwings as xw

xlfile = r'D:\清算文件\08-IRS客户日结单\2022利率IRS互换日结单\202212利率IRS互换日结单\202212利率IRS互换日结单-浦发银行\IRS客户日结单_20221202_0003321_6.xlsx'

xlapp = xw.App(visible=False, add_book=False)
# 不显示Excel消息框
xlapp.display_alerts = False
# 关闭屏幕更新,可加快宏的执行速度
xlapp.screen_updating = False
wb = xlapp.books.open(xlfile)
ws = wb.sheets['日终盯市市值明细']

start_time=time.time()
max_row = ws.cells.last_cell.row
# bottom = ws.used_range.last_cell.row          # used已使用的包含了有格式的行
bottom = ws.range('A' + str(max_row)).end('up').row+2
print('bottom:', bottom)
print(f'耗时：{time.time() - start_time}')

ws[f'{bottom}:{max_row - bottom}'].delete()

# 以循环的空值来获取有数值的最后一行
row = 1
while ws.range('A' + str(row)).value is not None:
    row += 1
print('row:', row-1)

# 输出打开的excel的绝对路径
# print(wb.fullname)
wb.save()
wb.close()
# 退出excel程序，
xlapp.quit()
