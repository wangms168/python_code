import xlwings as xw
from datetime import datetime, timedelta

def writeXL(xlfile, dateStr, SH_A, SH_jj, SH_kcb, SZ_A, SZ_jj):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts  = False       # 不显示Excel消息框
    app.screen_updating = False       # 关闭屏幕更新,可加快宏的执行速度
    wb = app.books.open(xlfile)
    shts = wb.sheets
    shts = [s.name for s in shts]
    # print(shts)
    # for sht in shts:  # for sheets
    #     name = sht.name
    #     print('name:', name)

    date_obj =  datetime.strptime(dateStr, '%Y-%m-%d')       # 日期字符串转换成日期对象
    # Ymd_str = date_obj.strftime('%Y年%#m月%#d日')
    ym_str = date_obj.strftime('%y%m')
    # print(ym_str)
    if ym_str in shts:
        sht = wb.sheets[ym_str]
        # bottom = sht.used_range.last_cell.row          # used已使用的包含了有格式的行
        max_row = sht.cells.last_cell.row
        bottom = sht.range('A' + str(max_row)).end('up').row
        # 对excel单元格每一行遍历，若行首单元格名称存在于列表中，则将其背景标黄
        for i in range(5,bottom+1):
            date = sht.range(f'A{i}').value
            if date == date_obj:
                sht.range(f'B{i}').value = SH_A
                sht.range(f'C{i}').value = SH_jj
                sht.range(f'D{i}').value = SH_kcb
                sht.range(f'E{i}').value = SZ_A
                sht.range(f'F{i}').value = SZ_jj
                break
        wb.save()
        wb.close()
        app.quit()
    else:
        print('没有'+ym_str+'这张表，请增加模板后再运行本程序！')

if __name__ == "__main__":
    xlfile = r'docs\2023年市场交易量统计表-RPA.xls'
    date = '2023-03-01'
    # writeXL(xlfile, date, SH_A, SH_jj, SH_kcb, SZ_A, SZ_jj)