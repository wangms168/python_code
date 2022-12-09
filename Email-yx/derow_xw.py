import time

import xlwings as xw


def derow(xlfile):
    xlapp = xw.App(visible=False, add_book=False)
    xlapp.display_alerts = False  # 不显示Excel消息框
    xlapp.screen_updating = False  # 关闭屏幕更新,可加快宏的执行速度
    book = xlapp.books.open(xlfile)
    shts = book.sheets

    for sht in shts:  # for sheets
        # name = sht.name
        # print('name:', name)

        max_row = sht.cells.last_cell.row
        # bottom = sht.used_range.last_cell.row          # used已使用的包含了有格式的行
        bottom = sht.range('A' + str(max_row)).end('up').row + 2
        # print('bottom:', bottom)
        sht[f'{bottom}:{max_row - bottom}'].delete()

    book.save()
    book.close()
    xlapp.quit()


if __name__ == '__main__':
    xlFile = r'D:\清算文件\08-IRS客户日结单\2022利率IRS互换日结单\202212利率IRS互换日结单\202212利率IRS互换日结单-浦发银行\IRS客户日结单_20221202_0003321_6.xlsx'
    start_time = time.time()
    derow(xlFile)
    print(f'耗时：{time.time() - start_time}')
