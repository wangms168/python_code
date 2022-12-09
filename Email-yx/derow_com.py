import os
from win32com.client import Dispatch, constants
# from win32com import client as com


class easyExcel:

    # 初始化excel
    def __init__(self, filename=None):
        self.xlApp = Dispatch('Excel.Application')
        # self.xlApp = com.gencache.EnsureDispatch('excel.application')
        self.xlApp.Visible = False
        self.xlApp.DisplayAlerts = False
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        self.xlApp.Quit()
        # del self.xlApp

    def getSheets(self):
        """获取sheet个数和列表。"""
        shtCount = self.xlApp.Worksheets.Count
        shtList = [sht.Name for sht in self.xlBook.Worksheets]
        return shtCount, shtList

    def getSht(self, sheet):
        """获取某个sheet,形参sheet以序号index或名字name均可以。"""
        sht = self.xlBook.Worksheets(sheet)
        return sht

    def getUsedRows(self, sheet):
        """获取某个sheet的最大行数"""
        sht = self.getSht(sheet)
        bottom = sht.Range('A' + str(sht.Rows.Count)).End(constants.xlUp).Row
        return bottom


def forExcel(path):
    path = os.path.abspath(path)  # win32不认识相对路径，故需转换为绝对路径。
    print('路径path=', path)
    file_list = os.listdir(path)  # 获取文件夹内文件名列表
    print('文件夹列表长度', len(file_list))

    for xlfile in file_list:  # for xlfiles
        print('文件名', xlfile)
        excel = easyExcel(path + '\\' + xlfile)
        shtCount, shtList = excel.getSheets()
        print('shtCount', shtCount)
        for i in range(shtCount):  # for sheets
            name = shtList[i]  # 列表序号从0开始
            print('sht_name:', name)
            ws = excel.getSht(i + 1)  # sheet用序号，从1开始
            # ws = excel.getSht(name)                    # sheet用名字
            MaxRow = ws.Rows.Count
            print('MaxRows:', MaxRow)
            bottom = excel.getUsedRows(i + 1) + 2
            print('bottom', bottom - 2)
            ws.Rows(f'{bottom}:{MaxRow}').Delete()
            excel.save()

        excel.close()


def doExcel(xlfile):
    excel = easyExcel(xlfile)
    shtCount, shtList = excel.getSheets()
    # print('shtCount', shtCount)
    for i in range(shtCount):  # for sheets
        name = shtList[i]  # 列表序号从0开始
        # print('sht_name:', name)
        ws = excel.getSht(i + 1)  # sheet用序号，从1开始
        # ws = excel.getSht(name)                    # sheet用名字
        MaxRow = ws.Rows.Count
        # print('MaxRows:', MaxRow)
        bottom = excel.getUsedRows(i + 1) + 2
        # print('bottom', bottom - 2)
        ws.Rows(f'{bottom}:{MaxRow}').Delete()
        excel.save()
    excel.close()


if __name__ == '__main__':
    Path = r'E:\202211利率IRS互换日结单-浦发银行'
    xlFile = r'E:\202211利率IRS互换日结单-浦发银行\IRS客户日结单_20221124_0003321_6.xlsx'
    forExcel(Path)
    # doExcel(xlFile)
    print('执行完毕！')
