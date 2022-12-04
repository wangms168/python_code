# https://www.cnblogs.com/xlovepython/p/14257023.html
# https://www.cnblogs.com/jasonli-01/articles/6612020.html

from win32com.client import Dispatch, constants


class easyExcel:

    def __init__(self, filename=None):              #打开文件或者新建文件（如果不存在的话） 
        self.xlApp = Dispatch('Excel.Application')
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
        del self.xlApp
    def getSheets(self):
        "获取sheet个数和列表。"
        shtCount = self.xlApp.Worksheets.Count;
        shtList = [sht.Name for sht in self.xlBook.Worksheets]
        return shtCount, shtList
    def getSht(self, sheet):
        "获取某个sheet,形参sheet以序号index或名字name均可以。"
        sht = self.xlBook.Worksheets(sheet)
        return sht
    def getCell(self, sheet, row, col):
        "Get value of one cell #获取单元格的数据,行列全是数字,从1开始"
        sht = self.getSht(sheet)
        return sht.Cells(row, col).Value
    def setCell(self, sheet, row, col, value):            
        "set value of one cell #设置单元格的数据"
        sht = self.getSht(sheet)
        sht.Cells(row, col).Value = value
    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.getSht(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
    def setRange(self, sheet, leftCol, topRow, data):
        """insert a 2d array starting at given location.
        Works out the size needed for itself"""
        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        sht = self.getSht(sheet)
        sht.Range(
            sht.Cells(topRow, leftCol),
            sht.Cells(bottomRow, rightCol)
            ).Value = data
    def setRangeCellformat(self, sheet, topRow, leftCol, bottomRow, rightCol):  
        "set format of one cell"    
        sht = self.getSht(sheet)
        Range = sht.Range(
            sht.Cells(topRow, leftCol),
            sht.Cells(bottomRow, rightCol)
            )
        Range.Font.Size = 11                                #字体大小  
        Range.Font.Bold = False                             #是否黑体  
        Range.Name = "Arial"                                #字体类型  
        # Range.Interior.ColorIndex = 3                       #表格背景  
        Range.BorderAround(1,2)                             #表格边框  
        sht.Rows(f'{topRow}:{bottomRow}').RowHeight = 15    #行高  
        Range.HorizontalAlignment = -4131                   #水平居中xlCenter  
        Range.VerticalAlignment = -4160 #
    def deleteRowsCols(self, sheet, row=None, col=None):
        "删除行row或列col,除某行 或者删除第几行到第几行 1   '1:3'   or  'A:C'"
        sht = self.getSht(sheet)
        if row:
            sht.Rows(row).Delete()
        if col:
            sht.Columns(col).Delete()
    def inserRow(self,sheet,row):
        sht = self.getSht(sheet)
        sht.Rows(row).Insert(1)
    def getUsedRows(self, sheet):
        "获取某个sheet已使用的最大行数"
        sht = self.getSht(sheet)
        # bottom = sht.UsedRange.Rows.Count               # 这个不准
        bottom = sht.Range('A'+str(sht.Rows.Count)).End(constants.xlUp).Row 
        return bottom
    def getUsedCols(self, sheet):
        "获取某个sheet已使用的最大列数"
        sht = self.getSht(sheet)
        # right = sht.UsedRange.Columns.Count               # 这个不准
        right = sht.Range('A'+str(sht.Rows.Count)).End(constants.xlUp).Col 
        return right
    def getMaxRows(self, sheet):
        "获取某个sheet的最大行数"
        sht = self.getSht(sheet)
        MaxRows = sht.Rows.Count
        return MaxRows
    def getMaxCols(self, sheet):
        "获取某个sheet的最大列数"
        sht = self.getSht(sheet)
        MaxCols = sht.Columns.Count
        return MaxCols
    def getContiguousRange(self, sheet, row, col):
        """Tracks down and across from top left cell until it
        encounters blank cells; returns the non-blank range.
        Looks at first row and column; blanks at bottom or right
        are OK and return None witin the array。毗连区,作用不大。"""
        sht = self.getSht(sheet)
        # find the bottom row
        bottom = row
        while sht.Cells(bottom + 1, col).Value not in [None, '']:
            bottom = bottom + 1
        # right column
        right = col
        while sht.Cells(row, right + 1).Value not in [None, '']:
            right = right + 1
        return sht.Range(sht.Cells(row, col), sht.Cells(bottom, right)).Value

