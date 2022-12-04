# 从写好的class导入方法
from easyExcel.easyExcel import easyExcel

path = r'E:\python_code\easyExcel'
#读取excle
excel = easyExcel(path + '\\test.xlsx')


#获取Sheet1  第9行2列内的数据
print(excel.getCell('Sheet1', 9, 2))

#修改数据
excel.setCell('Sheet1',9,2,"newdata")

# #删除12-13行
# excel.deleteRowsCols('sheet1', '12:13', None)

# #删除E:F列
# excel.deleteRowsCols('sheet1', None, 'E:F')

excel.setRangeCellformat('sheet1', 1, 'A', 39, 'F')

#保存文件
excel.save(path + '\\out.xlsx')

#关闭文件
excel.close()