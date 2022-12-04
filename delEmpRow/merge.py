import pandas as pd
import xlrd
import xlwt
import os


pd.set_option("display.max_rows", 1000)


def get_tabledata(table_name):
    print('路径：', path)
    print('文件名：', table_name)
    print('全路径文件名：', path+table_name)
    xlsx = xlrd.open_workbook(path+table_name)
    print('xlsx===', xlsx)

    table = xlsx.sheet_by_name(xlsx.sheet_names()[0])
    print('table===', table)

    #table_data = None
    new_col = ['结算机构','申请部门','调拨金额','调拨用途','收款账号']
    data0 = table.cell(5,3).value
    data0 = data0.replace(" ", "")
    data1 = table.cell(4,2).value
    data2 = table.cell(8,5).value
    data3 = table.cell(7,4).value[:12]
    data4 = table.cell(10,4).value
    df = pd.DataFrame([[data0,data1,data2,data3,data4]],columns=new_col)
    print(data3)
    return df

def _get_style(borders_major='tblr',width_major=1,width_minor=1,font_size=12):
    '''
    borders_major：选择主要边框 默认为4边全选 边框粗细为 width_major
    width_major：主要边框的粗细 默认为1
    wdith_minor：次要边框的粗细 默认为1
    font_size：字体大小 默认为10
    '''
    style = xlwt.XFStyle()       # Create Style
    font = xlwt.Font()           # Create Font
    borders = xlwt.Borders()     # Create Borders
    alignment = xlwt.Alignment() # Create Alignment
    font.name = '黑体'           # 设置字体为 宋体
    font.height = font_size*20   # 设置字体大小为 10（10*20）
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    # 可以选择: HORZ_GENERAL,HORZ_LEFT,HORZ_CENTER,HORZ_RIGHT,HORZ_FILLED,
    #          HORZ_JUSTIFIED,HORZ_CENTER_ACROSS_SEL,HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER
    # 可以选择: VERT_TOP,VERT_CENTER,VERT_BOTTOM,VERT_JUSTIFIED,VERT_DISTRIBUTED
    alignment.wrap = 1           # 自动换行
    # 设置边框宽度
    # borders.left = width_major if 'l' in borders_major else width_minor
    # borders.right = width_major if 'r' in borders_major else width_minor
    # borders.top = width_major if 't' in borders_major else width_minor
    # borders.bottom = width_major if 'b' in borders_major else width_minor
    # 向style输入格式
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style

def del_file(path_data):
    for i in os.listdir(path_data) :
        file_data = path_data + "\\" + i
        if os.path.isfile(file_data) == True:
            os.remove(file_data)
        else:
            del_file(file_data)

def doing():
    all_data = None
    for i in range(len(newList)):
        try:
            if i==0:
                all_data = get_tabledata(newList[i])
            else:
                all_data = pd.concat([all_data,get_tabledata(newList[i])],axis=0,join='outer',sort=False)
        except:
            print("数据有误!")
    if all_data is None:
        print("无法解析该表数据！")
    else:
        all_data.reset_index(inplace=True,drop=True)
        all_data.index = range(1,len(all_data)+1)
        all_data.index.name='序号'
        all_data.reset_index(drop=False)

        workbook = xlwt.Workbook(encoding='UTF-8')
        worksheet1 = workbook.add_sheet('Sheet1')
        # 设置格式

        # 表头
        font_style_lrtb     = _get_style('lrtb',2)

        cur_row = 0
        for i in range(len(all_data.columns)):
            worksheet1.write(cur_row,i,all_data.columns[i],font_style_lrtb)

        cur_row += 1
        for i in range(len(all_data)):
            for j in range(len(all_data.iloc[i].values)):
                cur_data = all_data.iloc[i].values[j]
                worksheet1.write(cur_row,j,cur_data)
            cur_row += 1

        worksheet1.col(0).width = 500*14
        worksheet1.col(1).width = 300*14
        worksheet1.col(2).width = 500*14
        worksheet1.col(3).width = 500*14
        worksheet1.col(4).width = 500*14

        workbook.save('E:/调自有汇总表/自有资金汇总表.xls')
        print('成功导出excel。文件路径：E:/调自有汇总表/')
        # del_file(path)

path_data = r"E:\调自有汇总表"
# del_file(path_data)

# path='E:/调自有汇总邮件附件/'
path=r'E:\python\DelEmpLine\202211利率IRS互换日结单-浦发银行/'

# path='F:/QQ接收到的文件/read_email/dist/read_email/附件/'
myList = os.listdir(path)
# print(myList)
if myList:
  newList = []
  for fileName in myList:
      if os.path.splitext(fileName)[1] == '.xlsx' or os.path.splitext(fileName)[1] == '.xls' :
        newList.append(fileName)
  print(newList)
  doing()
else:
  print('文件夹为空')
  value = '无新邮件，流程结束。'

