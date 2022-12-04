import pandas as pd
from win32com.client import constants

def append(xlapp_flag, out_ws, fl_list):
    if xlapp_flag == "win32com":
        # bottom = out_ws.UsedRange.Rows.Count                  # 不准
        bottom = out_ws.Range('A'+str(out_ws.Rows.Count)).End(constants.xlUp).Row 
        out_ws.Range(out_ws.Cells(bottom+1, 1), out_ws.Cells(bottom+1, len(fl_list))).Value = fl_list
    if xlapp_flag == "xlwings":
        bottom = out_ws.range('A' + str(out_ws.cells.last_cell.row)).end('up').row
        out_ws.range("A"+ str(bottom+1)).options(index=False, header=False).value = fl_list
    if xlapp_flag == "openpyxl":
        bottom = out_ws.max_row
        last_col = out_ws.max_column
        out_ws.append(fl_list)


def km66011501(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提工资'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '计提工资'
    fl_list[6] = '22110101'
    fl_list[8] = None
    fl_list[9] = None
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '发放工资'
    fl_list[6] = '22110101'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km66011501_D(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提工资'
    fl_list[6] = '66011509'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '计提工资'
    fl_list[6] = '22110105'
    fl_list[8] = None
    fl_list[9] = None
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '发放工资'
    fl_list[6] = '22110105'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km66011501_L(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提工资'
    fl_list[6] = '66011519'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '计提工资'
    fl_list[6] = '22110103'
    fl_list[8] = None
    fl_list[9] = None
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '发放工资'
    fl_list[6] = '22110103'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km66011502(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提奖金'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '计提奖金'
    fl_list[6] = '22110102'
    fl_list[8] = None
    fl_list[9] = None
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '发放奖金'
    fl_list[6] = '22110102'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km66011503(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提过节费'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011504(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提交通补贴'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011505(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提伙食补贴'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011506(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提通讯补贴'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011507(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提辞退福利'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '计提辞退福利'
    fl_list[6] = '221109'
    fl_list[8] = None
    fl_list[9] = None
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '发放辞退福利'
    fl_list[6] = '221109'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km66011510(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提劳保补贴'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011519(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提其他补贴'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km22110103(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提补贴小计'
    fl_list[6] = '22110103'
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '发放补贴小计'
    fl_list[6] = '22110103'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km66011601(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提医疗卫生'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011602(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提防暑降温'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011603(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提取暖费'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011604(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提独生子女费'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011605(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提文体宣传费'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011606(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提探亲路费'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011607(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提职工困难补助'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011608(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提丧葬抚恤救济费'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011609(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提食堂费用'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km66011619(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提其他福利费'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km221102(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提福利小计'
    fl_list[6] = '221102'
    fl_list[8] = None
    fl_list[9] = None
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None

    append(xlapp_flag, out_ws, fl_list)
    fl_list[5] = '【' + lb_2 + '】' + '发放福利小计'
    fl_list[6] = '221102'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km660115080101(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提佣金提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km660115080102(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提服务提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km660115080103(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提期权业务提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km660115080104(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提IB业务提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km660115080105(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提管理津贴'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km660115080106(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提投顾业务提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km660115080107(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提两融净息差提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km660115080108(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提开户奖'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601150802(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提基金保有量提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601150803(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提基金销售奖励'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601150804(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提基金销售手续费返还'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601150805(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提公司理财产品销售奖励'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    fl_list[17] = 'gl01001' + ':销售产品类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601150806(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提代理销售保险产品'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601150807(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提非公募产品销售奖励'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601150819(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '计提其他提成'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[15] = jbb + ':部门'
    fl_list[16] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km22110104(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    #    if '22411901' in in_df.columns:
    #        print('有22411901')
    #        var = round(in_df['22411901']['合计'], 2)  # 总部下拨奖励合计金额
    #        amt = amt - var
    #    else:
    #        print('无22411901')

    fl_list[5] = '【' + lb_2 + '】' + '计提提成小计'
    fl_list[6] = '22110104'
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '【' + lb_2 + '】' + '发放提成小计'
    fl_list[6] = '22110104'
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = lb_1 + ':人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km22411901(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    fl_list[5] = '发放总部下拨奖励'
    fl_list[6] = k
    fl_list[8] = str(amt)
    fl_list[9] = str(amt)
    append(xlapp_flag, out_ws, fl_list)


def km22410401(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    fl_list[5] = '扣个人养老'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    append(xlapp_flag, out_ws, fl_list)


def km22410402(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    fl_list[5] = '扣个人失业'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    append(xlapp_flag, out_ws, fl_list)


def km22410403(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    fl_list[5] = '扣个人医疗'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    append(xlapp_flag, out_ws, fl_list)


def km22410404(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    fl_list[5] = '扣个人公积金'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    append(xlapp_flag, out_ws, fl_list)


def km22410409(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    fl_list[5] = '扣其他保险(企业年金)'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    append(xlapp_flag, out_ws, fl_list)


def km222105(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '【' + lb_2 + '】' + '扣个人所得税'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km6601070201(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    fl_list[5] = '租入房产员工租用扣房租'
    fl_list[6] = k
    fl_list[8] = str(-amt)
    fl_list[9] = str(-amt)
    fl_list[15] = jbb + ':部门'
    append(xlapp_flag, out_ws, fl_list)


def km605102(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    amt = round(amt/1.05,2)
    fl_list[5] = '自有房产员工租用扣房租'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = jbb + ':部门'
    append(xlapp_flag, out_ws, fl_list)
    
    amt = round(amt*0.05,2)
    fl_list[5] = '自有房产员工租用扣房租5%增值税'
    fl_list[6] = '2221160201'
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = None
    append(xlapp_flag, out_ws, fl_list)


def km221105(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    lb_1 = lb.split(':')[0]
    lb_2 = lb.split(':')[1]

    fl_list[5] = '扣工会经费'
    fl_list[6] = k
    fl_list[11] = str(amt)
    fl_list[12] = str(amt)
    fl_list[15] = lb_1 + ':人员类别'
    append(xlapp_flag, out_ws, fl_list)


def km224107(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    for i, v in amt.items():                                                    # items是pd.Serie的每个项目，i是键(df的index)、v是值('224107'这列上的数据)
        if v and (not (pd.isnull(v))):
            i1 = i.split(':')[0]
            i2 = i.split(':')[1]
            # print('i1=', i1, 'i2=', i2, '  ', 'v=', v)
            fl_list[5] = '扣风险金'
            fl_list[6] = k
            fl_list[11] = str(round(v, 2))
            fl_list[12] = str(round(v, 2))
            fl_list[15] = i1 + ':人员档案'
            fl_list[16] = i2 + ':人员类别'
            append(xlapp_flag, out_ws, fl_list)


def km1001(xlapp_flag, SInfo_df, fl_list, amt, jbb, in_df, out_ws):
    YYB_bm = fl_list[10]

    if SInfo_df['工资结算户'][YYB_bm] == "总部统一结算":
        kmbm = "114305"

        fl_list[5] = '发放员工工资'
        fl_list[6] = kmbm
        fl_list[11] = str(amt)
        fl_list[12] = str(amt)
        fl_list[15] = '1101:客商'
        append(xlapp_flag, out_ws, fl_list)

    elif SInfo_df['工资结算户'][YYB_bm] == "基本户":
        kmbm = SInfo_df['基本户-科目编码'][YYB_bm]
        yhzh = SInfo_df['基本户-银行账户编码'][YYB_bm] + ':银行账户'

        if pd.isna(SInfo_df)['基本户-科目编码'][YYB_bm]:                         # 若“科目编码”为空、则使用1001。 pd.isna(SInfo_df)将各元素值转化为True或False
            kmbm = '1001'
            yhzh = ''

        fl_list[5] = '发放员工工资'
        fl_list[6] = kmbm
        fl_list[11] = str(amt)
        fl_list[12] = str(amt)
        fl_list[15] = yhzh
        append(xlapp_flag, out_ws, fl_list)

    elif SInfo_df['工资结算户'][YYB_bm] == "现金":
        fl_list[5] = '发放员工工资'
        fl_list[6] = '1001'
        fl_list[11] = str(amt)
        fl_list[12] = str(amt)
        append(xlapp_flag, out_ws, fl_list)

    var = round(in_df['66011501']['合计'], 2)  # 基本工资合计数
    if 'D小计' in in_df.index:
        var_D = round(in_df['66011501']['D小计'], 2)
        var -= var_D
    if 'L小计' in in_df.index:
        var_L = round(in_df['66011501']['L小计'], 2)
        var -= var_L

    # fl_list[5] = '计提2%的工会经费'
    # fl_list[6] = '660117'
    # fl_list[8] = str(round(var * 0.02, 2))
    # fl_list[9] = str(round(var * 0.02, 2))
    # fl_list[11] = None
    # fl_list[12] = None
    # fl_list[15] = jbb + ':部门'
    # fl_list[16] = '01:人员类别'
    # append(xlapp_flag, out_ws, fl_list)

    # fl_list[5] = '计提2%的工会经费'
    # fl_list[6] = '221105'
    # fl_list[8] = None
    # fl_list[9] = None
    # fl_list[11] = str(round(var * 0.02, 2))
    # fl_list[12] = str(round(var * 0.02, 2))
    # fl_list[15] = '01:人员类别'
    # fl_list[16] = None
    # append(xlapp_flag, out_ws, fl_list)


dict_AB = {
    "66011501": km66011501,             # 工资
    "66011502": km66011502,             # 奖金
    "66011503": km66011503,             # 过节费
    "66011504": km66011504,             # 交通补贴
    "66011505": km66011505,             # 伙食补贴
    "66011506": km66011506,             # 通讯补贴
    "66011510": km66011510,             # 劳保补贴
    "66011519": km66011519,             # 其他补贴
    "22110103": km22110103,             # 应付工资\其他  补贴小计
    "66011507": km66011507,             # 辞退福利

    "66011601": km66011601,             # 医疗卫生
    "66011602": km66011602,             # 防暑降温费
    "66011603": km66011603,             # 取暖费
    "66011604": km66011604,             # 独生子女费
    "66011605": km66011605,             # 文体宣传费
    "66011606": km66011606,             # 探亲路费
    "66011607": km66011607,             # 职工困难补助
    "66011608": km66011608,             # 丧葬抚恤救济费
    "66011609": km66011609,             # 食堂费用
    "66011619": km66011619,             # 其他福利
    "221102": km221102,                 # 应付福利

    "660115080101": km660115080101,     # 佣金提成
    "660115080102": km660115080102,     # 服务提成
    "660115080103": km660115080103,     # 期权业务提成
    "660115080104": km660115080104,     # IB业务提成
    "660115080105": km660115080105,     # 管理津贴
    "660115080106": km660115080106,     # 投顾业务提成
    "660115080107": km660115080107,     # 两融净息差提成
    "660115080108": km660115080108,     # 开户奖
    "6601150802": km6601150802,         # 基金保有量提成
    "6601150803": km6601150803,         # 基金销售奖励
    "6601150804": km6601150804,         # 基金销售手续费返还
    "6601150805": km6601150805,         # 公司理财产品销售奖励
    "6601150806": km6601150806,         # 代理销售保险产品
    "6601150807": km6601150807,         # 非公募产品销售奖励
    "6601150819": km6601150819,         # 其他
    "22110104": km22110104,             # 应付提成支出
    "222105": km222105,                 # 应付个税
}

dict_D = dict_AB.copy()
dict_D["66011501"] = km66011501_D       # D类人员
dict_L = dict_AB.copy()
dict_L["66011501"] = km66011501_L       # L类人员

dict = {
    "22411901": km22411901,             # 总部下拨奖励
    "22410401": km22410401,             # 应付个人养老
    "22410402": km22410402,             # 应付个人失业
    "22410403": km22410403,             # 应付个人医疗
    "22410404": km22410404,             # 应付个人公积金
    "22410409": km22410409,             # 应付企业年金(其他保险)
    "221105": km221105,                 # 应付工会经费
    "6601070201": km6601070201,         # 扣租入房产员工宿舍房租
    "605102": km605102,                 # 扣自有房产员工房租
    "224107": km224107,                 # 应付风险金
}

kmdm = [*dict_AB] + [*dict]             # kmdm = list(dict_AB)

def test():
    None

def switcher(dict, xlapp_flag, k, fl_list, amt, lb, jbb, out_ws):
    # func = dict.get(k, lambda xlapp_flag, k, fl_list, amt, lb, jbb, out_ws: None)
    # return func(xlapp_flag, k, fl_list, amt, lb, jbb, out_ws)
    dict.get(k, lambda xlapp_flag, k, fl_list, amt, lb, jbb, out_ws: None) \
        (xlapp_flag, k, fl_list, amt, lb, jbb, out_ws)

# dict.get(key[, value]) 
# 参数
# key -- 字典中要查找的键。
# value -- 可选，如果指定键的值不存在时，返回该默认值。
# 返回值
# 返回指定键的值，如果键不在字典中返回默认值 None 或者设置的默认值。
# 如果返回默认值 None 或者设置的默认值，则 None 或者设置的默认值后加()，并没有None 或者设置的默认值这样的函数，所以程序就会报错。
# 对应 value 位置用 lambda 函数 ,则如果键不在字典中就会返回函数体为 none 这样一个函数，样如：
# def test():
#     None


# result = {
#   'a': lambda x: x * 5,
#   'b': lambda x: x + 7,
#   'c': lambda x: x - 2
# }.get(whatToUse, lambda x: x - 22)(value)
# where

# .get('c', lambda x: x - 22)(23)
# looks up "lambda x: x - 2" in the dict and uses it with x=23

# .get('xxx', lambda x: x - 22)(44)
# doesn't find it in the dict and uses the default "lambda x: x - 22" with x=44.
# print((lambda x: x - 22)(44))
# https://stackoverflow.com/questions/60208/replacements-for-switch-statement-in-python