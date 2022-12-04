import os
import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.width', 180)
import datetime
from shutil import copyfile

import yg_fl
import jjr_fl
import sb_fl
import gjj_fl


class Convert:
    pass


def cshfl(zdr_bm, yyb_bm):  # 初始化分录
    today = datetime.date.today().strftime("%Y-%m-%d")
    kjfl = [                            # 会计分录
        None,                           # 0  A列
        yyb_bm + '-0002',               # 1  B列，核算账簿
        '01',                           # 2  C列，凭证类别编码
        zdr_bm,                         # 3  D列，制单人编码
        today,                          # 4  E列，制单日期
        None,                           # 5  F列，摘要
        None,                           # 6  G列，科目编码
        '人民币',                       # 7  H列，币种
        None,                           # 8  I列，原币借方金额
        None,                           # 9  J列，本币借方金额
        yyb_bm,                         # 10 K列，业务单元编码
        None,                           # 11 L列，原币贷方金额
        None,                           # 12 M列，本币贷方金额
        today,                          # 13 N列，业务日期
        '1',                            # 14 O列，组织本币汇率
        None,                           # 15 P列，辅助核算1
        None,                           # 16 Q列，辅助核算2
        None,                           # 17 R列，辅助核算3
    ]
    return kjfl


def SInfo():
    global SInfo_df
    SInfo_xlFile = "docs\\结算信息.xlsx"                     # 结算信息excel文件
    SInfo_df = pd.read_excel(SInfo_xlFile, index_col=2 - 1, skiprows=1 - 1)
    return SInfo_df


SInfo_df = SInfo()


def create_dir_not_exist(path):
    if not os.path.exists(path):
        os.mkdir(path)


create_dir_not_exist("output\\")


def yg_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, out_ws):
    kmdm_AB = yg_fl.dict_AB.keys()                          # 获取字典的所有键,字典dict键key列表
    kmdm_Z = yg_fl.dict.keys()

    xj_list = [                         # 小计列表
        "22110103",                     # 补贴小计
        "221102",                       # 福利小计
        "22110104",                     # 提成小计
    ]

    bt_list = [                         # 补贴列表
        "66011503",                     # 过节费
        "66011504",                     # 交通补贴
        "66011505",                     # 伙食补贴
        "66011506",                     # 通讯补贴
        # "66011507",                   # 辞退福利
        "66011510",                     # 劳保补贴
        "66011519",                     # 其他补贴
    ]

    flf_list = [k for k in kmdm_AB if k[0:6] == "660116"]   # 福利费列表       在kmdm_AB列表中提取福利费科目代码列表
    tc_list = [k for k in kmdm_AB if k[0:8] == "66011508"]  # 提成列表         在kmdm_AB列表中提取提成支出科目代码列表

    amt_a_bt = 0                        # A类人员补贴、福利、提成小计金额
    amt_a_fl = 0
    amt_a_tc = 0

    amt_b_bt = 0                        # B类人员补贴、福利、提成小计金额
    amt_b_fl = 0
    amt_b_tc = 0

    amt_d_bt = 0                        # D类人员补贴、福利、提成小计金额
    amt_d_fl = 0
    amt_d_tc = 0

    amt_l_bt = 0                        # L类人员补贴、福利、提成小计金额
    amt_l_fl = 0
    amt_l_tc = 0

    if 'A小计' in in_df.index:                                                  # A类人员（员工）-01
        for k in kmdm_AB:                                                       # 做账循序科目代码列表for循环
            fl_list = cshfl(zdr_bm, yyb_bm)                                     # 初始化分录各字段list列表
            if (k in in_df.columns) or (k in xj_list):                          # 元素是否在表头col(科目代码)或小计列表中
                if k in in_df.columns:
                    amount_A = in_df[k]['A小计']                                # 取“A小计”行数据
                    if (amount_A != 0) and (not (pd.isnull(amount_A))):
                        amount_A = round(in_df[k]['A小计'], 2)                  # 对合计数据四舍五入，并转为字符型。
                        if k in bt_list:
                            amt_a_bt += amount_A
                        if k in flf_list:
                            amt_a_fl += amount_A
                        if k in tc_list:
                            amt_a_tc += amount_A
                        if k == '22110103':
                            amount_A = amt_a_bt
                        if k == '221102':
                            amount_A = amt_a_fl
                        if k == '22110104':
                            amount_A = amt_a_tc
                        yg_fl.switcher(yg_fl.dict_AB, xlapp_flag, k, fl_list, amount_A, '01:员工', jbb_bm, out_ws)

                if (k not in in_df.columns) and (k in xj_list):
                    if k == '22110103':
                        amount_A = amt_a_bt
                    if k == '221102':
                        amount_A = amt_a_fl
                    if k == '22110104':
                        amount_A = amt_a_tc
                    if amount_A != 0:
                        yg_fl.switcher(yg_fl.dict_AB, xlapp_flag, k, fl_list, amount_A, '01:员工', jbb_bm, out_ws)

    if 'B小计' in in_df.index:                                                  # B类人员（全日制营销人员）-05  
        
        for k in kmdm_AB:                                                       # 做账循序科目代码列表for循环
            fl_list = cshfl(zdr_bm, yyb_bm)                                     # 初始化分录各字段list列表
            if (k in in_df.columns) or (k in xj_list):                          # 元素是否在表头col(科目代码)或小计列表中
                if k in in_df.columns:
                    amount_B = in_df[k]['B小计']                                # 取“B小计”行数据
                    if (amount_B != 0) and (not (pd.isnull(amount_B))):         # 当有多个“B小计”时，amount_B就不是一个数值
                        amount_B = round(in_df[k]['B小计'], 2)                  # 对合计数据四舍五入，并转为字符型。
                        if k in bt_list:
                            amt_b_bt += amount_B
                        if k in flf_list:
                            amt_b_fl += amount_B
                        if k in tc_list:
                            amt_b_tc += amount_B
                        if k == '22110103':
                            amount_B = amt_b_bt
                        if k == '221102':
                            amount_B = amt_b_fl
                        if k == '22110104':
                            amount_B = amt_b_tc
                        yg_fl.switcher(yg_fl.dict_AB, xlapp_flag, k, fl_list, amount_B, '05:营销人员', jbb_bm, out_ws)

                if (k not in in_df.columns) and (k in xj_list):
                    if k == '22110103':
                        amount_B = amt_b_bt
                    if k == '221102':
                        amount_B = amt_b_fl
                    if k == '22110104':
                        amount_B = amt_b_tc
                    if amount_B != 0:
                        yg_fl.switcher(yg_fl.dict_AB, xlapp_flag, k, fl_list, amount_B, '05:营销人员', jbb_bm, out_ws)

    if 'D小计' in in_df.index:                                                  # D类人员（实习生）-08
        for k in kmdm_AB:                                                       # 做账循序科目代码列表for循环
            fl_list = cshfl(zdr_bm, yyb_bm)                                     # 初始化分录各字段list列表
            if (k in in_df.columns) or (k in xj_list):                          # 元素是否在表头col(科目代码)或小计列表中
                if k in in_df.columns:
                    amount_D = in_df[k]['D小计']                                # 取“D小计”行数据
                    if (amount_D != 0) and (not (pd.isnull(amount_D))):
                        amount_D = round(in_df[k]['D小计'], 2)                  # 对合计数据四舍五入，并转为字符型。
                        if k in bt_list:
                            amt_d_bt += amount_D
                        if k in flf_list:
                            amt_d_fl += amount_D
                        if k in tc_list:
                            amt_d_tc += amount_D
                        if k == '22110103':
                            amount_D = amt_d_bt
                        if k == '221102':
                            amount_D = amt_d_fl
                        if k == '22110104':
                            amount_D = amt_d_tc
                        yg_fl.switcher(yg_fl.dict_D, xlapp_flag, k, fl_list, amount_D, '08:实习生', jbb_bm, out_ws)

                if (k not in in_df.columns) and (k in xj_list):
                    if k == '22110103':
                        amount_D = amt_d_bt
                    if k == '221102':
                        amount_D = amt_d_fl
                    if k == '22110104':
                        amount_D = amt_d_tc
                    if amount_D != 0:
                        yg_fl.switcher(yg_fl.dict_D, xlapp_flag, k, fl_list, amount_D, '08:实习生', jbb_bm, out_ws)

    if 'L小计' in in_df.index:                                                  # L类人员（劳务）-09
        for k in kmdm_AB:                                                       # 做账循序科目代码列表for循环
            fl_list = cshfl(zdr_bm, yyb_bm)                                     # 初始化分录各字段list列表
            if (k in in_df.columns) or (k in xj_list):                          # 元素是否在表头col(科目代码)或小计列表中
                if k in in_df.columns:
                    amount_L = in_df[k]['L小计']                                # 取“L小计”行数据
                    if (amount_L != 0) and (not (pd.isnull(amount_L))):
                        amount_L = round(in_df[k]['L小计'], 2)                  # 对合计数据四舍五入，并转为字符型。
                        if k in bt_list:
                            amt_l_bt += amount_L
                        if k in flf_list:
                            amt_l_fl += amount_L
                        if k in tc_list:
                            amt_l_tc += amount_L
                        if k == '22110103':
                            amount_L = amt_l_bt
                        if k == '221102':
                            amount_L = amt_l_fl
                        if k == '22110104':
                            amount_L = amt_l_tc
                        yg_fl.switcher(yg_fl.dict_L, xlapp_flag, k, fl_list, amount_L, '09:劳务', jbb_bm, out_ws)

                if (k not in in_df.columns) and (k in xj_list):
                    if k == '22110103':
                        amount_L = amt_l_bt
                    if k == '221102':
                        amount_L = amt_l_fl
                    if k == '22110104':
                        amount_L = amt_l_tc
                    if amount_L != 0:
                        yg_fl.switcher(yg_fl.dict_L, xlapp_flag, k, fl_list, amount_L, '09:劳务', jbb_bm, out_ws)

    if '负责人小计' in in_df.index:
        pass

    # 全部人员合计
    for k in kmdm_Z:                                                            # 做账循序科目代码列表for循环
        fl_list = cshfl(zdr_bm, yyb_bm)                                         # 初始化分录各字段list列表
        if k in in_df.columns:                                                  # 做账循序科目代码列表元素是否在表头col(科目代码)中
            amount = in_df[k]['合计']                                           # 取“合计”行数据
            if (amount != 0) and (not (pd.isnull(amount))):
                amount = round(in_df[k]['合计'], 2)                             # 对合计数据四舍五入，并转为字符型。
                if k == '224107':
                    s = in_df[k]                                                # 获取'224107'这一列pd.Serie这个对象 
                    tc_list = ['A1小计', 'A2小计', 'A小计', 'B小计', 'D小计', 'L小计', '合计']        # 对这个pd.Serie进行瘦身过滤下
                    s = s[~s.index.isin(tc_list)]                               # 对这个pd.Serie过滤掉ti-list  ~是对True或False逻辑值取反
                    amount = s
                yg_fl.switcher(yg_fl.dict, xlapp_flag, k, fl_list, amount, '01:员工', jbb_bm, out_ws)

    # 支付结算分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第四次、初始化分录各字段list列表
    amt = round(in_df['1001']['合计'], 2)                                       # 对合计数据四舍五入，并转为字符型。
    yg_fl.km1001(xlapp_flag, SInfo_df, fl_list, amt, jbb_bm, in_df, out_ws)


def jjr_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, out_ws):
    for k in in_df.columns:                                                     # col是列字段名(科目代码)
        fl_list = cshfl(zdr_bm, yyb_bm)                                         # 初始化分录各字段list列表
        amount = in_df[k]['合计']                                               # 从左至右取每列/字段的最后一行的合计数据
        if (amount != 0) and (not (pd.isnull(amount))):
            amount = round(in_df[k]['合计'], 2)                                 # 对合计数据四舍五入，并转为字符型。
            if k == '224107':
                s = in_df[k]                                                    # 获取'224107'这一列pd.Serie这个对象 
                
                tc_list = ['A1小计', 'A2小计', 'A小计', 'B小计', 'D小计', 'L小计', '合计']            # 对这个pd.Serie进行瘦身过滤下
                s = s[~s.index.isin(tc_list)]                                   # 对这个pd.Serie过滤掉ti-list  ~是对True或False逻辑值取反
                amount = s
            jjr_fl.switcher(jjr_fl.dict, xlapp_flag, k, fl_list, amount, jbb_bm, out_ws)

    # 支付结算分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第四次、初始化分录各字段list列表
    amt = round(in_df['1001']['合计'], 2)                                       # 对合计数据四舍五入，并转为字符型。
    jjr_fl.km1001(xlapp_flag, SInfo_df, fl_list, amt, jbb_bm, out_ws)

             
def sb_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, sf_df, dz_df, out_ws):
    # 一、 单位部分社保分录
    for k in in_df.columns:                                                     # col是列字段名(科目代码)
        fl_list = cshfl(zdr_bm, yyb_bm)                                         # 第一次、初始化分录各字段list列表
        amount = in_df[k]['成本数据']                                            # 从左至右取每列/字段的最后一行的合计数据
        if (amount != 0) and (not (pd.isnull(amount))):
            amount = round(in_df[k]['成本数据'], 2)                              # 对合计数据四舍五入，并转为字符型。
            sb_fl.switcher(sb_fl.dict_1, xlapp_flag, k, fl_list, amount, jbb_bm, in_df, dz_df, out_ws)

    # 二、 个人部分社保分录
    for k in in_df.columns:                                                     # col是列字段名(科目代码)
        fl_list = cshfl(zdr_bm, yyb_bm)                                         # 第二次、初始化分录各字段list列表
        amount = in_df[k]['成本数据']                                            # 从左至右取每列/字段的最后一行的合计数据
        if (amount != 0) and (not (pd.isnull(amount))):
            amount = round(in_df[k]['成本数据'], 2)                              # 对合计数据四舍五入，并转为字符型。
            sb_fl.switcher(sb_fl.dict_2, xlapp_flag, k, fl_list, amount, jbb_bm, in_df, dz_df, out_ws)

    # 三、代垫总部分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第三次、初始化分录各字段list列表
    sb_fl.km_dzgr(xlapp_flag, fl_list, dz_df, out_ws)

    # 四、应收应付分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第三次、初始化分录各字段list列表
    sb_fl.km_sfsb(xlapp_flag, fl_list, sf_df, out_ws)

    # 五、银行付款这一笔分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第四次、初始化分录各字段list列表
    sb_fl.km1001(xlapp_flag, SInfo_df, fl_list, in_df, out_ws)


def gjj_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, sf_df, dz_df, out_ws):
    # 一、 单位、个人公积金分录
    for k in in_df.columns:                                                     # col是列字段名(科目代码)
        fl_list = cshfl(zdr_bm, yyb_bm)                                         # 第一次、初始化分录各字段list列表
        amount = in_df[k]['成本数据']                                           # 从左至右取每列/字段的最后一行的合计数据
        if (amount != 0) and (not (pd.isnull(amount))):
            amount = round(in_df[k]['成本数据'], 2)                              # 对合计数据四舍五入，并转为字符型。
            gjj_fl.switcher(gjj_fl.dict, xlapp_flag, k, fl_list, amount, jbb_bm, in_df, dz_df, out_ws)

    # 二、代垫总部分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第三次、初始化分录各字段list列表
    gjj_fl.km_dzgr(xlapp_flag, fl_list, dz_df, out_ws)

    # 三、应收应付分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第二次、初始化分录各字段list列表
    gjj_fl.km_sfgjj(xlapp_flag, fl_list, sf_df, out_ws)

    # 四、银行付款这一笔分录
    fl_list = cshfl(zdr_bm, yyb_bm)                                             # 第三次、初始化分录各字段list列表
    gjj_fl.km1001(xlapp_flag, SInfo_df, fl_list, in_df, out_ws)


xlapp_flag = "win32com"                                                         # xlwings 最慢，它是在win32com基础上的封装
# xlwings 用 out_ws.range("A3").options(index=False, header=False).value = out_df 最后一次性在excel添加多行数据、即是out_df累加list数据这种方式，
# 将会在执行switcher()碰到1001科目代码时啥也没有做，out_df继续保持空df，于是接着执行km1001()时传入out为n空df参数，从而out.append（即即空df.append）报append（即none没有append属性的错误。

def out_xls(xlapp_flag, yyb_bm, pz):
    out_xlfile = "output\\" + yyb_bm + "_" + pz + ".xlsx"
    copyfile("docs\\template.xlsx", out_xlfile)
    if xlapp_flag == "win32com":
        from win32com.client import Dispatch
        xlapp = Dispatch("Excel.Application")
        xlapp.Visible = False
        xlapp.DisplayAlerts = False
        out_xlfile = os.path.abspath(out_xlfile)                                # win32不认识相对路径，故需转换为绝对路径。
        out_wb = xlapp.Workbooks.Open(out_xlfile)
        out_ws = out_wb.ActiveSheet
    if xlapp_flag == "openpyxl":
        from openpyxl import load_workbook
        xlapp = ''
        out_wb = load_workbook(out_xlfile)
        out_ws = out_wb.active
    if xlapp_flag == "xlwings":
        import xlwings as xw
        xlapp = xw.App(visible=False)
        out_wb = xw.Book(out_xlfile)
        out_ws = out_wb.sheets[0]
        # out_wb = xlapp.books.open(out_xlfile)
        # out_ws = out_wb.sheets.active

    return out_xlfile, xlapp, out_wb, out_ws


def out_save(xlapp_flag, xlapp, wb, out_xlfile):
    if xlapp_flag == "win32com":
        wb.Save()
        wb.Close()
        xlapp.Quit()
    if xlapp_flag == "xlwings":
        wb.save()
        wb.close()
        xlapp.kill()
    if xlapp_flag == "openpyxl":
        wb.save(filename=out_xlfile)
        wb.close()
        
        # 再用wim32com打开保存下，这样用友才能识别
        from win32com.client import Dispatch
        xlapp = Dispatch("Excel.Application")
        xlapp.Visible = False
        xlapp.DisplayAlerts = False
        out_xlfile = os.path.abspath(out_xlfile)                                # win32不认识相对路径，故需转换为绝对路径。
        wb = xlapp.Workbooks.Open(out_xlfile)
        wb.Save()                                                               # wb.SaveAs(xlfile)
        wb.Close()
        xlapp.Quit()


def convert_yg(zdr_bm, yyb_bm, jbb_bm, in_df):
    out_xlfile, xlapp, out_wb, out_ws = out_xls(xlapp_flag, yyb_bm, "yg")
    yg_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, out_ws)
    out_save(xlapp_flag, xlapp, out_wb, out_xlfile)
    

def convert_jjr(zdr_bm, yyb_bm, jbb_bm, in_df):
    out_xlfile, xlapp, out_wb, out_ws = out_xls(xlapp_flag, yyb_bm, "jjr")
    jjr_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, out_ws)
    out_save(xlapp_flag, xlapp, out_wb, out_xlfile)


def convert_sb(zdr_bm, yyb_bm, jbb_bm, in_df, sf_df, dz_df):
    out_xlfile, xlapp, out_wb, out_ws = out_xls(xlapp_flag, yyb_bm, "sb")
    sb_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, sf_df, dz_df, out_ws)
    out_save(xlapp_flag, xlapp, out_wb, out_xlfile)


def convert_gjj(zdr_bm, yyb_bm, jbb_bm, in_df, sf_df, dz_df):
    out_xlfile, xlapp, out_wb, out_ws = out_xls(xlapp_flag, yyb_bm, "gjj")
    gjj_add_fl(zdr_bm, yyb_bm, jbb_bm, in_df, sf_df, dz_df, out_ws)
    out_save(xlapp_flag, xlapp, out_wb, out_xlfile)


if __name__ == '__main__':
    pass