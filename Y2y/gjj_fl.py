import pandas as pd
import re
from tkinter import messagebox
from yg_fl import append


def km660111(xlapp_flag, k, fl_list, amt, jbb, in_df, dz_df, out_ws):
    if '准本分小计' in in_df.index:
        amt1 = amt - in_df[k]['准本分小计']                                     # 本分+应付
    else:
        amt1 = amt
    amt = str(amt)                                                              # 本分+应付+准本分
    amt1 = str(round(amt1, 2))
    fl_list[5] = '计提单位公积金 (本分+应付)'
    fl_list[6] = k
    fl_list[8] = amt1
    fl_list[9] = amt1
    fl_list[15] = jbb + ':部门'
    fl_list[16] = '01:人员类别'
    append(xlapp_flag, out_ws, fl_list)

    km_dzdw(xlapp_flag, k, fl_list, dz_df, out_ws)

    fl_list[5] = '计提单位公积金 (本分+应付+准本分)'
    fl_list[6] = '221104'
    fl_list[8] = None
    fl_list[9] = None
    fl_list[11] = amt
    fl_list[12] = amt
    fl_list[15] = '01:人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)

    fl_list[5] = '扣缴单位公积金 (本分+应付+准本分)'
    fl_list[6] = '221104'
    fl_list[8] = amt
    fl_list[9] = amt
    fl_list[11] = None
    fl_list[12] = None
    fl_list[15] = '01:人员类别'
    fl_list[16] = None
    append(xlapp_flag, out_ws, fl_list)


def km22410404(xlapp_flag, k, fl_list, amt, jbb, in_df, dz_df, out_ws):
    amt = str(amt)                                                              # 本分+应付,传入的“成本数据”amt，社保表中个人部分本身就是“本分+应付”，无需像上面单位部分样减除“准本分”。
    fl_list[5] = '应扣个人公积金（本分+应付）'
    fl_list[6] = k
    fl_list[8] = amt
    fl_list[9] = amt
    append(xlapp_flag, out_ws, fl_list)


def km_dzdw(xlapp_flag, k, fl_list, dz_df, out_ws):                             # 代垫总部单位社保
    s = dz_df[k]                                                                # 获取k即各项社保这一列pd.Serie这个对象
    n = 0
    for i, v in s.items():                                                      # items是pd.Serie的每个项目，i是键(df的index)也即'代垫'这一列、v是值(k这列上的数据)
        i = i[4:]                                                               # 截取“代垫总部”之后内容
        b1 = i.find(":")                                                        # 查找:的位置
        if b1 == -1:
            i1 = None
        else:
            i1 = i[:b1]                                                         # 截取人员编码(自第5个字符至:间的字符)
        b2 = re.search(r"[(,（,\u4e00-\u9fa5]", i)                              # 匹配(、（或中文字符
        if b2:
            b2 = b2.span()[0]
        i2 = i[b1+1:b2]                                                         # 截取部门编码

        if v and (not (pd.isnull(v))):
            fl_list[5] = '计提单位公积金' + ' (代垫总部' + '[' + dz_df['xm'].iloc[n] + ']单位部分）'
            fl_list[6] = k
            fl_list[8] = str(round(v, 2))
            fl_list[9] = str(round(v, 2))
            fl_list[11] = None
            fl_list[12] = None
            fl_list[15] = i2 + ':部门'
            fl_list[16] = '01:人员类别'
            append(xlapp_flag, out_ws, fl_list)
        n += 1


def km_dzgr(xlapp_flag, fl_list, dz_df, out_ws):                                # 代垫总部个人社保
    s = dz_df['22410404']                                                       # 获取k即各项社保这一列pd.Serie这个对象
    n = 0
    for i, v in s.items():                                                      # items是pd.Serie的每个项目，i是键(df的index)也即'代垫'这一列、v是值(k这列上的数据)
        i = i[4:]                                                               # 截取“代垫总部”之后内容
        b1 = i.find(":")                                                        # 查找:的位置
        if b1 == -1:
            i1 = None
        else:
            i1 = i[:b1]                                                         # 截取人员编码(自第5个字符至:间的字符)
        b2 = re.search(r"[(,（,\u4e00-\u9fa5]", i)                              # 匹配(、（或中文字符
        if b2:
            b2 = b2.span()[0]
        i2 = i[b1+1:b2]                                                         # 截取部门编码

        if v and (not (pd.isnull(v))):
            if i1 is None:
                messagebox.showinfo(title='提示', message="准本分区个人部分有数据，却没有人员编码，生成的凭证不完整！！")
                return
            fl_list[5] = '应收总部代扣' + ' (代垫总部' + '[' + dz_df['xm'].iloc[n] + ']个人公积金)'
            fl_list[6] = '12211901'
            fl_list[8] = str(round(v, 2))
            fl_list[9] = str(round(v, 2))
            fl_list[11] = None
            fl_list[12] = None
            fl_list[15] = i1 + ':人员档案'
            append(xlapp_flag, out_ws, fl_list)
        n += 1


def km_sfgjj(xlapp_flag, fl_list, sf_df, out_ws):                               # 代垫社保
    s = sf_df['1001']                                                           # 获取'660110'即社保合计数这一列pd.Serie这个对象
    n = 0
    for i, v in s.items():                                                      # items是pd.Serie的每个项目，i是键(df的index)也即'代垫'这一列、v是值('660110'这列上的数据)
        i1 = i[0:2]                                                             # 截取左边两个字符,即'应付'或'应收'二字。
        b = re.search(r"[(,（,\u4e00-\u9fa5]", i[2:])                           # 匹配(或中文字符
        if b:
            b = b.span()[0]
        i2 = i[2:][:b]                                                          # 截取第3到(或中文出现的地方，即如1109广分的编码。

        if v and (not (pd.isnull(v))):
            if i1 == '应付':
                fl_list[5] = '应付' + i2 + '部替本部[' + sf_df['xm'].iloc[n] + ']代垫公积金(应付数据)'
                fl_list[6] = '22410102'
                fl_list[8] = None
                fl_list[9] = None
                fl_list[11] = str(round(v, 2))
                fl_list[12] = str(round(v, 2))
                fl_list[15] = i2 + ':客商'
                append(xlapp_flag, out_ws, fl_list)
            if i1 == '应收':
                fl_list[5] = '应收本部替' + i2 + '部[' + sf_df['xm'].iloc[n] + ']代垫公积金(应收数据)'
                fl_list[6] = '12210102'
                fl_list[8] = str(round(v, 2))
                fl_list[9] = str(round(v, 2))
                fl_list[11] = None
                fl_list[12] = None
                fl_list[15] = i2 + ':客商'
                append(xlapp_flag, out_ws, fl_list)
        n += 1


def km1001(xlapp_flag, SInfo_df, fl_list, in_df, out_ws):                       # 银行托收
    global kmbm, yhzh
    YYB_bm = fl_list[10]
    var = round(in_df['1001']['实付数据'], 2)                                   # 社保合计列的实付数据
    if var == 0:
        print("公积金合计列实付数据为空，不能生成付款分录！")
        return

    if SInfo_df['公积金结算户'][YYB_bm] == "总部统一结算":
        kmbm = "114305"
        if var != 0:
            fl_list[5] = '支付公积金'
            fl_list[6] = kmbm
            fl_list[11] = str(var)
            fl_list[12] = str(var)
            fl_list[15] = '1101:客商'
            append(xlapp_flag, out_ws, fl_list)
    else:
        if SInfo_df['公积金结算户'][YYB_bm] == "基本户":
            kmbm = SInfo_df['基本户-科目编码'][YYB_bm]
            yhzh = SInfo_df['基本户-银行账户编码'][YYB_bm] + ':银行账户'
            if pd.isna(SInfo_df)['基本户-科目编码'][YYB_bm]:                     # pd.isna(SInfo_df)将各元素值转化为True或False
                kmbm = '1001'
                yhzh = ''
        elif SInfo_df['公积金结算户'][YYB_bm] == "专用户":
            kmbm = SInfo_df['公积金专用户-科目编码'][YYB_bm]
            yhzh = SInfo_df['公积金专用户-银行账户编码'][YYB_bm] + ':银行账户'
            if pd.isna(SInfo_df)['公积金专用户-科目编码'][YYB_bm]:                # pd.isna(SInfo_df)将各元素值转化为True或False
                kmbm = '1001'
                yhzh = ''
        elif SInfo_df['公积金结算户'][YYB_bm] == "现金":
            kmbm = '1001'
            yhzh = ''

        fl_list[5] = '支付公积金'
        fl_list[6] = kmbm
        fl_list[11] = str(var)
        fl_list[12] = str(var)
        fl_list[15] = yhzh
        append(xlapp_flag, out_ws, fl_list)


dict = {
    "660111": km660111,                 # 单位公积金
    "22410404": km22410404,             # 个人公积金
}


def switcher(dict, xlapp_flag, k, fl_list, amt, jbb, in_df, dz_df, out_ws):
    # func = dict.get(k, lambda xlapp_flag, k, fl_list, amt, jbb, in_df, dz_df, out_ws : None)
    # return func(xlapp_flag, k, fl_list, amt, jbb, in_df, dz_df, out_ws)
    dict.get(k, lambda xlapp_flag, k, fl_list, amt, jbb, in_df, dz_df, out_ws : None) \
        (xlapp_flag, k, fl_list, amt, jbb, in_df, dz_df, out_ws)