import sys
import configparser
import pandas as pd
import xlwt
from pandas import DataFrame


def transform():
    sec_ls = []
    opt_ls = []
    values_ls = []
    sec_a_ls = []
    conf = configparser.ConfigParser()
    conf.read("VisualSVN-SvnAuthz.ini", encoding="utf8")
    sections = conf.sections()

    for sec in sections:

        _sec = "[" + sec + "]"
        sec_a_ls.append(sec)
        sec_ls.append(_sec)
        a = conf.options(sec)
        opt_ls.extend(a)
    opt_ls = list(set(opt_ls))
    # print(sec_ls, len(sec_ls))
    # print(opt_ls,len(opt_ls))

    for sec in sec_a_ls:
        for opt in opt_ls:
            try:
                values = conf.get(sec, opt)
                values_ls.append(values)
            except:
                values_ls.append("None")

    ll = [values_ls[i:i + len(opt_ls)] for i in range(0, len(values_ls), len(opt_ls))]
    # print(ll,len(ll))

    def collect():
        writer = pd.ExcelWriter("123.xls")
        df = DataFrame({"opt_name": opt_ls})

        for i in range(len(sec_ls)):
            sec = sec_ls[i]
            df[sec] = ll[i]
        print(df)

        df.to_excel(writer, sheet_name='sheet1')
        writer.save()
    return collect


target = transform()
target()
