import sys
import xlwt
import time
import configparser
import pandas as pd
from pandas import DataFrame


def transform(arg):
    sec_ls = []
    opt_ls = []
    values_ls = []
    sec_a_ls = []

    conf = configparser.ConfigParser()
    conf.read("%s" % arg, encoding="utf8")
    sections = conf.sections()

    for sec in sections:
        _sec = "[" + sec + "]"
        sec_a_ls.append(sec)
        sec_ls.append(_sec)
        a = conf.options(sec)
        opt_ls.extend(a)
    opt_ls = list(set(opt_ls))

    for sec in sec_a_ls:
        for opt in opt_ls:
            try:
                values = conf.get(sec, opt)
                if values == " ":
                    values_ls.append(" ")
                values_ls.append(values)
            except:
                values_ls.append("无")

    ll = [values_ls[i:i + len(opt_ls)] for i in range(0, len(values_ls), len(opt_ls))]

    def collect():
        time_str = time.strftime("%Y/%m/%d") + " " + time.strftime("%I:%M")
        writer = pd.ExcelWriter("%s.xls" % (arg + "-report"))
        df = DataFrame({"扫描日期: %s" % time_str: opt_ls})

        for i in range(len(sec_ls)):
            sec = sec_ls[i]
            df[sec] = ll[i]
        print(df)
        df.to_excel(writer, sheet_name='sheet1')
        writer.save()
    return collect


def main():
    try:
        arg = sys.argv[1]
        print("开始扫描.")
        target = transform(arg)
        target()
        print("扫描完成.")
    except:
        print("输入提示: srs [infile]  [outfile]")


if __name__ == "__main__":
    main()
