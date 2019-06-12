import sys
import xlwt
import time
import configparser
import pandas as pd
from pandas import DataFrame


class SRS(object):

    def __init__(self, infile, outfile):
        self.infile = infile
        self.outfile = outfile

    def tranform(self):
        sec_ls = []
        sec_a_ls = []
        opt_ls = []
        values_ls = []

        conf = configparser.ConfigParser()
        conf.read("%s" % self.infile, encoding="utf8")
        sections = conf.sections()

        for sec in sections:
            _sec = "[" + sec + "]"
            sec_ls.append(_sec)
            sec_a_ls.append(sec)
            a = conf.options(sec)
            opt_ls.extend(a)

        opt_ls = list(set(opt_ls))

        for opt in opt_ls:
            for sec in sec_a_ls:
                try:
                    values = conf.get(sec, opt)
                    print(values,type(values),len(values))
                    if values == "":
                        values_ls.append("无")
                    elif values != "":
                        values_ls.append(values)
                    else:
                        values_ls.append(values)
                except:
                    values_ls.append("")

        ll = [values_ls[i:i + len(sec_a_ls)] for i in range(0, len(values_ls), len(sec_a_ls))]
        print(ll, len(ll))
        def collect():
            time_str = time.strftime("%Y/%m/%d") + " " + time.strftime("%I:%M")
            writer = pd.ExcelWriter("%s.xls" % (self.outfile + "-report"))
            df = DataFrame({"%s" % ("扫描日期: " + time_str): sec_ls})
            for i in range(len(opt_ls)):
                opt = opt_ls[i]
                df[opt] = ll[i]

            print(df)

            df.to_excel(writer, sheet_name='sheet1')
            writer.save()

        return collect


def main():
    try:
        infile = sys.argv[1]
        outfile = sys.argv[2]
        print("开始扫描.")
        s = SRS(infile, outfile)
        srs = s.tranform()
        srs()
        print("扫秒完成.")
    except:
        print("输入提示: srs [infile]  [outfile]")


if __name__ == "__main__":
    main()
