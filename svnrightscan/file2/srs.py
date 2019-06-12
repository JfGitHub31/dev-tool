import sys,os
import configparser
import pandas as pd
import xlwt
from pandas import DataFrame


"""
脚本运行
python3 srs.py 参数, 结果生成一个excel表格
注意: 脚本和文件在同一目录下
"""


class SRS:
    sec_ls = []
    opt_ls = []
    values_ls = []

    def fun(self):
        arg = sys.argv[1]
        print(arg,type(arg))
        conf = configparser.ConfigParser()
        file_path = conf.get
        conf.read("%s" % arg, encoding="utf8")
        sections = conf.sections()

        for sec in sections:
            sec1 = "[" + sec + "]"
            self.sec_ls.append(sec1)
            a = conf.options(sec)
            self.opt_ls.extend(a)

        # print(sec_ls,len(sec_ls))  # 73个首选项
        opt1_ls = list(set(self.opt_ls))
        # print(opt1_ls,len(opt1_ls))  # 12 个配置参数

        for index in range(len(opt1_ls)):
            for sec in sections:
                try:
                    values = conf.get(sec,opt1_ls[index])
                    self.values_ls.append(values)
                except:
                    self.values_ls.append("None")
        ll = [self.values_ls[i:i + len(self.sec_ls)] for i in range(0, len(self.values_ls), len(self.sec_ls))]
        # print(ll)

        def run():
            writer = pd.ExcelWriter("2.xls")
            df = DataFrame({"sections_name": self.sec_ls})
            for i in range(len(opt1_ls)):
                opt = opt1_ls[i]
                df[opt] = ll[i]
            print(df)

            df.to_excel(writer, sheet_name='sheet1')
            writer.save()
        return run


def main():
    srs = SRS()
    S = srs.fun()
    S()

main()


if __name__ == "__main__":
    main()
