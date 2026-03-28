import pandas as pd
import numpy as np


def handle(df, sheet):
    df = df.fillna(0)  # 无值为0
    df.rename(columns={'Unnamed: 0': "Code", 'Unnamed: 1': "Code2", 'Unnamed: 2': "Date"}, inplace=True)
    pcode, pn, po = 0, 0, 0
    for id, row in df.iterrows():
        icode = row['Code']
        d, e, f, g, h, l = row["Shares"], row["Cost"], row['average cost'], row["Shares.1"], row["Cost.1"], row[
            "Shares.3"]
        if icode != pcode:  # 第一行
            pcode = icode
        else:
            d, e = pn, po
            f = e / d if d != 0 else 0
        i = d + g
        j = e + h
        k = j / i if i != 0 else 0
        m = k * l
        n = i - l
        o = j - m
        df.loc[id, (
            'Shares', 'Cost', 'average cost', 'Shares.2', 'Cost.2', 'Average cost2', 'Cost.3', 'Shares.4',
            'Cost.4')] = (
            d, e, f, i, j, k, m, n, o)
        pn, po = n, o

    df.to_excel(writer, index=False, sheet_name=sheet)


if __name__ == "__main__":

    file = 'output/trans_result_WA.xlsx' # 这里填trans_result_WA.xlsx的地址

    df = pd.read_excel(file, skiprows=[0], dtype={'Unnamed: 0': 'str'}, sheet_name=None)
    writer = pd.ExcelWriter(file)

    for i in range(10):
        sheetname = 'transaction' + str(i + 1)
        print(sheetname)
        file = df[sheetname]
        handle(file, sheetname)
    writer.save()




