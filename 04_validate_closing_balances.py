import pandas as pd
import numpy as np
import math
import sys


def concat_sheets(df):
    sheet_names = list(df.keys())
    all_df = pd.DataFrame()
    for sheet_name in sheet_names:
        df_sheet = df[sheet_name]
        all_df = pd.concat([all_df, df_sheet])
    return all_df


if __name__ == "__main__":

    print_list_1 = []
    print_list_2 = []
    print_list_3 = []
    out_file_1 = 'wrong_with_closing.xlsx'
    out_file_2 = 'ending_shares_negative.xlsx'
    out_file_3 = 'ending_cost_negative.xlsx'

    trwa_file = "output/trans_result_WA.xlsx" # 这里填trans_result_WA.xlsx的地址
    closing_file = "input/closing.xlsx" # 这里填closing.xlsx的地址
    df_ret = pd.read_excel(
        trwa_file,
        dtype={'Code': 'str', 'A': 'str'},
        skiprows=[0, 1],
        usecols="A,B,C,D,E,F,G,H,I,J,K,L,M,N,O",
        header=None,
        names=["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"],
        sheet_name=None)
    df_ret = concat_sheets(df_ret)  # 多个sheet_name 合并
    df_c = pd.read_excel(
        closing_file,
        dtype={'A': 'str'},
        skiprows=[0],
        usecols="A, B, C",
        header=None,
        names=["A", "B", "C"])

    df_n = df_ret[df_ret['N'] < 0]
    df_n.to_excel(out_file_2, index=False)

    df_n = df_ret[df_ret['O'] < 0]
    df_n.to_excel(out_file_3, index=False)

    # o = df_ret.loc[:, "O"]
    # for k in o:
    #     if str(k) >= str(0):
    #         pass
    #     else:
    #         print_list_3.append(str(k))
    #
    #     with open(out_file_3, "w") as f:
    #         for k in print_list_3:
    #             f.write(k + '\n')
    #     f.close()

    pcode, prow = df_ret.iloc[0]["A"], 0
    k, all = 0, len(df_ret)
    for id, row in df_ret[1:].iterrows():
        k = k + 1
        print("\r", end="")
        print("进度: {}%: ".format(k / all * 100), "▓" * (k * 20 // all), end="")
        sys.stdout.flush()
        icode = row['A']

        if pcode == icode:  # 与前一行相同
            pass
        else:  # 与前一行不相同，前一行为最后一个
            row_c = df_c[df_c["A"] == pcode]
            if len(row_c) == 0:  # 找不到到相同code行
                # print('无相同code', pcode)
                pass
            else:  # 找到相同code行
                row_c = row_c.iloc[0]
                if (prow["N"] == row_c["B"]) and ((math.fabs((prow["O"] - row_c["C"])) <= row_c["C"] * 0.01) or (
                        math.fabs((prow["O"] - row_c["C"])) <= 10)):
                    # print("通过自检", pcode)
                    pass
                else:
                    # print("不通过自检", pcode, prow["N"], row_c["B"], math.fabs((prow["O"] - row_c["C"])), row_c["C"] * 0.01)

                    print_list_1.append(pcode)
                    print_list_1.append(str(prow["N"]))
                    print_list_1.append(str(row_c["B"]))
                    print_list_1.append(str(math.fabs((prow["O"] - row_c["C"]))))
                    print_list_1.append(str(row_c["C"] * 0.01))

        pcode, prow = icode, row

    # 最后一个数据
    row_c = df_c[df_c["A"] == pcode]
    if len(row_c) == 0:  # 找不到到相同code行
        # print('无相同code', pcode)
        pass
    else:  # 找到相同code行
        row_c = row_c.iloc[0]
        if (prow["N"] == row_c["B"]) and ((math.fabs((prow["O"] - row_c["C"])) <= row_c["C"] * 0.01) or (
                math.fabs((prow["O"] - row_c["C"])) <= 10)):
            # print("通过自检", pcode)
            pass
        else:
            # print("不通过自检", pcode, prow["N"], row_c["B"] , math.fabs((prow["O"] - row_c["C"])), row_c["C"]*0.01)
            print_list_1.append(pcode)
            print_list_1.append(str(prow["N"]))
            print_list_1.append(str(row_c["B"]))
            print_list_1.append(str(math.fabs((prow["O"] - row_c["C"]))))
            print_list_1.append(str(row_c["C"] * 0.01))

    # 保存

    with open(out_file_1, "w") as f:
        i = 0
        for code in print_list_1:
            i = i + 1
            f.write(code + ' ')
            if i % 5 == 0:
                f.write('\n')
    f.close()




