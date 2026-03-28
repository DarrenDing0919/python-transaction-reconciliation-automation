import pandas as pd
import openpyxl as op
import numpy as np
import os

def find_file(root):
    files = os.listdir(root)
    ret = []
    for file in files:
        if '.xlsx' in file and 'closing' not in file and 'opening' not in file:
            ret.append(file)
    ret.sort(key=lambda a: int(a[-6]) if len(a)==17 else int(a[-7:-5]))
    return ret

def handle(file,sheet_name):
    a = np.arange(65, 91)
    list_ABC = [chr(i) for i in a]
    dict1 = {}
    for i in zip(range(0, 27), list_ABC):
        dict1[i[1]] = i[0]
    list_code = []
    list_Date = []
    df1 = pd.read_excel(file)

    df2 = pd.read_excel(save_file)
    for i in zip(df1['Code'], df1['Date']):
        # print(i)

        list_code.append('0' * (6 - len(str(i[0]))) + str(i[0]))
        list_Date.append(i[1])
    list_DVP = []
    list_RVP = []
    list_Amount_D = []
    list_Amount_R = []
    for i in zip(df1['Code'], df1['Type'], df1['Quantity'], df1['Amount']):
        if i[1] == 'DVP':
            list_DVP.append(i[2])
            list_Amount_D.append(i[3])
            list_RVP.append('')
            list_Amount_R.append('')
        elif i[1] == 'RVP' or i[1] == 'BNS':
            list_RVP.append(i[2])
            list_Amount_R.append(i[3])
            list_DVP.append('')
            list_Amount_D.append('')
    list_all = []
    for idx, i in enumerate(list_code):
        list1 = ['' for i in range(1, df2.shape[1])]
        list1[dict1['A']] = list_code[idx]
        list1[dict1['C']] = list_Date[idx]
        list1[dict1['L']] = list_DVP[idx]
        list1[dict1['W']] = list_Amount_D[idx]
        list1[dict1['G']] = list_RVP[idx]
        list1[dict1['H']] = list_Amount_R[idx]

        # print(list1)

        list_all.append(list1)
    table = op.load_workbook(save_file)
    table1 = table[sheet_name]
    for i in range(3, len(list_DVP) + 3):
        for j in range(1, df2.shape[1]):
            # print(list_all[i-3][j-1])
            table1.cell(i, j, list_all[i - 3][j - 1])
    table.save(save_file)


if __name__ == "__main__":
    in_put = 'input' # 这里填input文件夹的地址
    save_file = 'output/trans_result_WA.xlsx'
    # 上面一行填子文件夹output中trans_result_WA.xlsx的地址
    files = find_file(in_put)
    for i in range(len(files)):
        sh_name = 'transaction'+str(i+1)
        path = os.path.join(in_put, files[i])
        handle(path, sh_name)