# -*- coding: utf-8 -*-
"""
Created on Sat Apr 25 10:28:58 2020

@author: an 橘
"""

import pandas as pd
import os
import time
import re



def get_path(path):
    """返回文件夹下所有文件地址：file_path列表"""
    
    g = os.walk(path)  
    file_path = []
    for path, dir_list, file_list in g:
        for file_name in file_list:
            if (re.search('xlsx$', file_name) != None) and (re.search('~\$', file_name) == None):
                # 只读xlsx文件地址，且不读取打开Excel时的临时文件
                file_path.append(os.path.join(path, file_name))
    return file_path



def read_excel(path):
    """每个Excel表存储为一个excel变量（以sheet名为键的字典）"""
    
    p = get_path(path)
    num_file = len(p)
    excel = [[] for _ in range(num_file)]
    for i,p0 in zip(range(num_file),p):
        excel[i] = pd.read_excel(p0, header=None, sheet_name=None)
    return excel



def sheets_name(excel):
    """返回所有sheet的名字"""
    
    return list(excel[0].keys())



def concat_one_sheet(excel,sheetname):
    """合并一张sheet"""
    
    sheet = []
    for i in range(len(excel)):
        sheet.append(excel[i][sheetname])
    return pd.concat(sheet)
    


def combine_sheets(excel,s_start):
    """遍历所有文件合并对应sheet"""
    """修改s_start值可以设置从哪张表开始进行合并"""
    
    sheets = sheets_name(excel)
    s = sheets[s_start-1:]
    df = []
    for si in s:
        df.append(concat_one_sheet(excel, si))    
    
    # 未提出部分工作簿不合并前，做的遍历所有sheet合并
    # for s in sheets:
    #     df.append(concat_one_sheet(excel,s))
    return df, s



def save_file(output_path,df,sheets):
    """保存输出结果"""
    
    writer = pd.ExcelWriter(output_path)
    for i,s in zip(range(len(sheets)),sheets):
        df[i].to_excel(writer, index=False,encoding='utf-8',header=None,sheet_name=s)
    writer.save()
    print("恭喜合并完成！合并结果已保存在输出文件。（文件名为“合并结果+日期+时间”）")
    


if __name__ == '__main__':
    path = input("Hello! 晓兰~\n请确保文件夹下只有需要合并的Excel文件\n第一步：粘贴Excel文件所在的文件夹地址：")
    output = input("第二步：粘贴你想要输出文件保存的文件夹：")
    start_sheet = int(input("第三步：选择从第几张子表开始实行合并（请填正整数）："))
    
    excel = read_excel(path)     # 从路径读取所有Excel文件
    dataframe, sheet_name = combine_sheets(excel,s_start=start_sheet)    # 合成后生成的DataFrame
    time = time.strftime("%Y%m%d-%H%M%S", time.localtime())  
    output_path = output + r'\合并结果' + time + '.xlsx'     # 输出文件位置
    save_file(output_path,dataframe,sheet_name)
    os.system("pause")    # 运行结束不关闭窗口
