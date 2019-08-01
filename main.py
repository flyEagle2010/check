
# main.py
# -*- coding: utf-8 -*-
from tkinter import *
import pandas as pd
import os
# import os
import shutil
import time
# from data import Excel
# import sys
# sys.path.append(r'E: \project\mtt2')


files = []
path = r'E:\project\src\mtt'


def loadFile():
    for filename in os.listdir(path):
        print(os.path.join(path, filename))
        files.append(filename)


def SetExcelStyle(df, fileName):
    print("set style filename:", fileName)

    writer = pd.ExcelWriter(fileName, engine='xlsxwriter')

    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format(
        {'bg_color': '#CCCCCC', 'font_color': '#ff0006'})

    worksheet.write_string(10, 10, "Total", format1)

    writer.save()


def CreateExcel(file):
    localtime = time.asctime(time.localtime(time.time()))
    print("本地时间为 :", localtime)
    print("time:", time.strftime("%Y%m%d", time.localtime()))
    shutil.copyfile(file, "./data/"+time.strftime("%Y%m%d",
                                                  time.localtime())+"-"+file.split("/")[-1])


def CreateResult(file):
    print("file:", file)
    pd.set_option('expand_frame_repr', False)
    df = pd.read_excel(file, sheet_name="Sheet1")

    print("df: ", df.dropna(axis=1, how='all'))
    print("df col row:", df.iterrows(), "col:", df.columns, df.rows)

    for index, row in df.itmerows():
        for col_name in df.columns:
            print("row[col_name]:", row[col_name])

        #     for i in range(len(df[0])):
        #         print("i:", i)

    SetExcelStyle(df, file)


loadFile()
root = Tk()
root.title("测试测试")

# 是x 不是*
root.geometry('400x300')
root.resizable(width=False, height=False)


frame = Frame(root)
frame.pack()


fline = Frame(frame)
fline.pack()
Label(fline, text='选择文件').pack(side='left')

v = StringVar(fline)
OptionMenu(fline, v, *files).pack(side='right')
v.set(files[0])


def createFile():
    print(v.get())
    CreateExcel(path+'/'+v.get())


def compareResult():
    file = "./data/"+time.strftime("%Y%m%d", time.localtime())+"-"+v.get()
    print(file)
    CreateResult(file)


btn_create = Button(frame, text="生成文件", width=15, command=createFile)
btn_create.pack()
Button(frame, text="比对结果", width=15, command=compareResult).pack()

# 进入消息循环
root.mainloop()
