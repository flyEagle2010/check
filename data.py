
# import pandas as pd
# import os
# import shutil
# import time


# class Excel:
#     def __init__(self):
#         print("init....")

#     def CreateExcel(file):
#         localtime = time.asctime(time.localtime(time.time()))
#         print("本地时间为 :", localtime)
#         print("time:", time.strftime("%Y%m%d", time.localtime()))
#         shutil.copyfile(file, "./data/"+time.strftime("%Y%m%d",
#                                                       time.localtime())+"-"+file.split("/")[-1])

#     def CreateResult(file):
#         df = pd.read_excel("./data/" + file, sheet_name=0)
#         SetExcelStyle(df, file)
#         print("df: ", df.dropna(axis=1, how='all'))
#         # for i in range(len(df[0])):
#         #     print("i:", i)

#     def SetExcelStyle(df, fileName):
#         writer = pd.ExcelWriter("setcolor" + fileName)

#         sheet = writer.sheets['Sheet1']

#         bold = df.add_format({
#             'bold':  True,  # 字体加粗
#             'border': 1,  # 单元格边框宽度
#             'align': 'left',  # 水平对齐方式
#             'valign': 'vcenter',  # 垂直对齐方式
#             'fg_color': '#F4B084',  # 单元格背景颜色
#             'text_wrap': True,  # 是否自动换行
#         })
#         sheet.write(10, 10, "+++", bold)
#         writer.save()
#         writer.close()
