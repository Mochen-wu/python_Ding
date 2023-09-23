# python_Ding
#python钉钉下载图专用
#源代码如下





#本源码代码适用于python3.7以后的版本
#本源码创建始于python3.7，经过测试，python3.9正常使用，其他版本请自行测试
#本源码代码适用于python3.7以后的版本
#本源码创建始于python3.7，经过测试，python3.9正常使用，其他版本请自行测试



print("======================尘染出品======================")
print("======================尘染出品========================")
print("==============请遵守相关法律法规下使用本源码==============")
print("========可根据需求自定义更改以上源码，符合用户需求==========")
print("====不得利用该py脚本去做出违反有损害法律、网络安全行为=======")
print("===使用过程中若出现一切问题，将由使用者自行承担一切法律责任===")
print("=======使用过程中若出现一切问题，与该py脚本开发者无关=======")
print("=======使用过程中若出现一切问题，与该py脚本开发者无关=======")
print("======================尘染出品========================")
print("======================尘染出品========================")
print("=================该代码为尘染本人原创=================")
print("===================转载需注明来源=====================")
print("==============请遵守相关法律法规下使用本源码============")
print("======================尘染出品======================")
print("======================尘染出品======================")




#使用前需安装相关的库方可使用，若已安装便不需要重复下载
#windows系统
# pip isntall requests
# pip install xlrd
#Mac系统下需要把pip换成pip3方可安装，其他不变
# pip3 isntall requests
# pip3 install xlrd





#导入库
import requests
import os
import xlrd
import tkinter as tk
from tkinter import filedialog

#打开
sudo = tk.Tk()
sudo.withdraw()


print("请选择下载至目录")
print("请选择下载至目录")
print("请选择下载至目录")
Folderpath = filedialog .askdirectory() #获得选择好的文件夹

print("请选择表格文件（xls文件）")
print("请选择表格文件（xls文件）")
print("请选择表格文件（xls文件）")
Filepath = filedialog .askopenfilename()#获得选择好的处理后的表格文件
bglujing = '%s'% Filepath
#导入需要处理的文件
xl = xlrd.open_workbook(r'%s' % bglujing )
table = xl.sheets()[0]
col = table.col_values(0)
i=0

while i<29:
     url = table.cell(i+1, 6).value
     mz = table.cell(i+1, 2).value+table.cell(i+1, 0).value+ url[-4:]
     dirpath="%s" % Folderpath
     print(dirpath)
     savedir=(dirpath+"/ "+mz)
     r = requests.get(url)
     with open(savedir,"wb") as code:
          code.write(r.content)
     print(url)
     print(mz)

     url2 = table.cell(i+1, 7).value
     mz2 = '-'+table.cell(i+1, 2).value+table.cell(i+1, 0).value+ url[-4:]
     dirpath="%s" % Folderpath
     print(dirpath)
     savedir=(dirpath+"/ "+mz2)
     r = requests.get(url2)
     with open(savedir,"wb") as code:
          code.write(r.content)
     print(url2)
     print(mz2)

     if i == 29:
          break
     i += 1
     
     
print("===看到此行文字就说明运行程序完毕了，请打开文件夹查看文件=====")
print("===看到此行文字就说明运行程序完毕了，请打开文件夹查看文件=====")

print("======================尘染出品========================")
print("======================尘染出品========================")
print("==============请遵守相关法律法规下使用本源码==============")
print("========可根据需求自定义更改以上源码，符合用户需求==========")
print("====不得利用该py脚本去做出违反有损害法律、网络安全行为=======")
print("===使用过程中若出现一切问题，将由使用者自行承担一切法律责任===")
print("=======使用过程中若出现一切问题，与该py脚本开发者无关=======")
print("=======使用过程中若出现一切问题，与该py脚本开发者无关=======")
print("======================尘染出品========================")
print("======================尘染出品========================")
print("==============请遵守相关法律法规下使用本源码============")
print("==============请遵守相关法律法规下使用本源码============")
