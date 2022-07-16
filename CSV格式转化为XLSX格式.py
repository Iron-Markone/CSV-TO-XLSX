#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import os.path
import win32com.client as win32


# In[2]:
hint = input('温馨提示,运行前请关闭所以excel文件,是否关闭：')

def csv2xlsx():
    rootdir = str(input('需要打开的文件夹目录：'))
    rootdir1 = str(input('需要存放的文件夹目录：'))
#     rootdir = r"path_file" #需要转换的xls文件存放处
#     rootdir1 = r"path_file1" #转换好的xlsx文件存放处
    files = os.listdir(rootdir) #列出xls文件夹下的所有文件
    num = len(files) #列出所有文件的个数
    for i in range(num): #按文件个数执行次数
        kname = os.path.splitext(files[i])[1] #分离文件名与扩展名，返回(f_name, f_extension)元组
#         if kname == '.xls': #判定扩展名是否为xls,屏蔽其它文件
#             fname = rootdir + '\\' + files[i] #合成需要转换的路径与文件名
#             fname1 = rootdir1 + '\\' + files[i] #合成准备存放转换好的路径与文件名
#             excel = win32.gencache.EnsureDispatch('Excel.Application') #调用win32模块
#             wb = excel.Workbooks.Open(fname) #打开需要转换的文件
#             wb.SaveAs(fname1+"x", FileFormat=51) #文件另存为xlsx扩展名的文件
#             wb.Close()
#             excel.Application.Quit()
        if kname == '.csv': #判定扩展名是否为csv,屏蔽其它文件
            fname = rootdir + '\\' + files[i] #合成需要转换的路径与文件名
            fname1 = rootdir1 + '\\' + files[i] #合成准备存放转换好的路径与文件名
            excel = win32.gencache.EnsureDispatch('Excel.Application') #调用win32模块
            wb = excel.Workbooks.Open(fname) #打开需要转换的文件
            wb.SaveAs(fname1.replace('csv', 'xlsx'), FileFormat=51) #文件另存为xlsx扩展名的文件
            wb.Close()
            excel.Application.Quit()


# In[3]:


csv2xlsx()


# In[ ]:




