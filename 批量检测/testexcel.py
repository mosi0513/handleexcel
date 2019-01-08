#!/usr/bin/env python
# -*- coding: utf-8 -*-
from win32com.client import Dispatch
import win32com.client
import random
import os
import time

class easyExcel:
    """A utility to make it easier to get at Excel.  Remembering
    to save the data is your problem, as is  error handling.
    Operates on one workbook at a time."""
    def __init__(self, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''
    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()
    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp
    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value
    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value
    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)
    def cpSheet(self, before):
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None,shts(1))

        


start =time.clock()

#  这里需要获取路径名称，然后进行拼接遍历
file_dir = r'C:\\Users\\mosi\\Desktop\\磨斯'



# 检查问题数据

faitnums = 0
count = 0 
for root, dirs, files in os.walk(file_dir):  
    # print(root) #当前目录路径  
    # print(dirs) #当前路径下所有子目录  
    # print(files) #当前路径下所有非目录子文件
    for file in files:
        filename = root + "\\" + file

        print(filename)
        myExcel=easyExcel(filename)

        k =13 
        for x in range(0,8):
            titlename=myExcel.getCell(2,k,21)#获取单元格内容
            k +=2
            if int(titlename) >= 6 or int(titlename)<=-6:
                faitnums +=1

            print(titlename)
        myExcel.close()

        count +=1

end = time.clock()


print('检查时间: %s 秒'%(end-start))
print('总共检查文件：%s 个'%count)
print("问题数据数量：%s"%faitnums)

