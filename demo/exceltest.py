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



# 创建两个对象
myExcel=easyExcel("C:\\Users\\mosi\\Desktop\\excel测试\\text01.xls") #实例化一个easyExcel类
title =easyExcel("C:\\Users\\mosi\\Desktop\\箱梁.xls") #实例化一个easyExcel类 ,来读取箱涵表里面的桩号与天数





# 创建总文件夹
filenamed = "张拉"
path =  "C:\\Users\\mosi\\Desktop\\" + filenamed
folder = os.path.exists(path)
if not folder:#判断是否存在文件夹如果不存在则创建为文件夹
    os.makedirs(path)#makedirs 创建文件时如果路径不存在会创建这个路径
    print("---  new folder...  ---")
    print("---  OK  ---")

else:
    print("---  There is this folder!  ---")


# 创建分类文件夹
filenames = ['第十七跨','第十八跨','第十九跨','第二十跨','第二十一跨','第二十二跨','第二十三跨']
for filename in filenames:
    path =  "C:\\Users\\mosi\\Desktop\\张拉\\" + filename
    folder = os.path.exists(path)
    if not folder:#判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)#makedirs 创建文件时如果路径不存在会创建这个路径
        print("---  new folder...  ---")
        print("---  OK  ---")

    else:
        print("---  There is this folder!  ---")




# 从箱涵读取文件， 填入模板文件，保存为新的文件
nums = 4
xhline1 = 2
xhline2 = 20
xhline3 = 15

count = 0

for x in range(4,56):
    titlename=title.getCell(1,x,xhline1)#获取单元格内容
    print("获取桩号:%s"%titlename)
    cqtime=title.getCell(1,x,xhline2)
    print("获取龄期:%s"%cqtime)
    zltime = title.getCell(1,x,xhline3)# 获取张拉日期
    print("获取张拉日期:%s"%zltime)
    # 将桩号信息填入报审表
    myExcel.setCell(1,7,5,titlename)
    print("桩号填入报审")

    # 张拉记录信息填入
    line1 = 13 # line1 line2表示行高s
    line2 = 14
    #填入标定日期
    myExcel.setCell(2,4,20,zltime)
    # 填入龄期
    chiqi = str(cqtime)[0] + str(cqtime)[1]  + '天'
    myExcel.setCell(2,8,22,chiqi)


    for x in range(0,8):
        # 获取两个随机浮点数
        a = random.uniform(1, 2)
        b = random.uniform(2, 3)
        c = random.uniform(24, 25)# 控制应力随机数字
        d = random.uniform(24, 26)# 控制应力随机数字
        #控制随机数的精度round(数值，精度)
        num1 = round(a, 1)
        num2 = round(b, 1)
        num3 = num1*2 - 0.2
        num4 = num2*2 - 0.2
        # A端申长量
        num5 = random.randint(12,22)
        # A端申长量
        num6 = random.randint(20,30)
        # A端申长量(二倍初应力)

        x1 = random.randint(-3,1)
        x2 = random.randint(1,3)
        x3 = random.randint(-3,1)
        x4 = random.randint(1,3)


        num7 = num5 *2 - random.randint(x1,x2)
        # A端申长量(二倍初应力)
        num8 = num6 *2 - random.randint(x3,x4)
        # 控制应力A端
        num9 = round(c,1)
        # 控制应力B端
        num10 = round(d,1)
        num11 = random.randint(107,120)
        num12 =  random.randint(107,120)
        print(num1,num2,num3,num4)
        print(num5,num6,num7,num8)
        print(num9,num10,num11,num12)
        myExcel.setCell(2,line1,4,num1)  #修改单元格内容，第一个参数是sheet的编号，第二个为行数，第三个为列数，（全部都以1开始，下面的xlrd那几个模块都以0开始的），最后是要修改的内容
        myExcel.setCell(2,line1,5,num2)
        myExcel.setCell(2,line1,6,num3)
        myExcel.setCell(2,line1,7,num4)
        myExcel.setCell(2,line2,4,num5)
        myExcel.setCell(2,line2,5,num6)
        myExcel.setCell(2,line2,6,num7)
        myExcel.setCell(2,line2,7,num8)
        myExcel.setCell(2,line1,13,num9)
        myExcel.setCell(2,line1,15,num10)
        myExcel.setCell(2,line2,13,num11)
        myExcel.setCell(2,line2,15,num12)




        '''
        此处做一个判断，如果生成的参数不符合要求，则从新运行上面的方法
        '''

        # 获取到偏差值
        result = myExcel.getCell(2,line1,21)
        print("-"*50)
        print("-"*50)

        print(result)

        print("-"*50)
        print("-"*50)



        # 第一次问题数据使用随机数字进行重写
        if result <= -6 or result >= 6:
            a = random.uniform(1, 2)
            b = random.uniform(2, 3)
            c = random.uniform(24, 25)# 控制应力随机数字
            d = random.uniform(24, 26)# 控制应力随机数字
            #控制随机数的精度round(数值，精度)
            num1 = round(a, 1)
            num2 = round(b, 1)
            num3 = num1*2 - 0.2
            num4 = num2*2 - 0.2
            # A端申长量
            num5 = random.randint(12,22)
            # A端申长量
            num6 = random.randint(20,30)
            # A端申长量(二倍初应力)

            x1 = random.randint(-3,1)
            x2 = random.randint(1,3)
            x3 = random.randint(-3,1)
            x4 = random.randint(1,3)


            num7 = num5 *2 - random.randint(x1,x2)
            # A端申长量(二倍初应力)
            num8 = num6 *2 - random.randint(x3,x4)
            # 控制应力A端
            num9 = round(c,1)
            # 控制应力B端
            num10 = round(d,1)
            num11 = random.randint(107,120)
            num12 =  random.randint(107,120)
            print(num1,num2,num3,num4)
            print(num5,num6,num7,num8)
            print(num9,num10,num11,num12)
            myExcel.setCell(2,line1,4,num1)  
            myExcel.setCell(2,line1,5,num2)
            myExcel.setCell(2,line1,6,num3)
            myExcel.setCell(2,line1,7,num4)
            myExcel.setCell(2,line2,4,num5)
            myExcel.setCell(2,line2,5,num6)
            myExcel.setCell(2,line2,6,num7)
            myExcel.setCell(2,line2,7,num8)
            myExcel.setCell(2,line1,13,num9)
            myExcel.setCell(2,line1,15,num10)
            myExcel.setCell(2,line2,13,num11)
            myExcel.setCell(2,line2,15,num12)


            result = myExcel.getCell(2,line1,21)


            # 筛选后对还存在问题的数据进行人工填写
            if result <= -6 or result >= 6:

                myExcel.setCell(2,line1,4,1.2)  
                myExcel.setCell(2,line1,6,2.2)
                myExcel.setCell(2,line1,7,5.4)
                myExcel.setCell(2,line2,4,20)
                myExcel.setCell(2,line2,5,20)
                myExcel.setCell(2,line2,6,37)
                myExcel.setCell(2,line2,7,39)
                myExcel.setCell(2,line1,13,24.1)
                myExcel.setCell(2,line1,15,25.2)
                myExcel.setCell(2,line2,13,109)
                myExcel.setCell(2,line2,15,117)






        # 循环完修改列高参数
        line1 +=2
        line2 +=2





# 压浆记录信息填入
    # k表示行高
    k =10
    for x in range(0,8):
        num = random.randint(1,2)
        num3 = random.randint(1,9)
        num2 = str(0.1) + str(num) + str(num3)
        content = float(num2)
        print(content)
        myExcel.setCell(4,k,19,content)
        k +=1


    # 外循环参数
    nums +=1




    # 此处给一个判断，从文件名中获取相关参数进行判断，然后存入对应的文件夹

    # n = 几跨
    n = titlename[1] + titlename[2]
    print('*'*50)
    print(n)
    print('*'*50)


    if n == "17":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第十七跨\\" + titlename + '.xls'
    elif n == "18":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第十八跨\\" + titlename + '.xls'
    elif n == "19":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第十九跨\\" + titlename + '.xls'
    elif n == "20":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第二十跨\\" + titlename + '.xls'
    elif n == "21":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第二十一跨\\" + titlename + '.xls'
    elif n == "22":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第二十二跨\\" + titlename + '.xls'
    elif n == "23":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第二十三跨\\" + titlename + '.xls'
    elif n == "24":
      filename = "C:\\Users\\mosi\\Desktop\\张拉\\第二十三跨\\" + titlename + '.xls'




   # filename = "C:\\Users\\mosi\\Desktop\\张拉\\" + titlename + '.xls'
    myExcel.save(filename) #保存文件，如果路径与打开时相同，即保存文件，如果不同即新建文件
    #myExcel.close()

    count +=1

end = time.clock()
print('运行时间: %s 秒'%(end-start))
print('总共生成文件:%s 个'%count)