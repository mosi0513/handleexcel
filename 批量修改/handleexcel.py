5#!/usr/bin/env python
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

        

# 创建总文件夹
def create_folder(folder_name,folder_path):

    path = folder_path + folder_name

    folder = os.path.exists(path)
    if not folder:
        os.makedirs(path)
        print("---  create new folder%s...  ---"%path)
    else:
        print("---  the folder is exist!  ---")


# 创建子文件夹
def create_child_folder(folder_names,folder_path):
    
    for folder_name in folder_names:
            path = folder_path + "\\" + folder_name
            folder = os.path.exists(path)
            if not folder:
                os.makedirs(path)
                print("---  create new folder%s...  ---"%path)
            else:
                print("---  the folder is exist!  ---")


# 将随机数值填入单元格
def write_to_excel(write_excel,line1,line2):
    num1 = float(str(1) + "." + str(random.randint(1,9)))
    num2 = float(str(2) + "." + str(random.randint(1,9)))

    random_add1 = float(str(num1*2 - float(str(0) + "." + str(random.randint(0,3))))[0:3])
    random_sub1 = float(str(num1*2 + float(str(0) + "." + str(random.randint(0,3))))[0:3])
    result1 = [random_add1,random_sub1]
    num3 = random.choice(result1)

    random_add2 = float(str(num2*2 - float(str(0) + "." + str(random.randint(0,3))))[0:3])
    random_sub2 = float(str(num2*2 + float(str(0) + "." + str(random.randint(0,3))))[0:3])
    result2 = [random_add2,random_sub2]
    num4 = random.choice(result2)

    num5 = float(str(2) + str(random.randint(4,5)) + "." + str(random.randint(2,9)))
    num6 = float(str(2) + str(random.randint(4,5)) + "." + str(random.randint(2,9)))

    num7 = random.randint(12,16)
    num8 = random.randint(16,30)

    num9 = num7*2 - random.randint(-3,3)
    num10 = num8*2 - random.randint(-3,3)

    num11 = random.randint(107,120)
    num12 = random.randint(107,120)

    write_excel.setCell(2,line1,4,num1)  
    write_excel.setCell(2,line1,5,num2)
    write_excel.setCell(2,line1,6,num3)
    write_excel.setCell(2,line1,7,num4)
    write_excel.setCell(2,line1,13,num5)
    write_excel.setCell(2,line1,15,num6)
    write_excel.setCell(2,line2,4,num7)
    write_excel.setCell(2,line2,5,num8)
    write_excel.setCell(2,line2,6,num9)
    write_excel.setCell(2,line2,7,num10)
    write_excel.setCell(2,line2,13,num11)
    write_excel.setCell(2,line2,15,num12)
    print("---  写入张拉信息:  ---")
    print(num1,num2,num3,num4,num5,num6,num7,num8,num9,num10,num11,num12)


#  处理sheet2  写入多个单元格内容
def write_data(write_excel):
    line1 = 13
    line2 = 14
    for count in range(0,8):

        write_to_excel(write_excel,line1,line2)

        result = write_excel.getCell(2,line1,21)
        if result <= -6 or result >= 6:
            print("---  出现坏数据，第一次重写...  ---")
            write_to_excel(write_excel,line1,line2)

            result = write_excel.getCell(2,line1,21)
            if result <= -6 or result >= 6:
                print("---  出现坏数据，第二次重写...  ---")  
                write_to_excel(write_excel,line1,line2)
                result = write_excel.getCell(2,line1,21)
                if result <= -6 or result >= 6:
                    print("---  出现坏数据，第三次手动重写...  ---")  
                    write_excel.setCell(2,line1,4,1.6)  
                    write_excel.setCell(2,line1,5,2.3)
                    write_excel.setCell(2,line1,6,3.5)
                    write_excel.setCell(2,line1,7,4.6)
                    write_excel.setCell(2,line1,13,24.6)
                    write_excel.setCell(2,line1,15,25.8)
                    write_excel.setCell(2,line2,4,15)
                    write_excel.setCell(2,line2,5,26)
                    write_excel.setCell(2,line2,6,27)
                    write_excel.setCell(2,line2,7,51)
                    write_excel.setCell(2,line2,13,107)
                    write_excel.setCell(2,line2,15,118)

        line1 +=2
        line2 +=2

# 从excel中获取数据
def get_data(filename,start_number,end_number,*args):

    # excel对象
    excelobject =easyExcel(filename)

    #  获取的内容   strat是excel里的行数(1,2,3,4,5...)
    count = 0
    for x in range(start_number,end_number):
        #  序号
        serial_number = excelobject.getCell(args[0],start_number,args[1])
        #  桩号 
        mileage = excelobject.getCell(args[0],start_number,args[2])
        #  型号
        model = excelobject.getCell(args[0],start_number,args[3])
        #  方量
        quantity = excelobject.getCell(args[0],start_number,args[4])


        #  如果序号存在，则将数据填入到指定的文件的单元格内
        if mileage:
            write_excel = easyExcel("C:\\Users\\mosi\\Desktop\\text01.xls")

            # sheet1
            write_excel.setCell(1,7,5,mileage)
            print("---  写入桩号:%s  ---"%mileage)

            # sheet2
            write_data(write_excel)

            # sheet3
            row =10
            for x in range(0,8):
                num = random.randint(1,2)
                num3 = random.randint(1,9)
                num2 = str(0.1) + str(num) + str(num3)
                content = float(num2)
                write_excel.setCell(4,row,19,content)
                print("---  写入压浆记录:%s  ---"%content)
                row +=1

            filename = "C:\\Users\\mosi\\Desktop\\磨斯\\" + str(mileage) + '.xls'
            write_excel.save(filename)
            print("create new file: %s"%filename)
            write_excel.close()
        else:
            return None
        count +=1
        start_number +=1
        print("-"*60)
        print("总共生成文件:%s 个"%count)



def run():

    # 创建总文件夹
    floder_name = "测试新建总文件夹"
    current_path = "C:\\Users\\mosi\\Desktop\\"
    create_folder(floder_name,current_path)


    # 创建子文件夹
    child_folder_names = ['子文件夹1','子文件夹2','子文件夹3','子文件夹4']
    current_path = "C:\\Users\\mosi\\Desktop\\测试新建总文件夹"
    create_child_folder(child_folder_names,current_path)


    # 从表中获取内容
    filename = "C:\\Users\\mosi\\Desktop\\箱梁.xls"
    #  sheet：第一个单元格 num代表列数(表内ABCDEFG...)
    sheet = 1
    num1 = 1
    num2 = 2
    num3 = 3
    num4 = 4
    get_data(filename,3,64,sheet,num1,num2,num3,num4)


 # 检测偏差值
def test_data():
       
    file_dir = r'C:\\Users\\mosi\\Desktop\\磨斯'
    count = 0 
    faitnums = 0
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
    return count,faitnums




if __name__ == '__main__':

    start =time.clock()
    run()
    count,faitnums = test_data()
    end = time.clock()
    print('总共检查文件：%s 个'%count)
    print("问题数据数量：%s"%faitnums)
    print('运行时间: %s 秒'%(end-start))
    print("-"*60)








