# -*- coding: utf-8 -*-
"""
Created on Sun Jun 17 10:48:13 2018

@author: star
"""

import pyttsx3
import time
import xlrd,xlwt
class PaperScore():
    def __init__(self):
        self.msg="日照香炉生紫烟，遥看瀑布挂前川。飞流直下三千尺，疑是银河落九天"
        self.ans="ABBCD DDACD BDAAC CCBBD ABBCD DDACD BDAAC CCBBD  BDAAC CCBBD"
        self.engine = pyttsx3.init()
        self.scoreList=[]
    
    '''
    播报答案选项
    '''
    def speechAnswer(self):       
        for i in range(10):
            #msg1=msg[i*8:(8*i)+8]
            if i%2==0:
                bh=str(int(i/2*10+1))+'至'+str(int(i/2*10+10))+'题'
                self.engine.say(bh)
            msg1=self.ans[i*6:(6*i)+6]
            # print(msg1)
            self.engine.say(msg1)
            time.sleep(1)
            self.engine.runAndWait()
    '''
    输入错误个数，每个2分，获得最终成绩，并播报结果
    '''        
    def calcScore(self):  
        errCount=input("请输入错误个数，输入0打印分数：")
        if errCount=='0':
            print("结束运行！打印成绩")
            #讲输入的成绩写入excel表格
            self.printScore()
            return
        if errCount=='':          
            print("请重新输入")
            self.calcScore()
            return
        score=100-int(errCount)*2
        print(score)
        self.scoreList.append(score)
        res1="错误"+errCount+"个，得分"+str(score)
        self.engine.say(res1)
        self.engine.runAndWait()     
        self.calcScore()
    '''
    打印成绩
    '''
    def printScore(self):
        self.writeExcel("成绩.csv",self.scoreList)
        '''
        for i,score in enumerate(self.scoreList):
            print(i+1,score )
            self.engine.say('成绩为'+str(score))
            time.sleep(2)
            self.engine.runAndWait()
        '''
        return
    '''
    核对成绩
    '''
    def checkScore(self):
        book=xlrd.open_workbook('成绩.csv')
        sheet=book.sheet_by_name('sheet1')
        rows=sheet.nrows
        cols=sheet.ncols
        for r in range(1,rows):
            for c in range(0,cols):
                values=sheet.cell_value(r,c)
                self.engine.say(values)
                self.engine.runAndWait()
                
        return
    #excel写入
    def writeExcel(self,filename,data):
        book=xlwt.Workbook()
        sheet=book.add_sheet('sheet1')
        c=1
        sheet.write(0,1,'成绩')
        for d in data:
            sheet.write(c,1,str(d))   
            c+=1
        book.save(filename)
        '''
    #excel读取
    def reddExcel(filename):
        book=xlrd.open_workbook(filename)
        sheet=book.sheet_by_name('sheet1')
        rows=sheet.nrows
        cols=sheet.ncols
        for c in range(cols):
            c_values=sheet.col_values(c)
            print(c_values)
        for r in range(rows):
            r_values=sheet.row_values(r)
            print(r_values)
        print(sheet.cell(1,1))
        '''
        
if __name__=='__main__':
    paper=PaperScore()
    while True:
        fun=input("0、退出程序\n1、播报答案\n2、计算成绩\n3、播报成绩\n4、核对姓名成绩\n请选择程序:")
        if fun=='1':
            paper.speechAnswer()
        elif fun=='2':
            paper.calcScore()
        elif fun=='3':
            paper.printScore()
        elif fun=='4':
            paper.checkScore()
        elif fun=='0':
            break
        else:
            print("输入数据有误，请重新输入！")
    
        



