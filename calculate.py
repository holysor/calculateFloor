#-*- coding:utf-8 -*-
import numpy as np
import os
from xlrd import open_workbook
import xlwt
from xlutils.copy import copy
from Tkinter import *
import sys
import logging

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                    datefmt='%a, %d %b %Y %H:%M:%S',
                    filename='calculateFloor.log',
                    filemode='w')

console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)

def calresult(h,hn,L1,L2,L3,L4):

    try:
        max_deviation = np.max(abs(h-hn)) #最大偏差
    except:
        max_deviation = False
    count = 0
    for i in hn:
        if type(i) is np.unicode_:
            count += 1
    if count!=5:
        try:
            range = np.max(hn) - np.min(hn)
        except:
            range = False
    else:
        range = False
    try:
        L1_2 = abs(L1 - L2)
    except:
        L1_2 = False
    try:
        L3_4 = abs(L3 - L4)
    except:
        L3_4 = False
    return max_deviation,range,L1_2,L3_4
#设置表格的字体和单元格样式
def fontStyle(link=None):
    #初始化表格样式
    style = xlwt.XFStyle()
    #单元格背景颜色
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5

    #字体样式
    font = xlwt.Font()
    font.name = u'新宋体'
    font.bold = False
    font.height = 240
    font.wigth = 240
    if link == True :
        font.underline = font.UNDERLINE_SINGLE
        font.colour_index = 4

    #单元格内文字对齐方式
    cell_alignment = xlwt.Alignment()
#         cell_style.horz = cell_style.HORZ_CENTER
    cell_alignment.vert = cell_alignment.VERT_CENTER
    #单元格边框大小及宽高设定
    cell_border = xlwt.Borders()
    cell_border.top = 1
    cell_border.right = 1
    cell_border.left = 1
    cell_border.bottom = 1
#         cell_border.diag = cell_border.THIN
#         cell_border.bottom_colour = 0x3A
    style.pattern = pattern
    style.alignment = cell_alignment
    style.borders = cell_border
    style.font = font
    return style

def runCalculate():
    filename = (__file__).split('/')[-1].split('.')[0]
    logging.info(filename)

    if (filename+'.app') in os.getcwd().split(os.sep):
        pathlist = os.getcwd().split('.app')[0].split('/')[:-1]
        filepath = os.sep + os.path.join(*pathlist) + os.sep  + 'calculatevalue.xls'
    else:
        filepath = os.getcwd() + os.sep  + 'calculatevalue.xls'

    print (filename+'.app') in os.getcwd().split(os.sep)
    logging.info(filepath)

    rb = open_workbook(filepath)
    rs = rb.sheet_by_index(0)
    wb = copy(rb)
    work_sheet = wb.get_sheet(0)

    resultArry = {}
    for row in range(rs.nrows):
        h = rs.cell_value(row,0)
        hn=np.array([rs.cell_value(row,1),rs.cell_value(row,2),rs.cell_value(row,3),rs.cell_value(row,4),rs.cell_value(row,5)])
        L1 = rs.cell_value(row,6)
        L2 = rs.cell_value(row,7)
        L3 = rs.cell_value(row, 8)
        L4 = rs.cell_value(row, 9)

        result = calresult(h, hn, L1, L2, L3, L4)
        list = []
        if result[0] and type(h) is not unicode:
            work_sheet.write(row,10,int(result[0]),fontStyle())
            list.append(str(int(result[0])))

        if result[1] or str(result[1])=='0.0':
            work_sheet.write(row,11,int(result[1]),fontStyle())
            list.append(str(int(result[1])))
        if (result[2] or str(result[2])=='0.0') and (result[3] or str(result[3])=='0.0'):
            dt = str(int(result[2])) +'/'+ str(int(result[3]))
            work_sheet.write(row,12,dt,fontStyle())
            list.append(dt)
        elif result[2] or str(result[2])=='0.0':
            dt = str(int(result[2])) + '/'
            work_sheet.write(row, 12, dt,fontStyle())
            list.append(dt)
        elif result[3] or str(result[3])=='0.0':
            dt = '/' + str(int(result[3]))
            work_sheet.write(row, 12, dt,fontStyle())
            list.append(dt)
        else:
            pass
        if list:
            resultArry[row] = list
    wb.save(filepath)
    return resultArry
class mainWindow(object):
    def __init__(self):
        self.root = Tk()
        self.root.title(u'计算差值')
        # self.root.geometry("400x300")
        self.setCenter(400, 300)

        self.root.resizable(width=False, height=False)
        self.button = Button(self.root,text='开始',command=self.runCal).pack()
        self.button1 = Button(self.root,text='清空',command=self.clear).pack()
        self.label = Label(self.root,text='点击开始之前请关闭计算值表格！').pack()
        self.text = Text(self.root)
        self.text.pack(side=LEFT,fill=BOTH,expand=1)
        self.text.focus_force()

        self.scroolbar = Scrollbar(self.root)
        self.text.config(yscrollcommand=self.scroolbar.set,width=20,height=20,background='#ffffff')

        self.scroolbar.config(command=self.text.yview)
        self.scroolbar.pack(side=RIGHT, fill=Y)

        self.root.mainloop()
    def clear(self):
        self.text.delete(1.0,END)

    def setCenter(self, width, height):

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight() - 100

        self.root.update_idletasks()
        self.root.deiconify()
        self.root.withdraw()
        self.root.geometry('%sx%s+%s+%s' % (
            width + 10, height + 10, (screen_width - width) / 2,
            (screen_height - height) / 2))
        self.root.deiconify()
    def runCal(self):
        try:
            resultcal = runCalculate()
        except:
            info = sys.exc_info()
            self.text.insert(1.0, str(info[0]) + ':' + str(info[1])+'\n')
            self.text.insert(1.0, '无法正常运行(请先查看表格是否关闭)！\n')

            return
        for key,value in resultcal.items():
            if len(value)==2:
                self.text.insert(1.0,'\n表格'+str(key+1)+'行: 最大偏差='+value[0]+',极差='+value[1]+'\n')
            elif len(value)==3:
                self.text.insert(1.0, '\n表格' + str(key + 1) + '行: 最大偏差=' + value[0] + ',极差=' + value[1] + ','+'极差(净开间)='+value[2]+'\n')
            elif len(value)==1:
                self.text.insert(1.0,'\n表格' + str(key + 1) +'行: 极差(净开间)='+value[0]+'\n')
        self.text.insert(1.0, '\n===最新计算结果===\n')

mw = mainWindow()

