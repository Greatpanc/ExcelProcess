# -*- coding: utf-8 -*-
from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from ExcelDealFunc import exceldealfunc
import os

class MyMainFace(object):
    """主界面类"""
    def __init__(self):
        """
            函数名：__init__(self)
            函数功能：MyMainFace类的构造函数,界面组件均在此构造出的
                输入	1: MyMainFace类的对象,自身,无需输入
                输出	1: 无
            其他说明：无
        """
        self.root = Tk()
        self.root.title('Excel处理程序')

        # 请选择目标路径
        Label(self.root,text = "请选择目标路径：",fg="red").grid(row = 0, column = 0,sticky="w")

        Label(self.root,text = "表1目标路径:").grid(row = 1, column = 0,sticky="e")
        self.path1 = StringVar()
        Entry(self.root, width=60,textvariable = self.path1,state='readonly').grid(row = 1, column = 1,columnspan=3)
        Button(self.root, text = "路径选择", command = self.selectPath1).grid(row = 1, column = 4)
        self.num1 = StringVar()
        Label(self.root,text = "  比较字段的ID").grid(row = 1, column = 5,sticky="e")
        Entry(self.root, width=5,textvariable = self.num1).grid(row = 1, column = 6)

        Label(self.root,text = "表2目标路径:").grid(row = 2, column = 0,sticky="e")
        self.path2 = StringVar()
        Entry(self.root,width=60,textvariable = self.path2,state='readonly').grid(row = 2, column = 1,columnspan=3)
        Button(self.root, text = "路径选择", command = self.selectPath2).grid(row = 2, column = 4)
        self.num2 = StringVar()
        Label(self.root,text = "  比较字段的ID").grid(row = 2, column = 5,sticky="e")
        Entry(self.root, width=5,textvariable = self.num2).grid(row = 2, column = 6)

        # 请选择生成表格路径
        Label(self.root,text = "请选择生成表格路径：",fg="red").grid(row = 4, column = 0,sticky="w")

        Label(self.root,text = "表格存放路径：").grid(row = 5, column = 0,sticky="e")
        self.path3 = StringVar()
        Entry(self.root,width=60,textvariable = self.path3,state='readonly').grid(row = 5, column = 1,columnspan=3)
        Button(self.root, text = "路径选择", command = self.selectPath3).grid(row = 5, column = 4)

        # 请输入生成表格名称
        Label(self.root,text = "请输入生成表格名称：",fg="red").grid(row = 6, column = 0,sticky="w")

        Label(self.root,text = "表1表2的交集表的表名：").grid(row = 7, column = 0,sticky="e")
        self.name1 = StringVar()
        Entry(self.root, textvariable = self.name1).grid(row = 7, column = 1,sticky="w")

        Label(self.root,text = "表1去除交集后的表名:").grid(row = 8, column = 0,sticky="e")
        self.name2 = StringVar()
        Entry(self.root, textvariable = self.name2).grid(row = 8, column = 1,sticky="w")

        Label(self.root,text = "表2去除交集后的表名:").grid(row = 9, column = 0,sticky="e")
        self.name3 = StringVar()
        Entry(self.root, textvariable = self.name3).grid(row = 9, column = 1,sticky="w")

        self.labeltxt=StringVar()
        self.labeltxt.set(" ")
        Label(self.root,textvariable = self.labeltxt,fg="red").grid(row = 7, column = 3)
        Label(self.root,text = "以表1为参考",fg="red").grid(row = 7, column = 2)

        self.var = StringVar()
        self.var.set("开始")
        self.button =  Button(self.root,textvariable = self.var,command = self.start, width = 5)
        self.button.grid(row = 8,column = 2,padx = 5)

        self.var2 = StringVar()
        self.var2.set("重置")
        self.button2 =  Button(self.root,textvariable = self.var2,command = self.reset, width = 5)
        self.button2.grid(row = 9,column = 3,padx = 5)

        # 创建一个背景色为白色的矩形
        self.canvas = Canvas(self.root,width = 170,height = 26,bg = "white")
        # 创建一个矩形外边框（距离左边,距离顶部,矩形宽度,矩形高度）,线型宽度，颜色
        self.out_line = self.canvas.create_rectangle(2,2,180,27,width = 1,outline = "black") 
        self.canvas.grid(row = 8,column = 3,ipadx = 5)

        self.root.mainloop()

    def start(self):
        """
            函数名：start(self)
            函数功能：开始按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        if self.path1.get():
            filename1=self.path1.get()
        else:
            self.labeltxt.set("请选择好表1")
            return

        if self.path2.get():
            filename2=self.path2.get()
        else:
            self.labeltxt.set("请选择好表2")
            return

        if self.path3.get():
            filename=self.path3.get()
        else:
            filename="data"
        
        if self.name1.get():
            filename3=filename+"/"+self.name1.get()+".xlsx"
        else:
            filename3=filename+"/Table1_Table2.xlsx"

        if self.name2.get():
            filename4=filename+"/"+self.name2.get()+".xlsx"
        else:
            filename4=filename+"/Table1_del.xlsx"

        if self.name3.get():
            filename5=filename+"/"+self.name3.get()+".xlsx"
        else:
            filename5=filename+"/Table2_del.xlsx"

        if self.num1.get():
            num1=int(self.num1.get())
        else:
            self.labeltxt.set("请选择好num1")
            return

        if self.num2.get():
            num2=int(self.num2.get())
        else:
            self.labeltxt.set("请选择好num2")
            return

        self.button.config(state="disable") # 关闭按钮1功能
        self.root.withdraw()
        os.system('cls')
        print("正在运行中请稍等...")
        resultinf=exceldealfunc(filename1,filename2,filename3,filename4,filename5,num1,num2)
        self.root.deiconify()
        self.labeltxt.set(resultinf)

    def scheduleshow(self,i):
        fill_line = self.canvas.create_rectangle(2,2,0,27,width = 0,fill = "blue") 
        self.canvas.coords(fill_line, (0, 0, 180*i, 30))
        self.var.set(str(round(100*i,1))+"%")
        self.root.update()

    def reset(self):
        """
            函数名：reset(self)
            函数功能：重置按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        self.button.config(state="active") # 激活按钮1

        fill_line = self.canvas.create_rectangle(2,2,0,27,width = 0,fill = "white") 
        self.var.set("开始")
        self.labeltxt.set(" ")
        self.canvas.coords(fill_line, (0, 0, 181, 30))
        self.root.update()

    def selectPath1(self):
        """
            函数名：selectPath1(self)
            函数功能：选择路径1按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        path_ = askopenfilename(filetypes = [('Excel', '*.xls*')])
        self.path1.set(path_)

    def selectPath2(self):
        """
            函数名：selectPath2(self)
            函数功能：选择路径2按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        path_ = askopenfilename(filetypes = [('Excel', '*.xls*')])
        self.path2.set(path_)

    def selectPath3(self):
        """
            函数名：selectPath3(self)
            函数功能：选择路径3按钮的功能函数
                输入	1: MyMainFace类的对象,自身
                输出	1: 无
            其他说明：无
        """
        path_ = askdirectory()
        self.path3.set(path_)

##############################程序入口#########################################
if __name__=="__main__":
    MyMainFace()
