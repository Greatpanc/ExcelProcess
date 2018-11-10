# -*- coding: utf-8 -*-
import xlrd,xlsxwriter
import sys

def writeoneline(table_from,table_to,line,nrows):
    """
        函数名：writeoneline(table_from,table_to,line,nrows)
        函数功能：读取table_from的第line行数据写入到table_to的第nrows行
            输入1: table_from：从该表内读取要写入的数据
            输入2: table_to：向该表内添加一行数据
            输入3: line：table_from内要读取数据的行数
            输入4: nrows：table_to内要写入数据的行数
            输出1: 无
        其他说明：无
    """
    write_data=table_from.row_values(line)
    for i in range(len(write_data)):
        table_to.write(nrows,i,write_data[i])

def readdata(filename,num):
    """
        函数名：readdata(filename,num)
        函数功能：通过文件路径,打开Excel文件,读取sheet
            输入	1: filename：读取表的文件路径
            输入	2: num：要比较的列号
            输出	1: table：sheet表信息
            输出	2: nrows：该Excel表的行数
            输出	3: count：count要比较列的数据列表
        其他说明：无
    """
    data = xlrd.open_workbook(filename)
    table = data.sheets()[0]
    nrows = table.nrows
    count=table.col_values(num)

    return table,nrows,count

def exceldealfunc(filename1,filename2,filename3,filename4,filename5,num1,num2):
    """
        函数名：exceldealfunc(filename1,filename2,filename3,filename4,filename5,num1,num2)
        函数功能：执行表1表2取公共部分,并生成公共集表3,表1的去除公共集的表4,表2的去除公共集的表5
            输入	1: filename1：读取表1的文件路径
            输入	1: filename2：读取表2的文件路径
            输入	1: filename3：写入表3的文件路径
            输入	1: filename4：写入表4的文件路径
            输入	1: filename5：写入表5的文件路径
            输入	2: num1：表1要比较的列号
            输入	2: num2：表2要比较的列号
            输出	1: 无
        其他说明：无
    """
    table1,nrows1,count1=readdata(filename1,num1)
    print("Read Table1 Successful!")
    table2,nrows2,count2=readdata(filename2,num2)
    print("Read Table2 Successful!")

    data3 = xlsxwriter.Workbook(filename3)    # 共同用户样本
    table3 = data3.add_worksheet()

    data4 = xlsxwriter.Workbook(filename4)    # Table1处理后的样本Table4
    table4 = data4.add_worksheet()

    data5 = xlsxwriter.Workbook(filename5)    # Table2处理后的样本Table5
    table5 = data5.add_worksheet()

    writeoneline(table2,table3,0,0)
    writeoneline(table1,table4,0,0)
    writeoneline(table2,table5,0,0)
    nrows3=1
    nrows4=1
    nrows5=1

    for i in range(1,nrows1):
        if count1[i]  in count2:
            writeoneline(table1,table3,i,nrows3)                # Table1写入共同用户样本Table3
            nrows3+=1
        else:
            writeoneline(table1,table4,i,nrows4)                # Table1写入处理后的样本Table4
            nrows4+=1
        if i%10000==0:
            done=i/(nrows1+nrows2)
            sys.stdout.write("\r[%s%s] %d%%" % ('█'*int(40*done),'  '*(40-int(40*done)),int(100*done)))
            sys.stdout.flush()
    data3.close()
    print("Write to Table3 Successful!")
    data4.close()
    print("Write to Table4 Successful!")

    for i in range(1,nrows2):
        if count2[i] not in count1:
            writeoneline(table2,table5,i,nrows5)                # table2写入处理后的样本Table5
            nrows5+=1
        if i%10000==0:
            done=(i+nrows2)/(nrows1+nrows2)
            sys.stdout.write("\r[%s%s] %d%%" % ('█'*int(40*done),'  '*(40-int(40*done)),int(100*done)))
            sys.stdout.flush()
    data5.close()
    print("Write to Table5 Successful!")


