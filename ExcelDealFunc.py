# -*- coding: utf-8 -*-

import pandas as pd
import time

def exceldealfunc(filename1,filename2,filename3,filename4,filename5,str1,str2):
    """
        函数名：exceldealfunc(filename1,filename2,filename3,filename4,filename5,num1,num2)
        函数功能：执行表1表2取公共部分,并生成公共集表3,表1的去除公共集的表4,表2的去除公共集的表5
            输入	1: filename1：读取表1的文件路径
            输入	1: filename2：读取表2的文件路径
            输入	1: filename3：写入表3的文件路径
            输入	1: filename4：写入表4的文件路径
            输入	1: filename5：写入表5的文件路径
            输入	2: str1：表1要比较的列号名
            输入	2: str2：表2要比较的列号名
            输出	1: 无
        其他说明：无
    """
    start=time.time()

    table1=pd.read_excel(filename1)     # 读取表1
    print("Time:%d s,Read Table1 Successful!"%(time.time()-start))
    table2=pd.read_excel(filename2)     # 读取表2
    print("Time:%d s,Read Table2 Successful!"%(time.time()-start))
    print()

    set1=set(table1[str1])              # 将表1要比较的列转换为集合格式(集合1)
    print("Time:%d s,File1 turn into a set Successful!"%(time.time()-start))
    set2=set(table2[str2])              # 将表2要比较的列转换为集合格式(集合2)
    print("Time:%d s,File2 turn into a set Successful!"%(time.time()-start))
    print()

    set3=set1 & set2                    # 取集合1和集合2的交集set3
    print("Time:%d s,set3 = set1 & set2 Successful!"%(time.time()-start))
    list3=list(set3)                    # 将set3转换为列表的格式
    print("Time:%d s,set3 turn to a list Successful!"%(time.time()-start))
    table1[table1[str1].isin(list3)].to_excel(filename3,index=False,)   # 将交集保存到表3
    print("Time:%d s,Write to Table3 Successful!"%(time.time()-start))
    print()

    list4=list(set1 - set3)             # 取集合1和交集的差集,并转换为列表格式
    print("Time:%d s,(set1 - set3) turn to a list Successful!"%(time.time()-start))
    table1[table1[str1].isin(list4)].to_excel(filename4,index=False,)   # 将差集1保存到表4
    print("Time:%d s,Write to Table4 Successful!"%(time.time()-start))
    print()

    list5=list(set2 - set3)             # 取集合2和交集的差集,并转换为列表格式
    print("Time:%d s,(set2 - set3) turn to a list Successful!"%(time.time()-start))
    table2[table2[str2].isin(list5)].to_excel(filename5,index=False,)   # 将差集2保存到表5
    print("Time:%d s,Write to Table5 Successful!"%(time.time()-start))
    print()

    print("All task finish! Using Time:%d s."%(time.time()-start))
