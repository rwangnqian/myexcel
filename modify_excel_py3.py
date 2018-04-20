"""
Copyright (C) 2016-2018
All rights reserved.
License: modify_excel
create_time:
_author_:wangqian
"""
import pandas as pd
import random
import numpy as np
import xlwings as xw
import os
import time
import random
##################################################
"""
pandas只能新建excel文件并进行写入并不能把表格写入到已经存在的表格中，
所以我写了如下的类文件把表格写入到已经存在的excel文件中去

# #     sht.range('A1').options(transpose=True).value=list(range(1,18))
# #     sht.range('A1').value=['name','sex','customer','wag']
# #     sht.range('e1').options(transpose=True).value=[random.randint(300,1090) for i in range(50)]
# #     sht.range('d2').value='yy'
# #     rng=sht.range('a1').current_region
# #     rng.autofit()


'''
class使用说明：
1.首先必须要实例化对象
2.必须要定义操作excel表格的函数（注意必须设置两个参数）
3.调用实例方法进行excel操作



# 定义对excel表格的操作函数。必须注意必须要有两个参数一个是workbook一个是worksheet
# def cellfunc(mbooks,sht):
#     sht.range('e10').value='wangqian'

# a=modify_excel_cells('d:\\test\\test5.xlsx','er')
# 实例化对象
# a.modify_excel(cellfunc)
# 调用实例的修改功能，修改存在excel中的具体操作。使用的语法是xlwings


'''


"""

##################################################


class modify_excel_cells:
    def __init__(self,filepath,tablename):
        self.filepath=filepath
        self.tablename=tablename
    def open_excel(self,sflag=None):
        if os.path.exists(self.filepath):
            mbooks=xw.Book(self.filepath)
            for a in mbooks.sheets:
                if self.tablename in a.name:
                    sflag=True    # 判断是否存在
            if sflag is True:
                sht=mbooks.sheets[self.tablename]
                sht.activate()
            else:
                sht=mbooks.sheets.add(self.tablename)
        else:
            mbooks=xw.Book()
            sht=mbooks.sheets.add(self.tablename)
        return mbooks,sht,sflag

    def modify_excel(self,func):
        '''
        这个函数是主要用于处理存在的sheet，想把对sheets操作的部分单独用一个函数引入
        参数引入工作薄和工作表，经过处理将保存好的工作薄和工作表返回给函数主体。
        当然可以先用
        :param func:mbooks,sht
        :return:mbooks,sht
        '''
        mbooks,sht,sflag=self.open_excel()
        if  not sflag:
            print ('该表格不存在请重新检查')
            exit()
        else:
            # 尝试把新的操作用函数独立出去
            def cellfunc(mbooks=mbooks,sht=sht):
                func(mbooks,sht)
                return mbooks,sht
            cellfunc()
            sht.activate()
            rng=sht.range('a1').current_region
            rng.autofit()
            mbooks.save()
            mbooks.close()

    def insert_new_sheet(self,table=None):
        mbooks,sht,sflag=self.open_excel()
        sht.activate()
        sht.clear()
        sht.range('a1').value=table
        rng=sht.range('a1').current_region
        rng.autofit()
        if sflag is True:
            mbooks.save()
            mbooks.close()
        else:
            mbooks.save(path=self.filepath)
            mbooks.close()





