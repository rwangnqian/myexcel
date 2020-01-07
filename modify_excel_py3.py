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

#回过头来发现以前用类的方法来对excel进行修改太麻烦，代码写的太渣。
#应该使用装饰器就可以了
#使用范例
@tables_to_sheets(locate=r"G:\my_new.xlsx", sheetname="bb")
def tt(app, wb, sht):
    sht.range('c20').options(transpose=True).value = ["hfd", "11", "er", "24", "op", "asjk"]

tt()

'''


"""

##################################################

def tables_to_sheets(locate, sheetname):
    # 该装饰器主要用于将表格或者list等数据存到指定的sheet中
    # 如果xlsx文件不存在则创建文件，如果sheet不存在则创建sheet。
    # 缺点是xlsx文件没有办法设定编码。
    def outer_sheet(func):
        def inner_sheet(*args, **kwargs):
            sheetlist=[]
            app = xw.App(visible=True, add_book=False)
            app.display_alerts = True
            app.screen_updating = False
            if not os.path.exists(locate):
                print("excel文件不存在，正在创建文件")
                app = xw.App(visible=True, add_book=False)
                wb = app.books.add()
                wb.save(locate)
                wb.close()
                app.quit()
            print("excel已经存在，正在打开。。。")
            wb = app.books.open(locate)
            for a in wb.sheets:
                sheetlist.append(a.name)
            if sheetname not in sheetlist:
                wb.sheets.add(sheetname)
            sht = wb.sheets[sheetname]
            sht.activate()
            func(app=app, wb=wb, sht=sht)
            rng = sht.range('a1').current_region
            rng.autofit()
            wb.save()
            wb.close()
            app.quit()
            return func
        return inner_sheet
    return outer_sheet



