import xmind
import time
import xlwings as excelM
# -*- coding:utf-8 -*-

class xmind2excel():


    def xmindtoexcel(self,xminds,excels):
        excelApp = excelM.App(False, False)
        excelFile = excelApp.books.add()   # 新增一个文件
        sht =excelFile.sheets.add('测试用例')# 新增一个表格
        now = time.strftime('%Y%m%d_%H%M%S', time.localtime())#获取本地当前时间



        workbook = xmind.load(xminds)#加载xmind文件
        #workbook = xmind.load('‪C:/Users/supaur/Desktop/xminds/中心主题.xmind')  # 加载xmind文件
        data=workbook.getData()#获取xmind文件数据
        print(data)

        titlelist = ['所属模块','用例标题','前置条件','步骤','预期','优先级']

        sht.range('A1').value = titlelist


        for i in (2,2):#往excel写进一行行的数值
            for model in data[0]["topic"]["topics"]:
                print(model["title"])
                for title in model["topics"]:
                    sht.range('B{}'.format(i)).options(transpose=True).value = title["title"]
                    print(title["title"])
                    for preconditions in title["topics"]:
                        sht.range('C{}'.format(i)).options(transpose=True).value = preconditions["title"]
                        print(preconditions["title"])
                        for step in preconditions["topics"]:
                            sht.range('D{}'.format(i)).options(transpose=True).value = step["title"]
                            print(step["title"])
                            for expect in step["topics"]:
                                sht.range('E{}'.format(i)).options(transpose=True).value = expect["title"]
                                print(expect["title"])
                                for priority in expect["topics"]:
                                    sht.range('A{}'.format(i)).options(transpose=True).value = model["title"]
                                    sht.range('F{}'.format(i)).options(transpose=True).value = priority["title"]
                                    print(priority["title"])
                                i = i + 1

        excelFile.save(r"{}/{}.xlsx".format(excels,now))#保存excel
        print("文件写入完成，文件路径"+ excels)
        #excelFile.save(r"C:/Users/supaur/Desktop/excels/{}.xlsx".format(now))  # 保存excel
