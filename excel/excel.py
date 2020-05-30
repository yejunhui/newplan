import xlrd ,openpyxl
import os

sheelTitle = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','aa','ab','ac','ad','ae','af','ag','ah','ai','aj','ak','al','am','an','ao','ap','ar','as','at','au','av','aw','ax','ay','az']

#Excel文件处理类
class excel:
    def __init__(self):
        pass

    #读取Excel
    def readExcel(self,readFile):
        # try:
            #Excel数据
            excelData = {}
            #Excel表数据
            excelSheetData = []
            #openFile
            e = xlrd.open_workbook(readFile)
            #获取工作表数量
            sheetQuantity = len(e.sheets())
            sheeetNames = e.sheet_names()
            #读取表内容
            for n in sheeetNames:
                #读取表
                s = e.sheet_by_name(n)
                #获得总行
                nrows = s.nrows
                #获得总列
                ncol = s.ncols
                for j in range(ncol):
                    excelData[sheelTitle[j]] = s.col_values(j)


            return excelData


        # except:
        #     print('There seems to be a problem!--> -->')

    #写入Excel
    def writeExcel(self,data):
        #Open File
        wb = openpyxl.Workbook()
        #crate Title
        Bom = wb.create_sheet(index=0,title='BOM')
        Publish = wb.create_sheet(index=1,title='Publish')
        Pass = wb.create_sheet(index=3,title='Pass')
        In = wb.create_sheet(index=4,title='In')
        Demand = wb.create_sheet(index=5,title='Demand')
        SysDemand = wb.create_sheet(index=6,title='SysDemand')

        for d in data :
            Bom.append(d)
            Publish.append(d)
            #Pass.append(d)
            In.append(d)
            Demand.append(d)
            SysDemand.append(d)

        row = 1
        for d in data :
            Pass.append(d)