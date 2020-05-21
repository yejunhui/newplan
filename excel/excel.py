import xlrd ,openpyxl
import os

#Excel文件处理类
class excel:
    def __init__(self,readFile,writeFile):
        self.thisFile = os.getcwd()
        self.readFile = readFile
        self.wriFile = writeFile

    #读取Excel
    def readExcel(self):
        try:
            #Excel数据
            excelData = {}
            #Excel表数据
            excelSheetData = []
            data = []
            #openFile
            e = xlrd.open_workbook(self.readFile)
            #获取工作表数量
            sheetQuantity = len(e.sheets())
            sheeetNames = e.sheet_names()
            #读取表内容
            for n in sheeetNames:
                #读取表
                s = e.sheet_by_name(n)
                #获得总行
                nrows = s.nrows
                for j in range(nrows):
                    data.append(s.row(j))
                excelSheetData.append(data)
                data = []
                excelData[n] = excelSheetData
                excelSheetData = []

            return excelData


        except:
            print('There seems to be a problem!--> -->')

    #写入Excel
    def writeExcel(self):
        #Open File
        wb = openpyxl.Workbook()
        #crate Title
        wb.create_sheet(index=0,title='BOM')
        wb.create_sheet(index=1,title='Publish')
        wb.create_sheet(index=3,title='Pass')
        wb.create_sheet(index=4,title='In')
        wb.create_sheet(index=5,title='Demand')
        wb.create_sheet(index=6,title='SysDemand')

