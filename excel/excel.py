import xlrd ,openpyxl,xlsxwriter,time
import os

sheelTitle = ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','aa','ab','ac','ad','ae','af','ag','ah','ai','aj','ak','al','am','an','ao','ap','ar','as','at','au','av','aw','ax','ay','az']

#Excel文件处理类
class excel:
    def __init__(self):
        self.arg = os.getcwd()+'/'
        self.time = time.ctime()
        '''定义计划头'''
        self.title = ['层级', 'ERP编码', '名称', '模号', 'BOM用量', '周期', '初期', '包装样式', '包装数量', '模穴类型', '用人', '工艺','项目']
        '''定义项目'''
        self.plan = ['需求', '计划生产', '实际生产', '扫描入库', '不良', '结存']
        '''定义列标，从第9列开始'''
        self.str = ['i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'aa',
                    'ab', 'ac', 'ad', 'ae', 'af', 'ag', 'ah', 'ai', 'aj', 'ak', 'al', 'am', 'an', 'ao', 'ap', 'aq',
                    'ar', 'as', 'at', 'au', 'av', 'aw', 'ax', 'ay', 'az']

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
                print(nrows)
                #获得总列
                ncol = s.ncols
                for l in range(0,nrows):
                    print(s.row_values(l))
                    excelSheetData.append(s.row_values(l))


            return excelSheetData


        # except:
        #     print('There seems to be a problem!--> -->')

    #写入推移表excel
    """写入计划excel"""

    def writePlan(self,data,newdata):
        print('开始写入推移表')
        f = xlsxwriter.Workbook('生产推移表.xlsx')
        sheet1 = f.add_worksheet('BOM')
        sheet2 = f.add_worksheet('生产发行表')
        sheet3 = f.add_worksheet('生产物料推移表')
        row = 0
        col = 0
        sheet1.write(row,col,'生产物料BOM&需求明细')
        row += 1
        for t in self.title :
            sheet1.write(row,col,t)
            col += 1
        i = 0
        for i in range(31) :
            sheet1.write(row,col,i+1)
            col += 1
        sheet1.write(row,col,'合计')
        row += 1
        col = 0
        for d in data :
            for li in d :
                print('正在写入%s'%li[1])
                #写入总成
                if '总成' in li[0] :
                    fn ='\"'+str(li[1])+'\"'
                    for l in li :
                        sheet1.write(row,col,l)
                        col += 1
                    sheet1.write(row,col,self.plan[0])
                    #sheet1.write_formula(row,col+32,'sum(j%s:an%s)'%(row+1,row+1))
                    #sheet1.write(row+1,col,self.plan[1])
                    #sheet1.write(row+2,col,self.plan[2])
                    #sheet1.write(row+3, col, self.plan[3])
                    #sheet1.write(row+4, col, self.plan[4])
                    col += 1
                    #写公式
                    i = 0
                    for i in range(32) :
                        sheet1.write_formula(row, col + i,'vlookup(%s,生产发行表!B:AZ,column(%s),0)' % (fn, self.str[i+3] + '1'))
                        """sheet1.write_formula(row + 2, col + i,'indirect(address(%s,%s))+indirect(address(%s,%s))-indirect(address(%s,%s))' % ( row + 3, col+ i, row+1, col + i+1 , row + 2, col + i+1))"""
                    sheet1.write_formula(row+1,col+31,'indirect(address(%s,%s))'%(row+3,col+31))
                    row += 1
                    col = 0
                #写入零件
                elif li[0] == '.1' :
                    name = '\"'+str(li[1])+'\"'
                    for l in li :
                        sheet1.write(row,col,l)
                        col += 1
                    sheet1.write(row,col,self.plan[1])
                    #sheet1.write(row+1,col-1,self.plan[2])
                    col += 1
                    '''写公式'''
                    i = 0
                    for i in range(32) :
                        sheet1.write_formula(row,col+i,'sumifs(%s:%s,A:A,\"总成\",B:B,%s)*INDIRECT(ADDRESS(%s,%s))'%(self.str[i+4],self.str[i+4],fn,row+1,5))
                        #sheet1.write_formula(row+1,col+i,'vlookup(%s,Summary!B:AZ,column(%s),0)'%(name,self.str[i]+'1'))
                    row += 1
                    col = 0
                else:
                    pass

        row = col = 0
        '''写入生产发行表'''
        sheet2.write(row, col, '生产发行表')
        row += 1
        sheet2.write(row,col,'本工作表只能对“需求”项进行操作')
        row += 1
        i = 0
        for t in self.title :
            sheet2.write(row,col,t)
            col += 1
        for i in range(31) :
            sheet2.write(row,col+i,i+1)
        sheet2.write(row,col+31,'合计')
        row += 1
        col = 0
        for d in data :
            for li in d :
                if '总成' in li[0] :
                    print('正在写入%s' % li)
                    for l in li :
                        sheet2.write(row,col,l)
                        col += 1
                    sheet2.write(row, col, self.plan[1])
                    col += 1
                    i = 0
                    for i in range(0,32):
                        sheet2.write_formula(row,col,'sumifs(生产物料推移表!%s:%s,生产物料推移表!M:M,\"计划生产\",生产物料推移表!B:B,\"%s\")'%(self.str[i+4],self.str[i+4],li[1]))
                        col += 1
                    sheet2.write_formula(row,col+32,'sum(j%s:an%s)'%(row+1,row+1))
                    col = 0
                    row += 1

        row = col = 0
        '''写入生产推移表'''
        sheet3.write(row, col, '生产推移表')
        row += 1
        sheet3.write(row,col,'本工作表只能对“计划生产”项进行操作')
        row += 1
        for t in self.title :
            sheet3.write(row,col,t)
            col += 1
        i = 0
        for i in range(31) :
            sheet3.write(row,col+i,i+1)
        sheet3.write(row,col+31,'合计')
        row += 1
        col = 0
        for d in newdata :
            name = '\"'+str(d[1])+'\"'
            for l in d :
                 sheet3.write(row,col,l)
                 sheet3.write(row+1,col,l)
                 col += 1
            sheet3.write(row + 2, col, str(d[1]))
            sheet3.write(row + 3, col, str(d[1]))
            sheet3.write(row + 4, col, str(d[1]))
            sheet3.write(row + 5, col, str(d[1]))

            sheet3.write(row,col,self.plan[0])
            sheet3.write(row+1,col,self.plan[1])
            sheet3.write(row+2,col,self.plan[2])
            sheet3.write(row+3,col,self.plan[3])
            sheet3.write(row+4, col, self.plan[4])
            sheet3.write(row+5,col,self.plan[5])
            col += 1
            i = 0
            for i in range(32):
                if d[0] != '总成':
                    sheet3.write_formula(row,col,'sumifs(BOM!%s:%s,BOM!A:A,".1",BOM!B:B,\"%s\")'%(self.str[i+4],self.str[i+4],d[1]))
                    if i == 0:
                        sheet3.write_formula(row+5,col,'IF(DAY(NOW())<M3,INDIRECT(ADDRESS(%s,7))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)),INDIRECT(ADDRESS(%s,7))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)))'%(row+6,row+1,col+1,row+2,col+1,row+5,col+1,row+6,row+1,col+1,row+4,col+1,row+5,col+1))
                    else:
                        sheet3.write_formula(row+5,col,'IF(DAY(NOW())<M3,INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)),INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)))'%(row+6,col,row+1,col+1,row+2,col+1,row+5,col+1,row+6,col,row+1,col+1,row+4,col+1,row+5,col+1))
                    col += 1
                else:
                    if i == 0:
                        sheet3.write_formula(row+5,col,'IF(DAY(NOW())<M3,INDIRECT(ADDRESS(%s,7))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)),INDIRECT(ADDRESS(%s,7))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)))'%(row+6,row+1,col+1,row+2,col+1,row+5,col+1,row+6,row+1,col+1,row+4,col+1,row+5,col+1))
                    else:
                        sheet3.write_formula(row+5,col,'IF(DAY(NOW())<M3,INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)),INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s))+INDIRECT(ADDRESS(%s,%s))-INDIRECT(ADDRESS(%s,%s)))'%(row+6,col,row+1,col+1,row+2,col+1,row+5,col+1,row+6,col,row+1,col+1,row+4,col+1,row+5,col+1))
                    col += 1
                #sheet3.write_formula(row+1, col + i,'sumif(Plan!B:B,%s,Plan!%s:%s)'%(name,self.str[i+1],self.str[i+1]))
                #sheet3.write_formula(row+2, col + i,'indirect(address(%s,%s))+indirect(address(%s,%s))-indirect(address(%s,%s))' % ( row + 3, col+ i, row+1, col + i+1, row + 2, col + i+1))
            sheet3.write_formula(row,col+31,'sum(j%s:an%s)'%(row+1,row+1))
            row += 6
            col = 0

        f.close()
        print('写入成功，文件保存在：%s，名为<<%s>>的Excel(.xlsx)文件'%(self.arg,'生产推移表'))

    def wirte(self,data):
       wb  = openpyxl.Workbook()
       ws = wb.active
       ws1 = wb.create_sheet(index=0,title='1')
       i = 0
       for d in data:
           ws1.append([d])

       wb.save('1.xlsx')
