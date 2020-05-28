import os,sys,time
from excel.excel import excel
import pandas as pd

print('*'*100)
print('New Time：',time.localtime(time.time()))
print('welcome to use!')
print('Please enter the working directory:')
print('Enter \'?q\' exit!')
print('*'*100)
directory = input('Directly enter the current directory!(%s)'%os.getcwd())

if directory == '?q' :
    sys.exit(0)
elif directory == '' :
    directory = os.getcwd()+'/'
else:
    directory = directory+'/'

list = os.listdir(directory)
datas = []

e = excel()

for l in list :
    if 'xls' in l :
        print('正在读取文件：%s'%l)
        datas = e.readExcel(directory+l)
    else:
        print(l,'当前文件个不是.xls或.xlsx文件，系统将会自动跳过该文件！')

df
print(datas[0][0])



print('*'*100)