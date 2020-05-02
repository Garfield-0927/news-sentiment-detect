# process the row data

# -*- coding: utf-8 -*-


#导入需要使用的包
import xlrd  #读取Excel文件的包
import xlsxwriter   #将文件写入Excel的包
import pandas as pd

#打开一个excel文件
def open_xls(file):
    f = xlrd.open_workbook(file)
    return f

#获取excel中所有的sheet表
def getsheet(f):
    return f.sheets()

#获取sheet表的行数
def get_Allrows(f,sheet):
    table=f.sheets()[sheet]
    return table.nrows

#读取文件内容并返回行内容
def getFile(file,shnum):
    f=open_xls(file)
    table=f.sheets()[shnum]
    num=table.nrows
    for row in range(num):
        rdata=table.row_values(row)
        datavalue.append(rdata)
    return datavalue

#获取sheet表的个数
def getshnum(f):
    x=0
    sh=getsheet(f)
    for sheet in sh:
        x+=1
    return x

#函数入口
if __name__ == '__main__':
    #定义要合并的excel文件列表
    allxls = ['D:\\garfield\\news-sentiment-detect\\data\\1.xlsx', 'D:\\garfield\\news-sentiment-detect\\data\\2.xlsx', 'D:\\garfield\\news-sentiment-detect\\data\\3.xlsx', 'D:\\garfield\\news-sentiment-detect\\data\\4.xlsx'] #列表中的为要读取文件的路径
    #存储所有读取的结果
    datavalue = []
    for fl in allxls:
        f=open_xls(fl)
        x=getshnum(f)
        for shnum in range(x):
            print("正在读取文件："+str(fl)+"的第"+str(shnum)+"个sheet表的内容...")
            rvalue = getFile(fl,shnum)
    #定义最终合并后生成的新文件
    endfile='D:\\garfield\\news-sentiment-detect\\data\\DataSet_merge.xlsx'
    wb=xlsxwriter.Workbook(endfile)
    #创建一个sheet工作对象
    ws=wb.add_worksheet()
    for a in range(len(rvalue)):
        for b in range(len(rvalue[a])):
            c=rvalue[a][b]
            ws.write(a,b,c)
    wb.close()

    print("文件合并完成")

#以上代码是将数据合并至一个excel

#一下代码是进行数据清理


data = pd.DataFrame(pd.read_excel('D:\\garfield\\news-sentiment-detect\\data\\DataSet_merge.xlsx', 'Sheet1'))

# 查看读取数据内容
print(data)

# 查看是否有重复行
re_row = data.duplicated()
print(re_row)

# 查看去除重复行的数据
no_re_row = data.drop_duplicates(subset=['id'], keep='first', inplace=False)


# 查看基于[物品]列去除重复行的数据
wp = data.drop_duplicates(['id'])
print(wp)

# 将去除重复行的数据输出到excel表中
no_re_row.to_excel("D:\\garfield\\news-sentiment-detect\\data\\train_data.xlsx")

#一下代码是去除有空白行，以及没有内容的新闻
data = pd.read_excel('D:\\garfield\\news-sentiment-detect\\data\\train_data.xlsx',sheet_name='Sheet1')
datanota = data[data['news_content_translate_en'].notna()]
datanota.to_excel("D:\\garfield\\news-sentiment-detect\\data\\train_data_final.xlsx")



