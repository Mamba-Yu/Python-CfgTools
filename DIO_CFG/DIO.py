#coding=utf-8
import pandas as pd
import xlrd
import openpyxl

### Open existed Excel file.
def open_exist_xlsx(diofile, bak):
    wb = openpyxl.load_workbook(diofile)

    sheet1 = wb[u'采集']
    sheet2 = wb[u'驱动']
    if bak == 0:
        files = ["AppDi.csv", "AppDo.csv"]
    if bak == 1:
        files = ["AppDi_bak.csv", "AppDo_bak.csv"]

    for file in files:
        f = open(file) 
        n = 1               
        li = list()
        data = f.readline()
        while data != "":
            li.append(data)
            data = f.readline()
            n += 1
        f.close()
        
        if file == "AppDi.csv" or file == "AppDi_bak.csv":
            write_appDio_data(li,sheet1,1)
        elif file == "AppDo.csv" or file == "AppDo_bak.csv":
            write_appDio_data(li,sheet2,2)

    wb.save(diofile)
    print("写入驱采信息表成功:%d!"%bak)
     
def write_appDio_data(list,ws,dio):        
    index = 0
    if dio == 1:
        k = 'di'
        m = 39
        n = 34
    elif dio == 2:
        k = 'do'
        m = 30
        n = 26
        
    for i in range(5):        
        for j in range(5):
            for col in ws.iter_cols(2+j*6, 2+j*6, 3+i*m, n+i*m):
                for cell in col:
                    cell.value = list[index].rstrip()                
                    index += 1
                    if index > len(list)-1:
                        print("写入 %s 点位数:%d"%(k,index))
                        return
                        

