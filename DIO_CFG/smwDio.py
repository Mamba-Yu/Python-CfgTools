#coding=utf-8
import pandas as pd
import xlrd
import xlwt

### Read every column valid data.
def column_valid_data(xlsfile, f, sheetName, startNum, endNum, columnName):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName, header=1)
    
    column = data_frame.loc[startNum:endNum,columnName]
    for i in column:
        if pd.isnull(i):
            pass
        else:
            f.write(i+"\n")
         
### generate smwDiCfg.txt and smwDoCfg.txt
def smwDioCfg(xlsfile):
    f1 = open("smwDiCfg.csv", 'w')
    f2 = open("smwDoCfg.csv", 'w')
    sheetName = ["采集", "驱动"]
    for sheet in sheetName:
        if sheet == "采集":       
            for i in range(5): 
                startNum = i * 39
                endNum = 31 + i * 39               
                for j in range(5):
                    if j is 0:
                        column_valid_data(xlsfile, f1, sheet, startNum, endNum, "采集信息名称")
                    else:
                        column_valid_data(xlsfile, f1, sheet, startNum, endNum, "采集信息名称."+str(j))

        elif sheet == "驱动":
            for i in range(4):
                startNum = i * 30
                endNum = 23 + i *30            
                for j in range(5):
                    if j is 0:
                        column_valid_data(xlsfile, f2, sheet, startNum, endNum, "驱动信息名称")
                    else:
                        column_valid_data(xlsfile, f2, sheet, startNum, endNum, "驱动信息名称."+str(j)) 

    f1.close()
    f2.close()
