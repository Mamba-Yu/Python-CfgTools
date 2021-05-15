#coding=utf-8
import pandas as pd
import xlrd
import xlwt

### Read every row ioConf data.
def row_ioConf(xlsfile, f, sheetName, headLine, endNo):
    df = pd.read_excel(xlsfile, sheet_name=sheetName, header=headLine)
    
    firstC = []
    secondC = []
    for i in df.columns:         
        if 'DO' in i:
            if i[1:-6] == 'I':
                k = 2
            elif i[1:-6] == 'II':
                k = 4
            elif i[1:-6] == 'III':
                k = 6
            firstC.append(k)
            secondC.append(i[-5:-3])
        elif 'DI' in i:
            if i[1:-6] == 'I':
                k = 1
            elif i[1:-6] == 'II':
                k = 3
            elif i[1:-6] == 'III':
                k = 5
            firstC.append(k)
            secondC.append(i[-5:-3])       

    columnCount = [0 for k in range(5)]
    for i in range(5):
        num = 0
        columnNum = 1 + i * 6
        column = df.iloc[1:endNo, columnNum]
        for j in column:
            if pd.isnull(j):
                pass
            else:
                num += 1
        columnCount[i] = num  
   
    for i in range(5):
        if columnCount[i] is 0:
            break
        out1 = firstC[i]
        out2 = int(secondC[i])
        out3 = columnCount[i]
        
        outList = [out1,out2,out3,0]
        
        out_str = map(str,outList)
        out_str = ",".join(out_str)+"\n"
        f.write(out_str)
    
### Generate ioConf.txt
def Every_row(xlsfile, doRowNum, diRowNum):
    f = open("ioConf_"+xlsfile[:-5]+".txt","w")
    for i in range(doRowNum):
        rowNum = i * 30  
        row_ioConf(xlsfile,f, "驱动", rowNum, 25)

    for i in range(diRowNum):
        rowNum = i * 39  
        row_ioConf(xlsfile,f, "采集", rowNum, 33)
    
    f.close()  
    print("生成ioConf-%s!"%(xlsfile[:-5]))