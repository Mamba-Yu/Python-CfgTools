#coding=utf-8
import pandas as pd
import xlrd
import xlwt
import os
from smwDio import smwDioCfg
from ioConf import Every_row
                              
### According to relation of app and smw, input data to smwDioCfg.xls 
def App_smw_relation(fileNo1, fileNo2):
    brk = 1
    smwExcel = xlwt.Workbook()
    sheet1 = smwExcel.add_sheet(u'DI', cell_overwrite_ok = True)
    sheet2 = smwExcel.add_sheet(u'DO', cell_overwrite_ok = True)
    sheet3 = smwExcel.add_sheet(u'DI_DEF', cell_overwrite_ok = True)
    sheet4 = smwExcel.add_sheet(u'DO_DEF', cell_overwrite_ok = True)
        
    sheet1.write(0,0,"App")
    sheet1.write(0,1,"Smw")
    sheet1.write(0,2,"Name")        
    sheet3.write(0,0,"App")
    sheet3.write(0,1,"Default")
    sheet2.write(0,0,"App")
    sheet2.write(0,1,"Smw")
    sheet2.write(0,2,"Name")        
    sheet4.write(0,0,"App")
    sheet4.write(0,1,"Default")  
    
    smwDioCfg(fileNo2)

    Every_row(fileNo2, 4, 5)

    files = ["smwDiCfg.csv", "smwDoCfg.csv", "AppDi.csv", "AppDo.csv"]
    Tlist = list()
    for file in files:
        f = open(file) 
        n = 1               
        li = list()
        data = f.readline()
        while data != "":
            li.append(data)
            if file == "AppDi.csv":
                sheet1.write(n, 0, n)            
                sheet3.write(n, 1, 0)
                sheet3.write(n, 0, n)
            elif file == "AppDo.csv":
                sheet2.write(n, 0, n)
                sheet4.write(n, 1, 0)
                sheet4.write(n, 0, n)
            else:
                pass
            data = f.readline()
            n += 1
        f.close()    
        Tlist.append(li)
    print("OK2")
    seen = set()
    duplicated = set()
    for Di in Tlist[0]:
        if Di not in seen:
            seen.add(Di)
        else:
            duplicated.add(Di)
    
    i = 0 
    appLen = len(Tlist[2])
    while i < appLen:
        if Tlist[2][i] in Tlist[0]:
            if Tlist[2][i] in duplicated:
                sheet1.write(i+1,1,Tlist[0].index(Tlist[2][i])+1)
                sheet1.write(i+2,1,Tlist[0].index(Tlist[2][i])+2)
                sheet1.write(i+1,2,Tlist[2][i])
                sheet1.write(i+2,2,Tlist[2][i])
                i += 2
            else:                          
                sheet1.write(i+1,1,Tlist[0].index(Tlist[2][i])+1)
                sheet1.write(i+1,2,Tlist[2][i])
                i += 1
        else:
            print("在APPDI中未找到：%s"%(Tlist[0][i]))
            brk = 0
            break
    print("OK3")
    j = 1        
    for Do in Tlist[3]:
        if Do in Tlist[1]:
            sheet2.write(j,1,Tlist[1].index(Do)+1)
            sheet2.write(j,2,Do)
            j += 1
        else:
            print("在APPDO中未找到：%s"%Do)
            brk = 0
            break
                    
    smwExcel.save("smwDioCfg_"+fileNo1[:-5]+".xls") 
    
    read_smwDio_data("smwDioCfg_"+fileNo1[:-5]+".xls") 

    fileList = ["smwDiCfg.csv", "smwDoCfg.csv", "AppDi.csv", "AppDo.csv"]
    for file in fileList:
        if os.path.exists(file):
            os.remove(file)

    return brk

#热备版APP与SMW点位对应关系
def App_smw_relation_backup(fileNo1, fileNo2, alarmDiNum=0):
    app_ret = 1
    smwExcel = xlwt.Workbook()
    sheet1 = smwExcel.add_sheet(u'DI', cell_overwrite_ok = True)
    sheet2 = smwExcel.add_sheet(u'DO', cell_overwrite_ok = True)
    sheet3 = smwExcel.add_sheet(u'DI_DEF', cell_overwrite_ok = True)
    sheet4 = smwExcel.add_sheet(u'DO_DEF', cell_overwrite_ok = True)
        
    appNum = []
    sheet1.write(0,0,"App")
    sheet1.write(0,1,"Smw")      
    sheet3.write(0,0,"App")
    sheet3.write(0,1,"Default")
    sheet2.write(0,0,"App")
    sheet2.write(0,1,"Smw")       
    sheet4.write(0,0,"App")
    sheet4.write(0,1,"Default")
          
    smwDioCfg(fileNo2)
    Every_row(fileNo2, 4, 5)           

    files = ["smwDiCfg.csv", "smwDoCfg.csv", "AppDi.csv", "AppDo.csv"]
    Tlist = list()
    for file in files:
        f = open(file) 
        li = list()
        data = f.readline()
        while data != "":
            li.append(data)            
            data = f.readline()            
        f.close()    
        Tlist.append(li)

    print("smwDI:%d, appDI:%d, smwDO:%d, appDO:%d"%(len(Tlist[0]), len(Tlist[2]), len(Tlist[1]), len(Tlist[3])))
    chk_di =  (len(Tlist[2]) - int(alarmDiNum)) * 2 - (len(Tlist[0]) - int(alarmDiNum))
    chk_do = (len(Tlist[3]) * 2 )- len(Tlist[1])
    if chk_di != 0 or chk_do != 0:
        print("驱采表中热备点数有误！")
        app_ret = 0
        return app_ret
    m = 1
    for n in range(1,len(Tlist[2])+1):
        sheet3.write(n, 1, 0)
        sheet3.write(n, 0, n)
        if n < len(Tlist[2]) - int(alarmDiNum) + 1:
            sheet1.write(m, 0, n)
            sheet1.write(m+1, 0, n)
            m += 2
        else:
            sheet1.write(m, 0, n)
            m += 1
        n += 1

    m = 1
    for n in range(1,len(Tlist[3])+1):
        sheet4.write(n, 1, 0)
        sheet4.write(n, 0, n)
        sheet4.write(len(Tlist[3])+n, 1, 0)
        sheet4.write(len(Tlist[3])+n, 0, len(Tlist[3])+n)
        if n < len(Tlist[3]) + 1:
            sheet2.write(m, 0, n)
            sheet2.write(m+1, 0, n)
            m += 2
        else:
            sheet2.write(m, 0, n)
            m += 1
        n += 1   

    i = 0   
    m = 1
    tmpdilist = set()
    while i < len(Tlist[2]): 
        j = 0
        if Tlist[0][j] in Tlist[2]: 
            index_list = [k for k,x in enumerate(Tlist[0]) if x == Tlist[2][i]]
            while j < len(Tlist[0]): 
                if Tlist[0][j] == Tlist[2][i]:
                    if j in tmpdilist:                  
                        pass                    
                    else:                            
                        tmpdilist.add(j)
                        sheet1.write(m,1,j+1)
                        sheet1.write(m,2,Tlist[0][j])
                        m += 1
                        if len(index_list) == 4:
                            j += 1
                j += 1                
        else:
            print("在APPDI中未找到：%s"%Tlist[0][j])
            app_ret = 0
            break
        i += 1

    i = 0     
    m = 1
    tmpdolist = set()
    while i < len(Tlist[3]): 
        j = 0
        if Tlist[1][j] in Tlist[3]: 
            while j < len(Tlist[1]):                
                if Tlist[1][j] == Tlist[3][i]:
                    if j in tmpdolist:
                        pass
                    else:
                        tmpdolist.add(j)
                        sheet2.write(m,1,j+1)
                        sheet2.write(m,2,Tlist[1][j])
                        m += 1
                j += 1                
        else:
            print("在APPDO中未找到：%s"%Tlist[1][j])
            app_ret = 0
            break
        i += 1       
                    
    smwExcel.save("smwDioCfg_"+fileNo1[:-5]+".xls") 
    
    read_smwDio_data("smwDioCfg_"+fileNo1[:-5]+".xls") 

    fileList = ["smwDiCfg.csv", "smwDoCfg.csv", "AppDi.csv", "AppDo.csv", "AppDi_bak.csv", "AppDo_bak.csv"]
    for file in fileList:
        if os.path.exists(file):
            os.remove(file)

    return app_ret
            
### Read smwDioCfg.xls data.
def read_smwDio_data(xlsfile):
    f = open(xlsfile[:-4]+".txt", "w")
    
    sheetList = ["DI","DO","DI_DEF","DO_DEF"]
    for sheet in sheetList:
        df = pd.read_excel(xlsfile, sheet_name=sheet)
        f.write(sheet+":")
        for i in range(len(df)):
            if i < len(df)-1:
                str_ele = str(df.iloc[i,0])+","+str(df.iloc[i,1])+";"
            elif i == len(df)-1:
                str_ele = str(df.iloc[i,0])+","+str(df.iloc[i,1])+"."
            f.write(str_ele)
    f.close()        
   
   



   