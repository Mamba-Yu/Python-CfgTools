#coding=utf-8
import pandas as pd
import xlrd
import xlwt

### Read DI data.
def read_di_data(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName, header=0)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 0:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-AJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ")
                elif row[2] == 1:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-AJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ")
                elif row[2] == 2:
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-AJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ")
                elif row[2] == 3:
                    X_list.append(row[1]+"-DJ")
                elif row[2] == 4:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ")    
                elif row[2] == 5:
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ") 
                elif row[2] == 6:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ")
                
    elif sheetName == "Point":
        rows = data_frame.iloc[:,[0,1,5]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-SJ")
                X_list.append(row[1]+"-DBJ")
                X_list.append(row[1]+"-FBJ")
                X_list.append(row[1]+"-DFH")

    elif sheetName == "Psd":
        rows = data_frame.iloc[:,[0,1,4]] 
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-MSJ")
                X_list.append(row[1]+"-MSJ")
                X_list.append(row[1]+"-IOJ")
                X_list.append(row[1]+"-IOJ")
        
    elif sheetName == "TVS":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-GJ")
                X_list.append(row[1]+"-GJ")
            
    elif sheetName == "Emp":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-JTJ")
                X_list.append(row[1]+"-JTJ")
            
    elif sheetName == "Floodgate":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-KSJ")
                X_list.append(row[1]+"-KSJ")
                X_list.append(row[1]+"-GQJ")
                X_list.append(row[1]+"-GQJ")
            
    elif sheetName == "Awm":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-XCJXJ")
            X_list.append(row[1]+"-XCJXJ")
            X_list.append(row[1]+"-HQJ")
            X_list.append(row[1]+"-HQJ")
            X_list.append(row[1]+"-CFJ1")
            X_list.append(row[1]+"-CFJ1")
            X_list.append(row[1]+"-CFJ2")
            X_list.append(row[1]+"-CFJ2")
            X_list.append(row[1]+"-JTJ")
            X_list.append(row[1]+"-JTJ")
            X_list.append(row[1]+"-FHQJ")
            X_list.append(row[1]+"-FHQJ")
    
    elif sheetName == "Grd":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-MKSJ")
            X_list.append(row[1]+"-MKSJ")
            X_list.append(row[1]+"-PLJ")
            X_list.append(row[1]+"-PLJ")
            X_list.append(row[1]+"-DMJ")
    
    elif sheetName == "Spks":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-RFJ")
                X_list.append(row[1]+"-RFJ")
    
    elif sheetName == "Pcb":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-ZGJ")
            X_list.append(row[1]+"-ZGJ")
            
    elif sheetName == "Pccb":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-QQJ")
            X_list.append(row[1]+"-QQJ")
            
    elif sheetName == "Drb":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-DRJ")
            X_list.append(row[1]+"-DRJ")
            
    elif sheetName == "Relay":
        rows = data_frame.iloc[:,[0,1,2,8]]
        for row in rows.values:
            if row[3] == 1:
                X_list.append(row[1])               
            elif row[3] == 2:
                X_list.append(row[1])
                X_list.append(row[1])
            elif row[3] == 3:
                X_list.append(row[1]+"-LJ")
                X_list.append(row[1]+"-YXJ")
                X_list.append(row[1]+"-DJ")
        
    return(X_list)
    
### Read DO data.
def read_do_data(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName, header=0)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 0:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-AJ")
                elif row[2] == 1:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-AJ")
                elif row[2] == 2:
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-AJ")
                elif row[2] == 4:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-YXJ")    
                elif row[2] == 5:
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                elif row[2] == 6:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                
    elif sheetName == "Point":
        rows = data_frame.iloc[:,[0,1,5]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-SJ")
                X_list.append(row[1]+"-DCJ")
                X_list.append(row[1]+"-FCJ")

    elif sheetName == "Psd":
        rows = data_frame.iloc[:,[0,1,4]] 
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-KMJ")
                X_list.append(row[1]+"-GMJ")
        
    elif sheetName == "TVS":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-YFLJ")           
            
    elif sheetName == "Floodgate":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-XYGJ")
            
    elif sheetName == "Awm":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-XCQJ")
            X_list.append(row[1]+"-TWJ1")
            X_list.append(row[1]+"-TWJ2")
                
    elif sheetName == "Grd":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-KMJ")
            X_list.append(row[1]+"-GMJ")
                                     
    elif sheetName == "Drb":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-DRDJ")
            
    elif sheetName == "Relay":
        rows = data_frame.iloc[:,[0,1,2,8]]
        for row in rows.values:
            if row[2] == 2:
                X_list.append(row[1])                       
        
    return(X_list)

### read CI table
def read_App_DIO_CBTC(xlsfile):
    Dilist = []
    Dolist = []
    sheetNames = ["Signal","Point","Psd","TVS","Emp","Floodgate","Drb","Relay"]

    f1 = open('AppDi.csv', 'w')
    f2 = open('AppDo.csv', 'w')
    for sheetName in sheetNames:
        di_list = read_di_data(xlsfile, sheetName, Dilist)
        do_list = read_do_data(xlsfile, sheetName, Dolist)
    
    for di in di_list:
        f1.write(di+"\n")
    for do in do_list:
        f2.write(do+"\n")
    f1.close()
    f2.close()       
    
    print("完成 %s 应用驱采点位读取!"%(xlsfile[:-5]))

### 读取的驱采点位进行热备处理
def app_dio_backup():
    ##对AppDi表进行热备处理：
    f1 = open('AppDi.csv', 'r')
    li1 = list()
    data = f1.readline()
    while data != "":
        li1.append(data)
        data = f1.readline()
    f1.close()

    f2 = open('AppDi_bak.csv', 'w')
    tmplist = list()
    count = len(li1) % 32
    num = int(len(li1) / 32)

    for i in range(num + 1):
        if i < num:
            tmplist.extend(li1[i * 32 : i * 32 + 32])
            tmplist.extend(li1[i * 32 : i * 32 + 32])
        else:
            if count == 0:
                tmplist.extend(li1[i * 32 : i * 32 + 32])
                tmplist.extend(li1[i * 32: i * 32 + 32])
            else:
                tmplist.extend(li1[i * 32: i * 32 + count])
                tmplist.extend(li1[i * 32: i * 32 + count])

    for di in tmplist:
        f2.write(di)
    f2.close()
    ##对AppDo表进行热备处理：
    f3 = open('AppDo.csv', 'r')
    li1 = list()
    data = f3.readline()
    while data != "":
        li1.append(data)
        data = f3.readline()
    f3.close()

    f4 = open('AppDo_bak.csv', 'w')
    tmplist = list()
    count = len(li1) % 24
    num = int(len(li1) / 24)

    for i in range(num + 1):
        if i < num:
            tmplist.extend(li1[i * 24: i * 24 + 24])
            tmplist.extend(li1[i * 24: i * 24 + 24])
        else:
            if count == 0:
                tmplist.extend(li1[i * 24: i * 24 + 24])
                tmplist.extend(li1[i * 24: i * 24 + 24])
            else:
                tmplist.extend(li1[i * 24: i * 24 + count])
                tmplist.extend(li1[i * 24: i * 24 + count])
    for do in tmplist:
        f4.write(do)
    f4.close()
    