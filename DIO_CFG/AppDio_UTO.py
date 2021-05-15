#coding=utf-8
import pandas as pd
import xlrd
import xlwt
from DIO import open_exist_xlsx

### Read DI data.
def read_di_data(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName, header=0)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,3,4]]
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
                    X_list.append(row[1]+"-DXJ")
                    X_list.append(row[1]+"-DJ")
                
    elif sheetName == "Point":
        rows = data_frame.iloc[:,[0,1,4]]
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
        rows = data_frame.iloc[:,[0,1,3,8]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                if row[3] == 'G':
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
        rows = data_frame.iloc[:,[0,1,3,4]]
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
                    X_list.append(row[1]+"-DXJ")
                
    elif sheetName == "Point":
        rows = data_frame.iloc[:,[0,1,4]]
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
        rows = data_frame.iloc[:,[0,1,3,8]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                if row[3] == 'G':
                    X_list.append(row[1]+"-YFLJ")            
            
    elif sheetName == "Floodgate":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"-KMJ")
            
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
def read_App_DIO_UTO(xlsfile):
    Dilist = []
    Dolist = []
    sheetNames = ["Signal","Point","Psd","TVS","Emp","Floodgate","Drb","Awm","Grd","Spks","Pcb","Pccb","Relay"]

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
    
    print("Write App data success!")
    
    
    