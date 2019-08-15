#coding=utf-8
import pandas as pd
import xlrd
import xlwt

######## Read DI data.
def read_di_data(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 0:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-MJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ")
                elif row[2] == 1:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-MJ")
                    X_list.append(row[1]+"-1DJ")
                    X_list.append(row[1]+"-2DJ")
                elif row[2] == 2:
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-MJ")
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
                    X_list.append(row[1] + "-LJ")
                    X_list.append(row[1] + "-UJ")
                    X_list.append(row[1] + "-YXJ")
                    X_list.append(row[1] + "-1DJ")
                    X_list.append(row[1] + "-2DJ")
                
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
            
    elif sheetName == "Drb":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"-DRJ")
            X_list.append(row[1]+"-DRJ")  
            
    elif sheetName == "Relay":
        rows = data_frame.iloc[:,[0,1,2,8]]
        for row in rows.values:
            if row[2] == 1:
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
    
######## Read DO data.
def read_do_data(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 0:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-MJ")
                elif row[2] == 1:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-MJ")
                elif row[2] == 2:
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                    X_list.append(row[1]+"-MJ")
                elif row[2] == 4:
                    X_list.append(row[1]+"-LJ")
                    X_list.append(row[1]+"-YXJ")    
                elif row[2] == 5:
                    X_list.append(row[1]+"-UJ")
                    X_list.append(row[1]+"-YXJ")
                elif row[2] == 6:
                    X_list.append(row[1] + "-LJ")
                    X_list.append(row[1] + "-UJ")
                    X_list.append(row[1] + "-YXJ")
                
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

######## Read CDI status.
def read_cdi_data(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [1,2,3]:
                X_list.append(row[1]+"_aspect")
                X_list.append(row[1]+"_block")
                X_list.append(row[1]+"_aspATP")
                X_list.append(row[1]+"_stopAss")
                X_list.append(row[1]+"_traStop")
                X_list.append(row[1]+"_supOlp")
                X_list.append(row[1]+"_ATapp")
                X_list.append(row[1]+"_NOUTapp")
                X_list.append(row[1]+"_colCon")
                X_list.append(row[1]+"_callOn")
                X_list.append(row[1]+"_wcuAppL")
                X_list.append(row[1]+"_admiss")
                
    

    elif sheetName == "Track":
        rows = data_frame.iloc[:,[0,1,4]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_block")
                X_list.append(row[1]+"_phyClr")
                X_list.append(row[1]+"_atpVaca")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_tsrSet")
                X_list.append(row[1]+"_tsrVal")
                X_list.append(row[1]+"_rotLck")
                X_list.append(row[1]+"_olpLck")                               
                X_list.append(row[1]+"_logClr")
                X_list.append(row[1]+"_rlsDel")
                X_list.append(row[1]+"_admiss")
    
    elif sheetName == "Point":
        rows = data_frame.iloc[:,[0,1,5]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_block")
                X_list.append(row[1]+"_phyClr")
                X_list.append(row[1]+"_atpVaca")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_tsrSet")
                X_list.append(row[1]+"_tsrVal")
                X_list.append(row[1]+"_rotLck")
                X_list.append(row[1]+"_olpLck")
                X_list.append(row[1]+"_position")
                X_list.append(row[1]+"_lock")
                X_list.append(row[1]+"_flkLck")
                X_list.append(row[1]+"_tvdFail")
                X_list.append(row[1]+"_trailed")
                X_list.append(row[1]+"_supVis")
                X_list.append(row[1]+"_logClr")
                
    elif sheetName == "TVS":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_tvsSta")
                X_list.append(row[1]+"_phyClr")
                X_list.append(row[1]+"_isReset")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_rotLck")
                X_list.append(row[1]+"_arb")
                X_list.append(row[1]+"_qjj")
                
    elif sheetName == "Psd":
        rows = data_frame.iloc[:,[0,1,4]] 
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_ovrCmd")
                X_list.append(row[1]+"_traStop")
                X_list.append(row[1]+"_type")
                X_list.append(row[1]+"_collect")
                X_list.append(row[1]+"_upstaSigID")
                X_list.append(row[1]+"_dnstaSinID")
                
    elif sheetName == "Emp":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_collect")

    elif sheetName == "Drb":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"_status")
            X_list.append(row[1]+"_cmd")
    
    elif sheetName == "PlatFormTrack":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"_hldSta")
            X_list.append(row[1]+"_skpSta") 
            
    elif sheetName == "Floodgate":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_collect")
                X_list.append(row[1]+"_clsReq")                                       
            
    elif sheetName == "Relay":
        rows = data_frame.iloc[:,[0,1,2,8]]
        for row in rows.values:
            X_list.append(row[1]+"_rlySta")  
    
    return(X_list)

######## Read logic output status.
def read_logic_output_data(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [1,2,3]:
                X_list.append(row[1] + "_id")
                X_list.append(row[1]+"_aspect")
                X_list.append(row[1]+"_aspATP")
                X_list.append(row[1]+"_collect")
                X_list.append(row[1]+"_autoMac")
                X_list.append(row[1]+"_clrCon")
                X_list.append(row[1]+"_byte1")
                X_list.append(row[1]+"_byte2")
                X_list.append(row[1]+"_byte3")
                               
    elif sheetName == "Point":
        rows = data_frame.iloc[:,[0,1,5]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1] + "_id")
                X_list.append(row[1]+"_position")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_reqPos")
                X_list.append(row[1]+"_assTrkId")
                X_list.append(row[1]+"_byte1")
                X_list.append(row[1]+"_byte2")
                
    elif sheetName == "Track":
        rows = data_frame.iloc[:,[0,1,4]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1] + "_id")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_lj1")
                X_list.append(row[1]+"_lj2")
                X_list.append(row[1]+"_qjj")
                X_list.append(row[1]+"_tsrSet")
                X_list.append(row[1]+"_tsrVal")
                X_list.append(row[1]+"_atpVcy")
                X_list.append(row[1]+"_rotId")                               
                X_list.append(row[1]+"_byte1")
                X_list.append(row[1]+"_byte2")                
    
    elif sheetName == "TVS":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1] + "_id")
                X_list.append(row[1]+"_lj1")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_lj2")
                X_list.append(row[1]+"_qjj")
                X_list.append(row[1]+"_rotId")
                X_list.append(row[1]+"_byte1")
                X_list.append(row[1]+"_byte2")
                
    elif sheetName == "Psd":
        rows = data_frame.iloc[:,[0,1,4]] 
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1] + "_id")
                X_list.append(row[1]+"_collect")
                X_list.append(row[1]+"_ovrCmd")
                X_list.append(row[1]+"_atpCmd")
                X_list.append(row[1]+"_type")                
                X_list.append(row[1]+"_byte1")
                
    elif sheetName == "Emp":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1] + "_id")
                X_list.append(row[1]+"_collect")
            
    elif sheetName == "Floodgate":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1] + "_id")
                X_list.append(row[1]+"_collect")
                X_list.append(row[1]+"_clsReq")
                X_list.append(row[1]+"_cmd")
                
    return(X_list)
    
######## Read adjci logic output cmd.
def read_adjci_logic_output_cmd(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName)

    if sheetName == "Track":
        rows = data_frame.iloc[:,[0,1,4]]
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_cmd0")
                X_list.append(row[1]+"_cmd1")
                X_list.append(row[1]+"_direct")                                 
                
    elif sheetName == "TVS":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_cmd0")
                X_list.append(row[1]+"_cmd1")
                X_list.append(row[1]+"_direct") 
                
    elif sheetName == "Psd":
        rows = data_frame.iloc[:,[0,1,4]] 
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_ovrCmd")
                X_list.append(row[1]+"_collect")                           
                
    elif sheetName == "Overlap":
        rows = data_frame.iloc[:,[0,1,9]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_ovpSta")
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_adjOvpCmd")
        for row in rows.values:
            if row[2] in [2,3]:
                X_list.append(row[1]+"_outOvFb")
                
    elif sheetName == "Route":
        rows = data_frame.iloc[:,[0,1,33]]
        for row in rows.values:
            if row[2] in [1, 2, 3]:
                X_list.append(row[1]+"_fatInfo")
                X_list.append(row[1]+"_cmdRet")
                X_list.append(row[1]+"_cmdFb")
                X_list.append(row[1]+"_eleId")
                X_list.append(row[1]+"_eleType")
        for row in rows.values:
            if row[2] in [1, 2, 3]:
                X_list.append(row[1]+"_status")
                
    return(X_list)

######## Recv adjci element status.
def recv_adjci_element_status(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName)

    if sheetName == "Track":
        rows = data_frame.iloc[:,[0,1,4]]
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_block")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_olpClm")
                X_list.append(row[1]+"_olpLck")
                X_list.append(row[1]+"_rotClm")
                X_list.append(row[1]+"_rotLck")
                X_list.append(row[1]+"_logClr")
                X_list.append(row[1]+"_phyClr")
                X_list.append(row[1]+"_atpVacany")
                X_list.append(row[1]+"_tvdFail")
                X_list.append(row[1]+"_tsrSet")
                X_list.append(row[1]+"_tsrVal")
                X_list.append(row[1]+"_enfRls")
                X_list.append(row[1]+"_qjj")                               
                X_list.append(row[1]+"_lj1")
                X_list.append(row[1]+"_lj2")
                X_list.append(row[1]+"_lockByCia")                               
    
    elif sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [4,5]:
                X_list.append(row[1]+"_block")
                X_list.append(row[1]+"_filament")
                X_list.append(row[1]+"_aspect")
                X_list.append(row[1]+"darkSig")
                X_list.append(row[1]+"_supOlp")
                X_list.append(row[1]+"autoMac")
                X_list.append(row[1]+"_ATapp")
                X_list.append(row[1]+"_traStop")
                X_list.append(row[1]+"_stopAss")
                X_list.append(row[1]+"_NOUTapp")
                X_list.append(row[1]+"_aspATP")
                
    elif sheetName == "TVS":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_qjj")
                X_list.append(row[1]+"_lj1")
                X_list.append(row[1]+"_lj2")
                X_list.append(row[1]+"_direct")
                X_list.append(row[1]+"_rotClm")
                X_list.append(row[1]+"_rotLck")
                X_list.append(row[1]+"_phyClr") 
                X_list.append(row[1]+"_failRls")
                X_list.append(row[1]+"_disable")
                X_list.append(row[1]+"_lockByCia") 
                X_list.append(row[1]+"_allTrkClr")
                
    elif sheetName == "Psd":
        rows = data_frame.iloc[:,[0,1,4]] 
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_atpCmd")
                X_list.append(row[1]+"_ovrCmd")
                X_list.append(row[1]+"_collect")
                X_list.append(row[1]+"_traStop")
                X_list.append(row[1]+"_type")
    
    elif sheetName == "Emp":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_status")
                    
    elif sheetName == "Overlap":
        rows = data_frame.iloc[:,[0,1,9]]
        for row in rows.values:
            if row[2] in [4,5]:
                X_list.append(row[1]+"_ovFbRet")              
                             
    return(X_list)

######## CAS inner data.
def cas_inner_property(xlsfile, sheetName,X_list):
    data_frame = pd.read_excel(xlsfile, sheet_name=sheetName)

    if sheetName == "Signal":
        rows = data_frame.iloc[:,[0,1,4,5]]
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 0:
                    X_list.append(row[1]+"_count")
                    X_list.append(row[1]+"_counTwoUp")
                    X_list.append(row[1]+"_collectCoun")
                    X_list.append(row[1]+"_countM")
                    X_list.append(row[1]+"_timeId")
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 1:
                    X_list.append(row[1]+"_count")
                    X_list.append(row[1]+"_counTwoUp")
                    X_list.append(row[1]+"_collectCoun")
                    X_list.append(row[1]+"_countM")
                    X_list.append(row[1]+"_timeId")
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 2:
                    X_list.append(row[1]+"_count")
                    X_list.append(row[1]+"_counTwoUp")
                    X_list.append(row[1]+"_collectCoun")
                    X_list.append(row[1]+"_countM")
                    X_list.append(row[1]+"_timeId")
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 4:
                    X_list.append(row[1]+"_count")
                    X_list.append(row[1]+"_collectCoun")
                    X_list.append(row[1]+"_timeId")   
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 5:
                    X_list.append(row[1]+"_count")
                    X_list.append(row[1]+"_collectCoun")
                    X_list.append(row[1]+"_timeId")
        for row in rows.values:
            if row[3] in [1,2,3]:
                if row[2] == 6:
                    X_list.append(row[1]+"_count")
                    X_list.append(row[1]+"_collectCoun")
                    X_list.append(row[1]+"_timeId")
                
    elif sheetName == "Point":
        rows = data_frame.iloc[:,[0,1,5]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_trTimeId")
                X_list.append(row[1]+"_sjTimeId")
                X_list.append(row[1]+"_reTimeId")
                X_list.append(row[1]+"_turn")
                X_list.append(row[1]+"_tmpTurn")
                X_list.append(row[1]+"_ilsValue")

    elif sheetName == "Psd":
        rows = data_frame.iloc[:,[0,1,4]] 
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_timeId")
                X_list.append(row[1]+"_turn")
                X_list.append(row[1]+"_ilsConSta")
                X_list.append(row[1]+"_ilsValue")
        
    elif sheetName == "TVS":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_jzResCoun")
                X_list.append(row[1]+"_jzRes")
            
    elif sheetName == "Floodgate":
        rows = data_frame.iloc[:,[0,1,3]]
        for row in rows.values:
            if row[2] in [1,2,3]:
                X_list.append(row[1]+"_ilsConSta")
                X_list.append(row[1]+"_ilsValue")
                X_list.append(row[1]+"_turn")
                X_list.append(row[1]+"_timeId")
            
    elif sheetName == "Drb":
        rows = data_frame.iloc[:,[0,1]]
        for row in rows.values:
            X_list.append(row[1]+"_lightFlhCoun")
            
    elif sheetName == "Relay":
        rows = data_frame.iloc[:,[0,1,2,8]]
        for row in rows.values:
            X_list.append(row[1]+"_timeId")
            X_list.append(row[1]+"_turn")
            X_list.append(row[1]+"_ilsValue")
            X_list.append(row[1]+"_outValue")
    
    return(X_list)

######## read App DIO
def read_App_DIO(xlsfile):
    Dilist = []
    Dolist = []
    sheetNames = ["Signal","Point","Track","Psd","TVS","Emp","Floodgate","Drb","Relay"]

    for sheetName in sheetNames:         
        di_list = read_di_data(xlsfile, sheetName, Dilist)
        do_list = read_do_data(xlsfile, sheetName, Dolist)
             
    dioList = di_list + do_list
    
    return(dioList)
    
######## read CDI data
def read_CDI_data(xlsfile):    
    CdiList = []
    sheetNames = ["Signal","Point","Track","TVS","Psd","Emp","Drb","PlatFormTrack","Floodgate","Relay"]

    for sheetName in sheetNames:         
        cdiList = read_cdi_data(xlsfile, sheetName, CdiList)
             
    cdiList.append("Console_Sta")
    cdiList.append("Occ_Console02")
    cdiList.append("Occ_Console01")
    
    return(cdiList)    

######## read logic data
def read_logic_data(xlsfile):    
    LogicList = []
    sheetNames = ["Signal","Point","Track","TVS","Psd","Emp","Floodgate"]

    for sheetName in sheetNames:         
        logicList = read_logic_output_data(xlsfile, sheetName, LogicList)
                 
    return(logicList)   

######## read adjci logic
def read_adjci_logic(xlsfile):
    AdjList = []
    sheetNames = ["Track","TVS","Psd","Overlap","Route"]

    for sheetName in sheetNames:         
        adjList = read_adjci_logic_output_cmd(xlsfile, sheetName, AdjList)
                 
    return(adjList)
    
######## recv adjci data
def recv_adjci_data(xlsfile):
    AdjList = []
    sheetNames = ["Track","Signal","TVS","Psd","Emp","Overlap"]

    for sheetName in sheetNames:         
        adjList = recv_adjci_element_status(xlsfile, sheetName, AdjList)
                 
    return(adjList)
    
######## cas inner data
def cas_inner_data(xlsfile):
    CasList = []
    sheetNames = ["Point","Signal","TVS","Psd","Floodgate","Drb","Relay"]

    for sheetName in sheetNames:         
        casList = cas_inner_property(xlsfile, sheetName, CasList)
                 
    return(casList)
    
######## frm timerPool data
def frm_timerPool_data():
    timerList = []
    timerList.append("TimerNum")
    for i in range(30):         
        timerList.append("Timer"+str(i+1)+"_Id")
        timerList.append("Timer"+str(i+1)+"_notUsed")
        timerList.append("Timer"+str(i+1)+"_count")
        
    return(timerList)
    
######## comm status data
def comm_status_data():
    commList = []
    commList.append("AppActualTime")
    commList.append("Rd_MsgNum")
    commList.append("Wt_MsgNum")
    commList.append("recvZCData")
    commList.append("recvZCMsgNum")
    commList.append("recvCILeft")
    commList.append("recvCILeftNum")
    commList.append("recvCIRight")
    commList.append("recvCIRightNum")
    commList.append("recvLEUNum")    

    for i in range(10):         
        commList.append("Ats"+str(i+1)+"_Id")
        commList.append("Ats"+str(i+1)+"_Alive")
        commList.append("Ats"+str(i+1)+"_LogOn")
    for i in range(15):
        commList.append("VobcHea"+str(i+1)+"_Id")
        commList.append("VobcHea"+str(i+1)+"_Alive")
    for i in range(15):
        commList.append("VobcPsd"+str(i+1)+"_Id")
        commList.append("VobcPsd"+str(i+1)+"_Alive")
    for i in range(15):
        commList.append("VobcReq"+str(i+1)+"_Id")
        commList.append("VobcReq"+str(i+1)+"_Alive")
    for i in range(20):         
        commList.append("CdiCmd"+str(i+1)+"_funNum")
        commList.append("CdiCmd"+str(i+1)+"_eleGlb")

    commList.append("frmAppCycId")
    for i in range(10):
        commList.append("Ats" + str(i + 1) + "_Id")
        commList.append("Ats" + str(i + 1) + "_MsgNum")
    for i in range(10):
        commList.append("Vobc" + str(i + 1) + "_Id")
        commList.append("Vobc" + str(i + 1) + "_MsgNum")

    return(commList)
    
######## According to relation of app and smw, input data to smwDioCfg.xls 
def CI_MDT_LIST_C(fileName):
    smwExcel = xlwt.Workbook()
    sheet1 = smwExcel.add_sheet(u'FieldInfo', cell_overwrite_ok = True)
    
    sheet1.write(0,0,"字段名")
    sheet1.write(0,1,"所属轴")
    sheet1.write(0,2,"起始字节索引")
    sheet1.write(0,3,"起始位索引")
    sheet1.write(0,4,"终止字节索引")
    sheet1.write(0,5,"终止位索引")
    sheet1.write(0,6,"字段编号")
    
    for i in range(100):
        for j in range(8):            
            sheet1.write(i*8+j+1,2,i)
            sheet1.write(i*8+j+1,4,i)
            sheet1.write(i*8+j+1,3,j)
            sheet1.write(i*8+j+1,5,j)
            sheet1.write(i*8+j+1,6,i*8+j+1)
                
    dioList = read_App_DIO(fileName)
    cdiList = read_CDI_data(fileName)
    logicList = read_logic_data(fileName)
    adjList = read_adjci_logic(fileName)
    recvAdjList = recv_adjci_data(fileName)
    casList = cas_inner_data(fileName)
    timerList = frm_timerPool_data()
    commList = comm_status_data()
    
    Total = cdiList + logicList + adjList + recvAdjList\
            + casList + timerList

    n = 1
    for data in dioList:
        sheet1.write(n,0,data)
        sheet1.write(n,1,'Y')
        n += 1
    DioNum = len(dioList)
    for i in range(DioNum,800):
        sheet1.write(i+1,0,'F'+str(i+1))
        sheet1.write(i+1,1,'Y')
        
    n = 801
    for data in Total:
        sheet1.write(n,0,data)
        sheet1.write(n,1,'Y')
        sheet1.write(n,2,n-701)
        sheet1.write(n,3,0)
        sheet1.write(n,4,n-701)
        sheet1.write(n,5,7)
        sheet1.write(n,6,n)
        n += 1
    k = n
    for data in commList:
        sheet1.write(n, 0, data)
        sheet1.write(n, 1, 'Y')
        sheet1.write(n, 3, 0)
        sheet1.write(n, 5, 7)
        sheet1.write(n, 6, n)
        # or "_funNum" or "_eleGlb"
        if "_Id"  in data:
            sheet1.write(n, 2, k - 701)
            sheet1.write(n, 4, k + 1 - 701)
            k += 2
        elif "_funNum"  in data:
            sheet1.write(n, 2, k - 701)
            sheet1.write(n, 4, k + 1 - 701)
            k += 2
        elif "_eleGlb"  in data:
            sheet1.write(n, 2, k - 701)
            sheet1.write(n, 4, k + 1 - 701)
            k += 2
        elif "frmAppCycId"  is data:
            sheet1.write(n, 2, k - 701)
            sheet1.write(n, 4, k + 3 - 701)
            k += 4
        else:
            sheet1.write(n, 2, k - 701)
            sheet1.write(n, 4, k - 701)
            k += 1
        n += 1

    smwExcel.save("FieldInfo-"+fileName[:-4]+".xls")
    print("Generate FieldInfo.xls!")
     
   



   