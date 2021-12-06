

import tkinter as tk
from tkinter import filedialog
import datetime
import numpy as np 
import pandas as pd
import chardet
import os

def ad_timestampdate(timestamp):
    if timestamp != 0:
        return (datetime.datetime(1601, 1, 1) + datetime.timedelta(seconds=timestamp/10000000)).date()
    return np.nan
def ad_timestamptime(timestamp):
    if timestamp != 0:
        return (datetime.datetime(1601, 1, 1) + datetime.timedelta(seconds=timestamp/10000000)).time()
    return np.nan

root = tk.Tk()
root.withdraw()

dir_path = filedialog.askdirectory()
allscans = pd.DataFrame(columns= [i for i in range (1,15)])
for (root,dirs,files) in os.walk(dir_path, topdown=True):
    
    if files and "Active Directory Results" in str(root):
        if 'AD-userResult.csv' in files and 'AD-computerResult.csv' in files and 'AD-trustsResult.csv' in files and 'AD-usersAndGroupsResult.csv' in files:
            userresults = root + '/AD-userResult.csv'
            computerresults = root + '/AD-computerResult.csv'
            trustsresults = root + '/AD-trustsResult.csv'
            usersandgroupsresults = root + '/AD-usersAndGroupsResult.csv'
            #-----------------------------------------
            df = pd.read_csv(userresults,encoding='UTF-16')
            #with open(userresults) as f:
             #   print(f)

            # print(dir_path)
            # print(userresults)
            df['lastLogonTime']=df['lastLogonTimestamp'].fillna(0).apply(ad_timestamptime)
            df['lastLogonDate']=df['lastLogonTimestamp'].fillna(0).apply(ad_timestampdate)

            # print(df.head(5)[['DN','lastLogonTimestamp', 'lastLogonDate', 'lastLogonTime']])

            #df.to_csv(dir_path+'/AD-userResult-Excel.csv')

            #----------------------------------------------
            df2 = pd.read_csv(computerresults, encoding='UTF-16')

            df2['lastLogonTime']=df2['lastLogonTimestamp'].fillna(0).apply(ad_timestamptime)
            df2['lastLogonDate']=df2['lastLogonTimestamp'].fillna(0).apply(ad_timestampdate)

            #df.to_csv(dir_path+'/AD-computerResult-Excel.csv')

            #----------------------------------------------
            df3 = pd.read_csv(trustsresults, encoding='UTF-16', sep=';')
            #df.to_csv(dir_path+'/AD-trustsResult-Excel.csv')

            #----------------------------------------------
            df4 = pd.read_csv(usersandgroupsresults, encoding='UTF16', sep=';')
            #df4.to_excel(dir_path+'/AD-usersAndGroupsResult-Excel.xlsx')

            with pd.ExcelWriter(root+'/AD-results-Excel.xlsx') as writer: 
                df.to_excel(writer, sheet_name='UserResult', index=False)
                df2.to_excel(writer, sheet_name='ComputerResult', index=False)
                df3.to_excel(writer, sheet_name='TrustsResult', index=False)
                df4.to_excel(writer, sheet_name='UsersAndGroupsResult', index=False) 
      
    elif files and "DIT Results" in root:
        for file in files:
            if "scanResult" in file:
                results_path = str(root) + "/"+str(file)
                df = pd.read_csv(results_path,encoding='UTF-16',sep=";", header=None, names= [i for i in range(1,15)])
                allscans = allscans.append(df, ignore_index=True)


osdf = allscans[allscans[1] == "[OS-CONTENT]"]
osdf = osdf.rename(columns={3:"OS Name",4:"OS Version"})
osdf = osdf.groupby(["OS Name","OS Version"]).size().to_frame("Anzahl").reset_index()

arpdf = allscans[allscans[1] == "[ARP-CONTENT]"]
arpdf = arpdf.rename(columns={3:"Software Name",5:"Version"})
arpdf = arpdf.groupby(["Software Name","Version"]).size().to_frame("Anzahl").reset_index()

rdsdf = allscans[allscans[1] == "[RDS-CONTENT]"]
rdsdf = rdsdf.rename(columns={6:"User Name",7:"SAM Account Name"})
rdsdf = rdsdf.groupby(["User Name","SAM Account Name"]).size().to_frame("Anzahl").reset_index()

hwdf = allscans[allscans[1] == "[HARDWARE-CONTENT]"]
hwdf = hwdf.rename(columns={3:"Processor Name",4:"Number of Processors", 5:"Number of Cores", 6:"Number of Logical Processors"})
hwdf = hwdf.groupby(["Processor Name"])["Number of Processors", "Number of Cores", "Number of Logical Processors"].sum().reset_index()

sqldf = allscans[allscans[1] == "[SQL-CONTENT]"]
sqldf = sqldf.rename(columns={3:"SQL Instance",4:"SQL Host Name", 5:"Edition", 6:"Version"})
sqldf = sqldf.groupby(["SQL Instance","SQL Host Name", "Edition", "Version"]).size().to_frame("Anzahl").reset_index()

visiodf = allscans[allscans[1] == "[VISIO-CONTENT]"]
visiodf = visiodf.rename(columns={3:"License Cache"})
visiodf = visiodf.groupby(["License Cache"]).size().to_frame("Anzahl").reset_index()

authdf = allscans[allscans[1] == "[AUTH-CONTENT]"]
authdf = authdf.rename(columns={4:"Application"})
authdf = authdf.groupby(["Application"]).size().to_frame("Anzahl").reset_index()

def multiple_dfs(df_list, sheets, file_name, spaces):
    writer = pd.ExcelWriter(dir_path+"/"+file_name)   
    col = 0
    for dataframe in df_list:
        dataframe.to_excel(writer,sheet_name=sheets,startrow=0 , startcol=col, index=False)   
        col = col + len(dataframe.columns) + spaces + 1
    writer.save()

multiple_dfs([osdf, arpdf, rdsdf, hwdf, sqldf, visiodf, authdf], "ELP", "ScanResultSummary.xlsx", 1)