#!/usr/bin/env python
# coding: utf-8
# For no more download file from MRP to update if itemno.xlsx updated.
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import matplotlib.dates as mdates
import numpy as np
import datetime
import sys
import pathlib
import os
import re
from time import sleep

# download bt_billing and save to daily_billing
# print(os.listdir('.'))
files = []
for f in os.listdir('.'):
    if re.match(r'BT_Billing_\d{12}.xls$',f):
        files.append(f)
print(files)
sleep(3)
## No need remove old file
# for f in files:
#     os.remove(f)
#     print(f,'is removed')

# # No need seleium webdriver that for download file. Line 30~63.
# from selenium import webdriver
# # from chromedriver_py import binary_path
# from selenium.webdriver.chrome.options import Options
# from webdriver_manager.chrome import ChromeDriverManager
# options = Options()
# options.add_experimental_option("prefs",{"download.default_directory": r"D:\coding\daily_billing"})
# driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), chrome_options=options) 


# # launch chrome to open the following URL 
# url = 'http://mpserver.supermicro.com/mrpbt/BtBillingQueryByRef.aspx'
# driver.get(url)
# # Click Display button to show latest data 
# driver.find_element_by_name('btnDisplay').click()

# # Timeout to wait for query 
# sleep(3)

# # 43~45 Click “Display in Excel” to save the query result as BoQueryTW.aspx
# driver.find_element_by_name('btnExport').click() 
# sleep(10)
# driver.close()
# sleep(3)
# # finish download
# #! ls *.xls
# print(os.listdir('.'))
# files = []
# for f in os.listdir('.'):
#     if re.match(r'BT_Billing_\d{12}.xls$',f):
#         files.append(f)
# print(files)
# input('Press any Key to continue......')

# # # finish download


# Shorten key long file name. Just key mmddhhmm.
a_Fname = 'BT_Billing_'
b_Fname = str( datetime.datetime.today().year)
c_Fname = input('Key file date "mmddhhmm"')
d_Fname = '.xls'
filename = a_Fname + b_Fname + c_Fname + d_Fname
print (filename)


currentPath= pathlib.Path().absolute()

bt_bill=pd.read_html(filename)
billdata=pd.DataFrame(bt_bill[0])
# print(billdata)
# Read html file (downlad named as xxx.xls but actually is HTML). Not necessary -Skip first row as excel file but use "[0]". 
# Read excel file either xls or xlsx. Skip first row to get correct title. 

salesName = pd.DataFrame(pd.read_excel('SalesList.xlsx')).drop(['Team','Sales Forecast Y/N','Group leader','Sales','Head count  by team','Head count by group','in SJ office','Location','Current month Hire'],axis=1)
itemno_cat= pd.DataFrame(pd.read_excel('itemno_cat.xlsx'))
solddata=billdata.drop(['Order/DN Num','Invoice','Customer','ZDSR','Type'], axis=1)
# solddata=billdata.drop(['Order/DN Num','Invoice','Customer','ZDSR','Type'], axis=1)
# Write dataframe to new .xls file in new dir(currenet+newfilename)

newfile = 'GC%s' %filename
# print("New File Name: ",newfile)
newDirName= filename[:-4]
# print("New Dir Name:",newDirName)

currPath=pathlib.Path().absolute()

print("Current Path Name: ",currPath)

newPath=f"{currPath}/{newDirName}/{newfile}"
# newPath=f"{newDirName}/{newfile}"

print("New Path and file name is: ",newPath)

# Create New Dir for store new file in new Path.
newDir=f"{currPath}/{newDirName}"

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

createFolder(newDir)

writer = pd.ExcelWriter(newPath)
billdata.to_excel(writer, sheet_name= 'Raw_Data',na_rep=False,index=False,header=True)

solddata= solddata.merge(salesName,how='left',left_on='Sales',right_on='Rep code')
# Merge 2 tables for Sales ID (code) to find sales name that easy to know. (May also use .join method)
solddata = solddata.merge(itemno_cat, how='left', left_on='Itemno', right_on='Itemno')

solddata['Date']= pd.to_datetime(solddata['Date'].astype(str), format='%Y%m%d')
solddata['Date']= solddata['Date'].apply(lambda x:x.strftime('%Y-%m-%d'))
# Covert Date type from object to datetime and Change Date format to be 2020-01-01
# print(solddata)
# uptoShip = solddata.cumsum(solddata['Qty'], axis =1)

# Write new dataframe to sheet name= New_DF;

solddata.to_excel(writer, sheet_name= 'New_DF',na_rep=False,index=True,header=True)

# Summary each day shipment amount;
# solddata.groupby(['Date']).sum().plot(kind='bar',x='Date', figsize=(12,6), fontsize= 7, rot=45, title='Daily Shipment')
# solddata.groupby(['Date']).agg(sum).reset_index().plot(kind='bar',x='Date', figsize=(8,6), fontsize= 7, rot=45, title='Daily Shipment')

dailyShip=solddata.groupby(['Date']).agg(sum).reset_index()
temp=dailyShip[['Qty','Node Qty']]

# dailySum=dailyShip.append(dailyShip.sum(numeric_only=True), ignore_index=True)
dailySum= dailyShip[['Qty','Node Qty']].sum()
dailySum['Date']= 'UptoDate Total'
dailyShip = dailyShip.append(dailySum,ignore_index=True)
dailyShip.to_excel(writer, sheet_name= 'by_Date', index=False,header=True)


# List Itemno Top sold.
top_item=solddata.groupby('Itemno')[['Qty','Node Qty']].agg(sum).sort_values(by='Node Qty',ascending=False)
print(top_item)
top_item.to_excel(writer, sheet_name= 'top_Item', index=True,header=True)


# plt.subplots_adjust(bottom = 0.2)

dailyShip.plot(kind='bar',x='Date', figsize=(8,6), fontsize= 7, rot=75, title='Daily Shipment')
plt.tight_layout()
plt.savefig(f'{newDir}/dailyship%s.png' %filename)

# plt.show()

# print(solddata.groupby('Itemno')['Qty','Node Qty'].agg(sum).nlargest(30,'Qty'))

# List Model name by_Itemno and sum up sold Qty (System) and Node count;
solddata.groupby('Itemno')[['Qty','Node Qty']].agg(sum).nlargest(20,'Qty').plot(kind='barh',title='Top 20 Models', figsize=(8,6),fontsize=7)
plt.tight_layout()
plt.savefig(f'{newDir}/Top20Model_%s.png' %filename)
#plt.show()


# plot (by_ShipTo)
solddata.groupby('Ship To')['Qty','Node Qty'].agg(sum).nlargest(10,'Qty').plot.pie(autopct='%.1f%%', fontsize=8, figsize=(12,8), legend=False, subplots=True)
# by_ShipTo=solddata.groupby('Ship To')[['Qty']].agg(sum).sort_values(by='Qty',ascending=True)
# print(by_ShipTo)
plt.savefig(f'{newDir}/ShipTo_%s_bySys.png' % filename)
# plt.show()

# solddata.groupby('Ship To')['Qty'].nlargest(10).plot.pie(autopct='%.2f', fontsize=10, figsize=(8,6), legend=False, subplots=False)
# plt.title('By Country Sold Amount -System',fontsize=18, fontweight='bold')

# plt.savefig(f'{newDir}/ShipTo_%s_bySys.png' % filename)
# plt.show()

# by_ShipTo['Node Qty'].nlargest(10).plot.pie(autopct='%.2f', fontsize=10, figsize=(8,6), legend=False, subplots=False)
# plt.title('By Country Sold Amount -Nodes',fontsize=18, fontweight='bold')
# plt.savefig(f'{newDir}/ShipTo_%s_byNode.png' % filename)

# plt.show()

# List Sales name and sum up sold Qty (System) and Node count;
solddata.groupby('Name')[['Qty','Node Qty']].agg(sum).nlargest(20,'Qty').plot(kind='barh',figsize=(8,5))
plt.title('Top 20 Sales', fontsize=18, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{newDir}/Top Sale_%s.png' %filename)
#plt.show()

# print(solddata)
#By Country and Sales 
# by_sales=solddata.groupby(['Ship To','Sales']).Qty.sum()
by_sales_ship_item=solddata.groupby(['Name','Ship To','Itemno']).agg(sum)
# print(by_sales_ship_item)

# List Sales Name and what Itemno they sold.
# print(solddata.groupby(['Name','Itemno']).agg(sum))
sales_item=solddata.groupby(['Name','Itemno']).agg(sum)
sales_item.to_excel(writer, sheet_name= 'by_Sales_Item', index=True,header=True)




# Pivot method which list by Ship to country and Sales Name and summary them.
pv_contrySale = solddata.pivot_table(index=['Ship To','Name'],aggfunc=[np.sum])

# print(pv_contrySale)
pv_contrySale.to_excel(writer, sheet_name= 'by_Country_Sales', index=True,header=True)

# Test povit with 3 columns:
pv_ConSalIte = solddata.pivot_table(index=['Ship To','Name','Itemno'],aggfunc=[np.sum])
pv_IteConSal = solddata.pivot_table(index=['Itemno','Ship To','Name'],aggfunc=[np.sum])
pv_CatIteConSal = solddata.pivot_table(index=['Cat','Itemno','Ship To','Name'],aggfunc=[np.sum])

# print(pv_ConSalIte)
pv_ConSalIte.to_excel(writer, sheet_name= 'by_Country_Sales_Item', index=True,header=True)
pv_CatIteConSal.to_excel(writer, sheet_name= 'by_Cat_Item_Country_Sales', index=True,header=True)


# writer.close()


# pv_contrySale.plot(kind='bar',figsize=(8,6),subplots=True,rot=270, fontsize=6)
# plt.subplots_adjust(bottom = 0.3)
# plt.tight_layout()
# plt.savefig(f'{newDir}/Sale_Country_%s.png' %filename)
#plt.show()

print("=============================================")
print("Up to date total sold Systems Qty: ", solddata['Qty'].sum())
print("Up to date total sold Nodes total: ", solddata['Node Qty'].sum())
print("=============================================")
print(top_item.head())
# Write a file for daily update;
solddata.to_excel('daily_output.xlsx')

# Check new items which missing CAT
naCat = solddata[solddata['Cat'].isnull()]

naCat = naCat.drop(['Date','Qty','Node Qty','Sales','Ship To','Region','Order Type','Rep code','Name','Supervisor'], axis=1)

naCat = naCat.drop_duplicates(subset = 'Itemno')
print(naCat)
itemno_new = itemno_cat.append(naCat).reset_index(drop=True)

itemno_new.to_excel('itemno_cat.xlsx', index=False)
writer.close()
plt.show()