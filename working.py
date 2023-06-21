import pandas as pd
from openpyxl import Workbook, load_workbook

#df =pd.ExcelFile('attendence.xls')
#ws = pd.read_excel(df,'Sheet2')
#print(ws)
number_of_members = 0

#sheet1 HQ
df = pd.read_excel('attendence.xls', sheet_name='Sheet1')

#convert each column to list, remove NAN values 
name = df.columns[3]
filtered_column_names = df.loc[9:,name].loc[~df[name].isin(['Name'])].dropna().tolist()
ecode = df.columns[2]
filtered_column_code = df.loc[9:,ecode].loc[~df[ecode].isin(['E. Code'])].dropna().tolist()
TD_time = df.columns[10]                                                                                #TD Time is same as A in time
filtered_column_TD = df.loc[9:,TD_time].loc[~df[TD_time].isin(['A. InTime'])].dropna().tolist()
SE_time = df.columns[8]                                                                                #SE time is same as  S out time
filtered_column_SE = df.loc[9:,SE_time].loc[~df[SE_time].isin(['S. OutTime'])].dropna().tolist()
status = df.columns[17]
filtered_column_status = df.loc[9:,status].loc[~df[status].isin(['Status'])].dropna().tolist() 

#finding total number of members in the company
size = len(filtered_column_names)
members =[]
x=0
while x <size:
    members.append('HQ')
    x+=1
number_of_members+=size

#adding their names to final excell sheet
data = {
    'Company' : members,
    'E. Code': filtered_column_code,
    'Name' : filtered_column_names,
    'TD' :  filtered_column_TD,
    'SE' : filtered_column_SE,
    'Status' : filtered_column_status
    }

fs = pd.DataFrame(data)
fs.to_excel('final.xlsx', index=False)

'''
#sheet2 LRW
df = pd.read_excel('attendence.xls', sheet_name='Sheet2')

#convert each column to list, remove NAN values 
name = df.columns[3]
filtered_column_names = df.loc[9:,name].loc[~df[name].isin(['PARBAT ALI','Name'])].dropna().tolist()
ecode = df.columns[2]
filtered_column_code = df.loc[9:,ecode].loc[~df[ecode].isin(['E. Code'])].dropna().tolist()
TD_time = df.columns[10]                                                                                #TD Time is same as A in time
filtered_column_TD = df.loc[9:,TD_time].loc[~df[TD_time].isin(['A. InTime'])].dropna().tolist()
SE_time = df.columns[8]                                                                                #SE time is same as  S out time
filtered_column_SE = df.loc[9:,SE_time].loc[~df[SE_time].isin(['S. OutTime'])].dropna().tolist()
status = df.columns[17]
filtered_column_status = df.loc[9:,status].loc[~df[status].isin(['Status'])].dropna().tolist() 

#finding total number of members in the company
size = len(filtered_column_names)
members =[]
x=0
while x <size:
    members.append('LRW')
    x+=1

#adding their names to final excell sheet

existing_column_values = fs['Company'].tolist()
existing_column_values.extend(members)
print(len(existing_column_values))
#print(fs['Company'])
fs['Company'] = existing_column_values



#data = {
#    'Sl No' : [],
#    'Company' : members,
#    'E. Code': filtered_column_code,
#    'Name' : filtered_column_names,
#    'TD' :  filtered_column_TD,
#    'SE' : filtered_column_SE,
#    'Status' : filtered_column_status
#    }
#
#fs = pd.DataFrame(data)
#fs.to_excel('final.xlsx', index=False)

#print(filtered_column_code)

#name = df.iloc[:, 3]
#print(name)
#column = df.iloc[:, 6]
#column_list = column.dropna().tolist()
#print(column_list)
'''