import pandas as pd

import math


number_of_members = 0
sm =['Sheet1','Sheet3','Sheet4','Sheet5']                  # make sure to add the names of all the sheets here under sheetname variable sm.
cn = ['HQ','P BTY','Q BTY','R BTY']                           # make sure to add all the names of the companies.
count=0 
m=0

writer = pd.ExcelWriter('final_attendence.xlsx', engine='xlsxwriter')          # instead of 'final.xlsx', enter path of the new file to be created (This is where the data has to be saved)
for i in sm:
    
    status=[]
    df = pd.read_excel('test_final_1.xlsx', sheet_name= i)             # instead of 'attendence.xlsx', enter path of the attendance sheet from which data has to be taken.
    

    #convert each column to list, remove NAN values 
    name = df.columns[3]
    filtered_column_names = df.loc[9:,name].loc[~df[name].isin(['PARBAT ALI','Name'])].tolist()
    ecode = df.columns[2]
    filtered_column_code = df.loc[9:,ecode].loc[~df[ecode].isin(['E. Code'])].tolist()
    TD_time = df.columns[10]                                                                                #TD Time is same as A in time
    filtered_column_TD = df.loc[9:,TD_time].loc[~df[TD_time].isin(['A. InTime'])].fillna(-1).tolist()
    SE_time = df.columns[8]                                                                                #SE time is same as  S out time
    filtered_column_SE = df.loc[9:,SE_time].loc[~df[SE_time].isin(['S. OutTime'])].fillna(-1).tolist()
    Sin = df.columns[6]
    filtered_column_Sin = df.iloc[9:df[Sin].last_valid_index() + 1, df.columns.get_loc(Sin)].fillna(-1).tolist()



    size = len(filtered_column_names)
    members =[]
    x=0
    while x <size: 
        members.append(cn[count])
        x+=1
    number_of_members+=size
    count=count+1

    for j in filtered_column_Sin: 
        if j == -1 :
            status.append('Absent')
        else :
            status.append('Present')
    for k in filtered_column_TD:
       
        if k== -1 :
            if status[filtered_column_TD.index(k)] != 'Absent' :
                status[filtered_column_TD.index(k)] = 'Absent'

    data = {
        'Company' : members,
        'E. Code': filtered_column_code,
        'Name' : filtered_column_names,
        'TD' :  filtered_column_TD,
        'S IN' : filtered_column_Sin,
        'S OUT' : filtered_column_SE,
        'Status' : status
        }

    
    # now creating saving the data into a seperate sheet per company name 
    fs = pd.DataFrame(data)
    fs.to_excel(writer,sheet_name= cn[m] , index=False)
    m+=1
writer._save()



'''
#sheet2 LRW
status=[]
df = pd.read_excel('test_final_1.xlsx', sheet_name='Sheet2')
df_filtered = df[df.iloc[:, 3] != 'PARBAT ALI']

#convert each column to list, remove NAN values 
name = df_filtered.columns[3]
filtered_column_names = df_filtered.loc[9:,name].loc[~df_filtered[name].isin(['Name'])].dropna().tolist()
ecode = df_filtered.columns[2]
filtered_column_code = df_filtered.loc[9:,ecode].loc[~df_filtered[ecode].isin(['E. Code'])].dropna().tolist()
TD_time = df_filtered.columns[10]                                                                                #TD Time is same as A in time
filtered_column_TD = df_filtered.loc[9:,TD_time].loc[~df_filtered[TD_time].isin(['A. InTime'])].dropna().tolist()
SE_time = df_filtered.columns[8]                                                                                #SE time is same as  S out time
filtered_column_SE = df_filtered.loc[9:,SE_time].loc[~df_filtered[SE_time].isin(['S. OutTime'])].dropna().tolist()
Sin = df_filtered.columns[6]
filtered_column_Sin = df_filtered.loc[9:,Sin].loc[~df_filtered[Sin].isin(['S. InTime'])].dropna().tolist()


#finding total number of members in the company
size = len(filtered_column_names)
members =[]
x=0
while x <size:
    members.append('LRW')
    x+=1

for j in filtered_column_Sin: 
        if j == -1 :
            status.append('Absent')
        else :
            status.append('Present')
for k in filtered_column_TD:
        #print(k)
        #print(type(k))
        #print('---------------')
        if k== -1 :
            if status[filtered_column_TD.index(k)] != 'Absent' :
                status[filtered_column_TD.index(k)] = 'Absent'

#adding their names to final excell sheet
data = {
        'Company' : members,
        'E. Code': filtered_column_code,
        'Name' : filtered_column_names,
        'TD' :  filtered_column_TD,
        'S IN' : filtered_column_Sin,
        'S OUT' : filtered_column_SE,
        'Status' : status
        }
fs = pd.DataFrame(data)
fs.to_excel(writer,sheet_name= 'sheet2', index=False)
writer._save()
'''