import pandas as pd
#from openpyxl import Workbook, load_workbook
#from xlsxwriter import Workbook

number_of_members = 0
sm =['Sheet1','Sheet2','Sheet3','Sheet4','Sheet5']                  # make sure to add the names of all the sheets here under sheetname variable sm.
cn = ['HQ','LRW','P BTY','Q BTY','R BTY']                           # make sure to add all the names of the companies.
count=0 
writer = pd.ExcelWriter('final_attendence.xlsx', engine='xlsxwriter')          # instead of 'final.xlsx', enter path of the new file to be created (where the data has to be saved)
for i in sm:
    
    df = pd.read_excel('attendence.xls', sheet_name= i)             # instead of 'attendence.xlsx', enter path of the attendance sheet from which data has to be taken.
    

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
  
    size = len(filtered_column_names)
    members =[]
    x=0
    while x <size: 
        members.append(cn[count])
        x+=1
    number_of_members+=size
    count=count+1

    #adding their names to final excell sheet
    data = {
        'Company' : members,
        'E. Code': filtered_column_code,
        'Name' : filtered_column_names,
        'TD' :  filtered_column_TD,
        'SE' : filtered_column_SE,
        'Status' : filtered_column_status
        }

    # now creating saving the data into a seperate sheet per company name 
    fs = pd.DataFrame(data)
    fs.to_excel(writer,sheet_name= i, index=False)
writer._save()