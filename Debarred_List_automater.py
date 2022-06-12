from tracemalloc import start
import pandas as pd 

filepath=input("Enter the File path : ")

# input for creating attendance sheets
range_input=list(map(int,input("Enter range list : ").split(" ")))
#print(filepath)

data_orig = pd.read_excel(filepath)

#creating attendance sheets and index for attendance percentagw
total_ent = []
unique_ent = []
attendance = []                                   # reading excel file by its path

for i in range_input:
    sh_name = '<'+str(i)+' ATTENDANCE'
    select=data_orig[data_orig['ATTENDANCE PERCENTAGE'].between(0,i)]          # Sorting of elements based on the range inputs from user
    print(select)                                                # displaying selected values
    writer_new=pd.ExcelWriter(path=filepath, if_sheet_exists = 'replace',mode='a',engine='openpyxl')  
    select.to_excel(excel_writer=writer_new, sheet_name = sh_name)                         # appending the selected values on the specific sheet of the excel file
    writer_new.save() 

    total_ent.append(select.shape[0])
    unique_ent.append(select.nunique( axis = 'rows')['REG. NO.'])
    attendance.append(sh_name)
    
    writer_new.save()
    
try:    
    data_index = pd.read_excel(filepath, sheet_name = 'Index')
    row_count = data_index.shape[0] 

except:
    row_count = 0

writer_ovr = pd.ExcelWriter(path=filepath, if_sheet_exists = 'overlay', mode='a',engine='openpyxl')
index_df = pd.DataFrame({'Total Number of Entries': total_ent, 'No. of Students': unique_ent, 'Attendance': attendance})
if row_count == 0:
    index_df.to_excel(excel_writer=writer_ovr, sheet_name = 'Index', index=False)
else:
    index_df.to_excel(excel_writer=writer_ovr, sheet_name = 'Index', index=False, startrow=row_count+1, header = False)

writer_ovr.save()
writer_new.close() 
writer_ovr.close()


#creating attendance sheets and index and avg attendance


std_regno = ''
att_sum = 0
count = 0
avg_attendance = []

for ind in data_orig.index:
    regno = data_orig['REG. NO.'][ind]

    if regno == std_regno:
        att_sum += data_orig['ATTENDANCE PERCENTAGE'][ind]
        count += 1
    else:
        if std_regno != '':
            for _ in range(count):
                avg_attendance.append((round(att_sum/count)))

        std_regno = regno
        att_sum = data_orig['ATTENDANCE PERCENTAGE'][ind]
        count = 1

for _ in range(count):
    avg_attendance.append((round(att_sum/count)))

data_orig['ATTENDANCE AVG'] = avg_attendance


writer_new=pd.ExcelWriter(path=filepath, if_sheet_exists = 'replace', mode='a', engine='openpyxl')  
data_orig.to_excel(excel_writer=writer_new, sheet_name = 'ATTENDANCE_REPORT_FSPG', index=False)                         # appending the selected values on the specific sheet of the excel file
writer_new.save() 

total_ent = []
unique_ent = []
attendance = [] 
for i in range_input:
    sh_name = '<'+str(i)+' ATTENDANCE AND OVERALL_AVG <'+str(i)
    select=data_orig[data_orig['ATTENDANCE PERCENTAGE'].between(0,i)]
    select=data_orig[data_orig['ATTENDANCE AVG'].between(0,i)]         # Sorting of elements based on the range inputs from user
    print(select)                                                # displaying selected values  
    select.to_excel(excel_writer=writer_new, sheet_name = sh_name)                         # appending the selected values on the specific sheet of the excel file
    writer_new.save() 

    total_ent.append(select.shape[0])
    unique_ent.append(select.nunique( axis = 'rows')['REG. NO.'])
    attendance.append(sh_name)
    
    writer_new.save()

try:    
    data_index = pd.read_excel(filepath, sheet_name = 'Index')
    row_count = data_index.shape[0] 

except:
    row_count = 0

writer_ovr = pd.ExcelWriter(path=filepath, if_sheet_exists = 'overlay', mode='a',engine='openpyxl')
index_df = pd.DataFrame({'Total Number of Entries': total_ent, 'No. of Students': unique_ent, 'Attendance': attendance})
if row_count == 0:
    index_df.to_excel(excel_writer=writer_ovr, sheet_name = 'Index', index=False)
else:
    index_df.to_excel(excel_writer=writer_ovr, sheet_name = 'Index', index=False, startrow=row_count+1, header = False)

writer_ovr.save()
writer_new.close() 
writer_ovr.close()
