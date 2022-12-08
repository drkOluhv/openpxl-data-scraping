import os
import pandas as  pd 
import openpyxl

# Select the directory where the Excel files are located
os.chdir('./Test')

#sheet = wb.get_sheet_by_name('Sheet1')
        
# Loop function 
for root, dirs, files in os.walk('.'):
    attributes = ['Total Consolidation Settlement excl creep to 40yrs']
    #initialise dictionary, create empty list for attributes with dict comprehension 
    data = {attribute: [] for attribute in attributes}
    #append a key:value for File, will use this as unique identifier/index
    data.update({'File': []})
    for file in files:
        wb = openpyxl.load_workbook(file)
        ws = wb["Sheet1"]
        data['File'].append(file)
        for attribute in attributes:
            data[attribute].append(ws.cell(row=2,column=2).value)
data

df =pd.DataFrame.from_dict(data)
df.to_excel('Scraped_Data.xlsx',sheet_name='Sheet1')