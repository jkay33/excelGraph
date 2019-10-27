##################################################
#           Parse xlsx/csv to graph              #
#                           By Jay Kim           #
##################################################
import os
import pandas as pd

##################################################
#  May make file finding process into function   #
##################################################
# Getting the current working directory
current = os.getcwd()

# Find the xlsx or csv file
folder_list = os.listdir(current)
for files in folder_list:
    if files.endswith('.xlsx') or files.endswith('.csv'):
        working_file = files
        break

# Determine what kind of file and create dataframe
if working_file.endswith('.xlsx'):
    df = pd.read_excel(working_file)
else:
    df = pd.read_csv(working_file)

# Accessing and creating exel writer to create workbook
excelFile = 'graphBook.xlsx'
sheetName = 'excelGraph'
writer = pd.ExcelWriter(excelFile, engine='xlsxwriter')

# inserting data in dataframe into created excel
df.to_excel(writer, sheet_name=sheetName, index=False)

# Accessing workbook and worksheet objects from df
workbook = writer.book
worksheet = writer.sheets[sheetName]

# Header information
headers = list(df)

# Create chart
chart = workbook.add_chart({'type': 'column'})

# Find length of all columns
# Assumption, the length of first column will be consistant throughout
col_length = len(df.iloc[:, 0])

# Setting location of last data
data_loc = str(col_length+1)

# Defining chart location on excel
chart_loc = str(col_length + 2)

# Defining data points for excel chart function
chart.add_series({
    'categories': '=' + sheetName + '!$A$2:$A$' + data_loc,
    'values': '=' + sheetName + '!$B$2:$B$' + data_loc,
    'name': headers[1]
    })
chart.add_series({
    'values': '=' + sheetName + '!$C$2:$C$' + data_loc,
    'name': headers[2]
    })
# Setting label
chart.set_x_axis({'name': headers[0]})
# Inserting chart in chart location
worksheet.insert_chart('A' + chart_loc, chart)
# Saving
writer.save()
