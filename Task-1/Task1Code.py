import ClointFusion as cf

def dateformat(rowend,name):
    path = "D:\CreatedDocuments\\"+name
    split = cf.excel_copy_range_from_sheet(excel_path=path,sheet_name="Sheet1", startRow=2, endRow=rowend, startCol=2, endCol=2)
    # print(split)
    temp=0
    for i in range(len(split)):
        try:
            split[i][0] = str(split[i][0]) #converting the data type to string from datetime data type so that we can change it
            split[i][0] = split[i][0][8:10]+"-"+split[i][0][5:7]+"-"+split[i][0][0:4]
            #removing the last 00:00:00
        except:
            # if any unexpected error happens then this will be printed in terminal
            print("Cannot remove the time from OrderDate column")
            temp=1
    if(temp == 0):
        cf.excel_copy_paste_range_from_to_sheet(excel_path=path, sheet_name="Sheet1",startCol=2, endCol=2, startRow=2, endRow=rowend, copiedData=split)

count = cf.excel_get_row_column_count(excel_path="D:\OriginalDocuments\Excel.xlsx", sheet_name="Sheet1")

# copying the data in the given excel file
sheetdata = cf.excel_copy_range_from_sheet(excel_path="D:\OriginalDocuments\Excel.xlsx", sheet_name="Sheet1",startCol=1, endCol=count[1], startRow=1, endRow=count[0])

# creating a folder 
cf.folder_create(strFolderPath="D:\CreatedDocuments")

# creating a excel file 
cf.excel_create_excel_file_in_given_folder(fullPathToTheFolder="D:\CreatedDocuments", excelFileName="Excel.xlsx")

# pasting the copied data into created excel file 
cf.excel_copy_paste_range_from_to_sheet(excel_path="D:\CreatedDocuments\Excel.xlsx", sheet_name="Sheet1",startCol=1, endCol=count[1], startRow=1, endRow=count[0], copiedData=sheetdata)

# 1. Getting all the header columns.
headerName = cf.excel_get_all_header_columns( excel_path="D:\CreatedDocuments\Excel.xlsx", sheet_name="Sheet1")
print("headerNames are : ",end=" ")
print(headerName)

# 2. Getting Row & Column count.
count = cf.excel_get_row_column_count(excel_path="D:\CreatedDocuments\Excel.xlsx", sheet_name="Sheet1")
print("(row,column) : ",end=" ")
print(count)

# 3. Getting all sheet names in the ‘Excel.xlsx’.
sheetnames = cf.excel_get_all_sheet_names(excelFilePath="D:\CreatedDocuments\Excel.xlsx")
print("sheet names are : ",end=" ")
print(sheetnames)

# 4. Remove the duplicate data w.r.t ‘ID’ column.
cf.excel_remove_duplicates(excel_path="D:\CreatedDocuments\Excel.xlsx",sheet_name="Sheet1", header=0, columnName="ID ", saveResultsInSameExcel=True)

# 5. Sort the data w.r.t ‘OrderDate’ column.
cf.excel_sort_columns(excel_path="D:\CreatedDocuments\Excel.xlsx", sheet_name="Sheet1", header=0,firstColumnToBeSorted="OrderDate")

# 6. Store the following data in a python dictionary and insert the data at the last row respectively
dict = {"ID ": 1027,"OrderDate": "4/14/2020","Region": "East","Rep": "Jones","Item": "Binder","Units": 60,
"UnitCost": 4.99,"Total": 449.1}
count2 = cf.excel_get_row_column_count(excel_path="D:\CreatedDocuments\Excel.xlsx", sheet_name="Sheet1")
for i in dict.keys():
    cf.excel_set_single_cell(excel_path="D:\CreatedDocuments\Excel.xlsx",sheet_name="Sheet1",header=0,columnName=i,cellNumber=count2[0]-1,setText=dict[i])

# 7.	Split the excel on row count ‘12’. 
cf.excel_split_the_file_on_row_count(excel_path="D:\CreatedDocuments\Excel.xlsx", sheet_name="Sheet1",rowSplitLimit=12, outputFolderPath="D:\CreatedDocuments", outputTemplateFileName="Split")

# changing the OrderDate column
dateformat(13,"Split-1.xlsx")
dateformat(13,"Split-2.xlsx")
dateformat(13,"Split-3.xlsx")
dateformat(9,"Split-4.xlsx")


# 8.	Create a python dictionary named ‘data’ such that it stores the ‘ID’ and ‘Units’ of each row data in the excel file
data={}

id = cf.excel_copy_range_from_sheet(excel_path="D:\CreatedDocuments\Excel.xlsx",
sheet_name="Sheet1",startCol=1, startRow=2, endCol=1, endRow=count2[0]+1)

units = cf.excel_copy_range_from_sheet(excel_path="D:\CreatedDocuments\Excel.xlsx",
sheet_name="Sheet1",startCol=6, startRow=2,endCol=6, endRow=count2[0]+1 )

for i in range(len(id)):
    if id[i][0] not in data:
        data[id[i][0]] = units[i][0]
    else:
        data[id[i][0]] = [data[id[i][0]]]
        data[id[i][0]].append(units[i][0])
print("data dictionay is : ",end="")
print(data)

# changing the OrderDate column
dateformat(count2[0]+1,"Excel.xlsx")


