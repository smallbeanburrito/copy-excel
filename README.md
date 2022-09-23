# Created using python 3.9 with openpyxl library and standard sys and csv library
# Python script used to selectively copy columns and rows from Excel sheet to Excel sheet

# To use, call python3 copyExcel.py with 3 arguments or with an optional 4th argument
# arg1 is the column letter where the values of the column contain the value 'add' if the row is to be copied (will not be copied if contains any other value)
# arg2 is the file containing comma separated values with the names of each column to be copied (custom column label not the column letter)
# arg3 the path to source excel file you wish to copy
# optional arg4 the path to a destination excel file 

# example 
# 4 args
# python3 copyExcel.py 'A' .\copyColumns.txt .\source.xlsx .\des.xlsx
# 3 args
# python3 copyExcel.py 'A' .\copyColumns.txt .\source.xlsx

# if only 3 arguments are passed, a new excel file named 'destination.xlsx' will be created and used as the destination worksheet
