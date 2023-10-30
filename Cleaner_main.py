import openpyxl as xl
import pandas as pd
import numpy as np


#setup
path=str(input("Path: ")).strip('"')
sheetname=str(input("Worksheet: "))
y_cord=str(input("Y coordinate of the column: ")).upper()
lenght=int(input("Lenght: "))

def __main__(path,sheetname,y_cord,lenght):

    al='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    #Creating a new list with only the float values
        #pandas to read dataframe
    data=pd.read_excel(path)
    new=[]
    #Iterating through the column
    for i in data.iloc[:,al.find(y_cord)]:
        strtofloat=""
        #Already floats dont get edited
        if isinstance(i,int) or isinstance(i,float) or i=='':
            new.append(float(i))
        #Strings get extra characters stripped and converted to floats
        elif isinstance(i,str):
            for p in i:
                if p in "1234567890.":
                    strtofloat=strtofloat+p
            new.append(float(strtofloat))
    new1=np.array(new)

    file=xl.load_workbook(path)
    sheet=file.get_sheet_by_name(sheetname)

    #Iterating through the columns and replacing values
    def change_cell_column(lenght):
        for i in range (2,lenght+2):
            sheet[f'{y_cord}{i}']=new[i-2]

    change_cell_column(lenght)

    file.save(path)

__main__(path,sheetname,y_cord,lenght)
