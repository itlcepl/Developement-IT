import os
import pandas as pd
import xlrd
import csv
from datetime import time
from datetime import datetime
from datetime import datetime
from datetime import timedelta

Flow_Table_CSV = "G:\\scada\\subset.csv"
Excel_File_Directory = "G:\\scada\\To-Merge"


def check_last_data(date_compare):
    date1 = datetime.strptime(date_compare, '%d-%m-%Y').date()
    date1 = date1 + timedelta(days=1)
    date1 = str(date1)
    date_compare = date1[8:] + date1[4:8] + date1[:4]

def change_actual_data(sheet, comapre_date, csv_data, mapped_sheet):
    for i in range(7, sheet.nrows):

        #print("Cell Value is",sheet.cell_value(i, 0))
        if sheet.cell_value(i, 0) == "24:00":
            #check_last_data(comapre_date)
            date1 = datetime.strptime(comapre_date, '%d-%m-%Y').date()
            date1 = date1 + timedelta(days=1)
            date1 = str(date1)
            comapre_date = date1[8:] + date1[4:8] + date1[:4]
            time_to_compare = "00:00:00"
        else:

            time_to_compare = int(sheet.cell_value(i, 0) * 24)
            time_to_compare = str(time(time_to_compare , 00))
            print(time_to_compare[:3])

        for index, rows in mapped_sheet.iterrows():
            print("Printing Data of column " + mapped_sheet.iloc[index]['Title'])
            print("Tag Index is  ", mapped_sheet.iloc[index]['TagIndex'])
            column = mapped_sheet.iloc[index]['Index In Excel Sheet']
            Tag_Index = mapped_sheet.iloc[index]['TagIndex']
            initial_value = sheet.cell_value(i, column)
            try:
                upto_value = sheet.cell_value(i + 1, column)
            except IndexError:
                upto_value = initial_value

            time_matched = csv_data.index[csv_data['DateAndTime'].str.match(comapre_date + " " + time_to_compare[:3])].tolist()
            add_value = (upto_value - initial_value) / len(time_matched)
            print(add_value)
            print("Initial Value is ", initial_value)
            csv_data.Val.iloc[[time_matched[0]]] = initial_value
            print("Updated value for time: ", csv_data.iloc[time_matched[0]]['DateAndTime'], " is ",csv_data.iloc[time_matched[0]]['Val'])
            for each_index in time_matched[1:]:
                if csv_data.iloc[each_index]['TagIndex'] == Tag_Index:
                    initial_value = initial_value + add_value
                    csv_data.Val.iloc[[each_index]] = initial_value

                    print(
                         "Updated value for time: ", csv_data.iloc[each_index]['DateAndTime'], " is ",
                          csv_data.iloc[each_index]['Val'])

data = [['RAW WATER INLET FLOW', 1, 16], ['TOTALIZED INLET FLOW', 2, 69], ['PH AT RW IN', 3, 17],
        ['TURBIDITY AT RW IN', 4, 18], ['RESIDUAL CLORINE', 5, 28], ['TURBIDITY AT CW OUT', 6, 29],
        ['BACKWASH TANK LEVEL', 7, 20]]

map_sheet1 = pd.DataFrame(data, columns=['Title', 'Index In Excel Sheet', 'TagIndex'])

data = [['HRS-1', 1, 118], ['Min-1', 2, 119], ['HRS-2', 3, 120], ['Min-2', 4, 121], ['HRS-3', 5, 122],
        ['Min-3', 6, 123],
        ['HRS-4', 7, 0], ['Min-4', 8, 0], ['HRS-5', 9, 0], ['Min-5', 10, 0], ['HRS-6', 11, 0], ['Min-6', 12, 0],
        ['HRS-7', 13, 38], ['Min-7', 14, 39], ['HRS-8', 15, 40], ['Min-8', 16, 41], ['HRS-9', 17, 42],
        ['Min-9', 18, 43]]
map_sheet2 = pd.DataFrame(data, columns=['Title', 'Index In Excel Sheet', 'TagIndex'])

data = [['HRS-1', 1, 44], ['Min-1', 2, 45], ['HRS-2', 3, 46], ['Min-2', 4, 47], ['HRS-3', 5, 53], ['Min-3', 6, 54],
        ['HRS-4', 7, 55], ['Min-4', 8, 56], ['HRS-5', 9, 57], ['Min-5', 10, 58], ['HRS-6', 11, 0], ['Min-6', 12, 0],
        ['HRS-7', 13, 61], ['Min-7', 14, 62], ['HRS-8', 15, 63], ['Min-8', 16, 64], ['HRS-9', 17, 65],
        ['Min-9', 18, 66],
        ['HRS-10', 13, 67], ['Min-10', 14, 68]]

map_sheet3 = pd.DataFrame(data, columns=['Title', 'Index In Excel Sheet', 'TagIndex'])

data = [['HRS-1', 1, 30], ['Min-1', 2, 31], ['HRS-2', 3, 32], ['Min-2', 4, 33], ['HRS-3', 5, 49], ['Min-3', 6, 50],
        ['HRS-4', 7, 51], ['Min-4', 8, 52], ['HRS-5', 113, ], ['Min-5', 10, 114], ['HRS-6', 11, 115],
        ['Min-6', 12, 116],
        ['HRS-7', 13, 74], ['Min-7', 14, 75], ['HRS-8', 15, 76], ['Min-8', 16, 77], ['HRS-9', 17, 78],
        ['Min-9', 18, 79],
        ['HRS-10', 19, 71], ['Min-10', 20, 72], ['HRS-11', 21, 0], ['Min-11', 22, 0], ['HRS-12', 23, 90],
        ['Min-12', 24, 91],
        ['HRS-13', 25, 92], ['Min-13', 26, 93], ['HRS-14', 27, 94], ['Min-14', 28, 95], ['HRS-15', 29, 0],
        ['Min-15', 30, 0],
        ['HRS-16', 31, 82], ['Min-16', 32, 83], ['HRS-17', 33, 84], ['Min-17', 34, 85], ['HRS-18', 35, 0],
        ['Min-18', 36, 0],
        ['HRS-19', 37, 0], ['Min-19', 38, 0], ['HRS-20', 39, 100], ['Min-20', 40, 101], ['HRS-21', 41, 102],
        ['Min-21', 42, 102]
        ]

map_sheet4 = pd.DataFrame(data, columns=['Title', 'Index In Excel Sheet', 'TagIndex'])


def analysis_seperate_sheet(file_path, comapre_date, csv_data):
    xls = xlrd.open_workbook(file_path)
    sheet1 = xls.sheet_by_index(0)
    sheet2 = xls.sheet_by_index(1)
    sheet3 = xls.sheet_by_index(2)
    sheet4 = xls.sheet_by_index(3)
    change_actual_data(sheet1, comapre_date, csv_data, map_sheet1)
    change_actual_data(sheet2, comapre_date, csv_data, map_sheet2)
    # change_actual_data(sheet3, comapre_date, csv_data, map_sheet3)
    # change_actual_data(sheet4, comapre_date, csv_data, map_sheet4)


csv_data = pd.read_csv(Flow_Table_CSV)
csv_data.sort_values(by=['DateAndTime'])

for file in os.listdir(Excel_File_Directory):
    if file.endswith(".xls"):
        path = os.path.join(Excel_File_Directory, file)
        print("Current file is :: " + file)
        file_to_compare = file[:10]
        analysis_seperate_sheet(path, file_to_compare, csv_data)
