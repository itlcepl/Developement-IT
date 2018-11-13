import os
import pandas as pd
import xlrd
import csv
from datetime import time

filename = "G:\\scada\\FloatTable.csv"
directory = "G:\\scada\\To-Merge"
count = 0


def open_and_check(path, tocompare):
    xls = xlrd.open_workbook(path)
    sheet1 = xls.sheet_by_index(0)
    sheet2 = xls.sheet_by_index(1)
    sheet3 = xls.sheet_by_index(2)
    sheet4 = xls.sheet_by_index(3)

    count = 0

    with open(filename, "rb") as csvfile:
        datareader = csv.reader(csvfile)
        next(datareader)
        for i in range(7, sheet1.nrows - 1):
            time1 = int(sheet1.cell_value(i, 0) * 24)
            time1 = time(time1, 00)
            time1 = str(time1)
            if type(sheet1.cell_value(i, 1))!='float' or type(sheet1.cell_value(i+1, 1))!='float':
                next(datareader)

            initial_value = sheet1.cell_value(i, 1)
            print("Initial value: ",type(sheet1.cell_value(i+1, 1)))
            upto_value = sheet1.cell_value(i+1, 1)
            #print("Upto value", upto_value)
            differenece = upto_value - initial_value
            #print(differenece)
            #print(time1)
            for row in datareader:
                #print(row[0])
                if row[0][:10] == tocompare:
                    count = count + 1


    print("Count for file "+path+"is : ",count)

data = pd.read_csv(filename)
print(data.count())
for file in os.listdir(directory):
    if file.endswith(".xls"):
        path = os.path.join(directory, file)
        print("Current file is :: " + path)
        tocompare = file[:10]
        print(tocompare)
        #open_and_check(path, tocompare)
