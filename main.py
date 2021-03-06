from numpy import matrix
from pandas import *
import os

def findDocument():
    #path =  os.path.realpath(__file__)
    directoryList = os.listdir('.')

    print('Which file would you like to use?')

    excelDocs = []

    for item in directoryList:
        if item.find('.~lock') != -1:
            continue 
        elif item.find('.xlsx') != -1:
            excelDocs.append(item)

    if len(excelDocs) < 1:
        print('Please move the course report to my folder!')
        exit()
    elif len(excelDocs) > 1:
        print('Which of the following files would you like to use?')
        for item in excelDocs:
            print(item)
    else:
        print('Using: ')
        for item in excelDocs: 
            print(item)

        return excelDocs[0]

def main():
    document = findDocument()

    matrixDF = read_excel(document)

    matrixDF = matrixDF[['Day', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']]
    matrixDF = matrixDF[matrixDF.Day != 'Time']

    #matrixDF.head(10)

    # Thank you to this blessing post:
    # https://stackoverflow.com/questions/61335563/how-to-reorder-of-rows-of-a-dataframe-based-on-values-in-a-column
    order = ['8:00am - 9:15am', '9:30am - 10:45am', '11:00am - 12:15pm', '12:30pm - 1:45pm', '2:00pm - 3:15pm', '3:30pm - 4:45pm', '5:00pm - 6:15pm', '6:30pm - 7:45pm', '8:00pm - 9:15pm']
    matrixDF = matrixDF.reindex(matrixDF['Day'].map(dict(zip(order, range(len(order))))).sort_values().index)

    matrixDF.to_excel('Output/exported.xlsx')

main()