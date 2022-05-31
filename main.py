from numpy import matrix
import pandas
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

    matrixDF = pandas.read_excel(document)

    matrixDF = matrixDF[['Day', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']]

    matrixDF.to_excel('Output/exported.xlsx')

main()