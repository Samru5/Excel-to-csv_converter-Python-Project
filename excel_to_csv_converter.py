# excel_to_csv_converter

import os
import csv
# openpyxl modeule helps to deal with excel files
import openpyxl


# Method to convert excel file into csv format
def excelToCsv(folder):
    # listdir returns all folders & files present in the given path
    for file in os.listdir(folder):

        # If file name doesn't ends with xlsx then again continue to find till we get xlsx file
        if not file.endswith('xlsx'):
            continue

        # load_workbook() takes in the filename and returns a value of the workbook data type. This Workbook object represents the Excel file
        myFile = openpyxl.load_workbook(file)

        # we get list of all the sheet names in the workbook by calling the get_sheet_names()
        for sheetName in myFile.get_sheet_names():

            # Each sheet is represented by a Worksheet object, which you can obtain by passing the sheet name string to the get_sheet_by_name() workbook method.
            eachSheet = myFile.get_sheet_by_name(sheetName)

            # Splitting the file name and taking only name of file without extension part & then appending it with each sheet title with csv extension
            csvFileName = file.split('.')[0] + '_' + eachSheet.title + '.csv'

            # Opening file in write mode
            csvFileObj = open(csvFileName, 'w', newline='')

            csvWriter = csv.writer(csvFileObj)

            # For each row in a sheet it will scan cells and will extract value of each cell
            for rowObj in eachSheet.rows:
                rowData = []
                for cellObj in rowObj:
                    rowData.append(cellObj.value)
                csvWriter.writerow(rowData)
        csvFileObj.close()

#Main method
if __name__ == "__main__":
    # Passing path of current folder
    excelToCsv('.')
