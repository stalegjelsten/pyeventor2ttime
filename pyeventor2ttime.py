import sys
import pandas as pd
from bs4 import BeautifulSoup

# Read Excel XML file as list of lists
def read_excel_xml(path):
    print("Reading data from .xls file")

    # Open the file as utf-8
    file = open(path,'r',encoding="utf-8").read()

    soup = BeautifulSoup(file,'xml')
    workbook = []
    for sheet in soup.findAll('Table'): 
        sheet_as_list = []
        for row in sheet.findAll('Row'):
            row_as_list = []
            for cell in row.findAll('Cell'):
                row_as_list.append(cell.Data.text)
            sheet_as_list.append(row_as_list)
        workbook.append(sheet_as_list)
    return workbook

# Split the file name at spaces and full stops to get the eventor eventID of the event
def getEventID():
    filename = sys.argv[1].split()
    return filename[2].split('.')[0]

# Create a dataframe from the data
def createDataframe(data,eventID):
    print("Converting data and writing .csv file")

    # The eventor entries are stored in sheet 7. The first row contains the column headers.
    df = pd.DataFrame(data[7][1:],columns = data[7][0])

    # Concatenate first and last names
    df["Navn"] = df["Fornavn"] + " " + df["Etternavn"]

    # The 'x' flag means entered in ttime
    df["Status"] = "x"

    # Creating a column for ttime special data
    df["Spesial"] = "A," + eventID + ",18," + df["Organisasjons-id"] + "," + df["Person-id"]

    # Reordering and dropping columns to conform with ttime database formatj
    df = df[["Person-id","Status","Navn","Klasse","Klubb","Spesial","Emit"]]

    # Choosing file name for csv database
    outfile = "EntryDB " + str(eventID) + ".csv"

    # Writing dataframe to csv file with correct encoding for ttime.
    df.to_csv(outfile,sep=";",header=False,index=False,encoding="iso-8859-1")


data = read_excel_xml(sys.argv[1])
eventID = getEventID()
createDataframe(data,eventID)

