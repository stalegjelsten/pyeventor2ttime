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

    # The entries are stored in worksheet and table no. 7
    sheet = soup.findAll('Table')[7]

    for row in sheet.findAll('Row'):
        rows_as_list = []
        for cell in row.findAll('Cell'):
            rows_as_list.append(cell.Data.text)
        workbook.append(rows_as_list)    
    return workbook

# Split the file name at spaces and full stops to get the eventor eventID of the event
def getEventID():
    filename = sys.argv[1].split()
    return filename[2].split('.')[0]

# Create a dataframe from the data
def createDataframe(data,eventID):
    print("Converting data and writing .csv file")

    # The first row contains the column headers.
    df = pd.DataFrame(data[1:],columns = data[0])

    # Concatenate first and last names
    df["Navn"] = df["Fornavn"] + " " + df["Etternavn"]

    # The 'x' flag means entered in ttime
    df["Status"] = "x"

    # Creating a column for ttime special data. Format string in next line.
    # A : Eventor ID's: event ID, organization ID, club ID, runner ID, class ID
    df["Spesial"] = "A:" + eventID + ",18," + df["Organisasjons-id"] + "," + df["Person-id"]

    # Reordering and dropping columns to conform with ttime database format
    df = df[["Person-id","Status","Navn","Klasse","Klubb","Spesial","Emit"]]

    # Choosing file name for csv database
    outfile = "EntryDB " + str(eventID) + ".csv"

    # Writing dataframe to csv file with correct encoding for ttime.
    # If you want to change the csv separator, you also need to change the
    # commas in the ttime special data field above.
    df.to_csv(outfile,sep=";",header=False,index=False,encoding="iso-8859-1")

data = read_excel_xml(sys.argv[1])
eventID = getEventID()
createDataframe(data,eventID)
