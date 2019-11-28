import sys
import pandas as pd
from bs4 import BeautifulSoup

def read_excel_xml(path):
    file = open(path,'r',encoding="utf-8").read()
    soup = BeautifulSoup(file,'xml')
    workbook = []
    print("Reading data from .xls file")
    for sheet in soup.findAll('Table'): 
        sheet_as_list = []
        for row in sheet.findAll('Row'):
            row_as_list = []
            for cell in row.findAll('Cell'):
                row_as_list.append(cell.Data.text)
            sheet_as_list.append(row_as_list)
        workbook.append(sheet_as_list)
    return workbook

data = read_excel_xml(sys.argv[1])
filename = sys.argv[1].split()
eventID = filename[2].split('.')[0]
print("Converting data and writing .csv file")
df = pd.DataFrame(data[7][1:],columns = data[7][0])
df["Navn"] = df["Fornavn"] + " " + df["Etternavn"]
df["Status"] = "x"
df["Spesial"] = "A," + eventID + ",18," + df["Organisasjons-id"] + "," + df["Person-id"]
df = df[["Person-id","Status","Navn","Klasse","Klubb","Spesial","Emit"]]
outfile = "EntryDB " + str(eventID) + ".csv"
df.to_csv(outfile,sep=";",header=False,index=False,encoding="iso-8859-1")
