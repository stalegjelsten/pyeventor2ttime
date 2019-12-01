import sys
import pandas as pd
from bs4 import BeautifulSoup

def read_IOF_xml(path):
    # parse XML file and output as a list of lists

    # open file and choose parser
    soup = open_and_soup(path)

    # initialize the entries table
    entries = [["Person-id","Fornavn","Etternavn","Organisasjons-id","Klubb","Emit","Klasse"]]

    # The Eventor event ID is the first Id tag in the file
    event_ID = soup.find("Id").text

    for personentry in soup.findAll("PersonEntry"):
        # looping through every personentry in the file

        # finding the children of personentry in the .xml file
        personal = personentry.find("Person")
        organisation = personentry.find("Organisation")
        classtag = personentry.find("Class")

        # trying to extract personentry data if available
        try:
            person_id = personal.find("Id").text
        except Exception:
            person_id = ""
       
        try:
            organisasjons_id = organisation.find("Id").text
            klubb = organisation.find("Name").text
        except Exception:
            # setting organisation id to 999 and team to NOTEAM when
            # club is unavailable
            organisasjons_id = "999"
            klubb = "NOTEAM"

        try: 
            ecard = personentry.find("ControlCard").text
        except Exception:
            # setting emit ecard to 999 for entries w/o ecard
            ecard = "999"

        try:
            klasse = classtag.find("Name").text
        except Exception:
            klasse = "NOCLASS"

        # every entry needs a given and family name, so no test required
        fornavn = personal.find("Given").text
        etternavn = personal.find("Family").text

        # appending all the personal data to the entries table
        entries.append([person_id, fornavn, etternavn, organisasjons_id, klubb, ecard, klasse])

    return entries, event_ID



def read_excel_xml(path):
    # Read Excel XML file and output as list of lists

    # open file and choose parser
    soup = open_and_soup(path)

    # initialize entries table
    entries = []

    # The entries are stored in worksheet no. 7
    sheet = soup.findAll('Worksheet')[7]

    for row in sheet.findAll('Row'):
        # Looping over all the rows in the worksheet
        rows_as_list = []

        for cell in row.findAll('Cell'):
            # Looping over all the cells in the row

            if cell.Data == None:
                # if the cell does not cotain any data,
                # append empty cell to rows_as_list
                rows_as_list.append("")

            else:
                # if the cell contains data, append data
                rows_as_list.append(cell.Data.text)

        entries.append(rows_as_list)    
    return entries

def open_and_soup(path):
    # open file and parse it with beautifulsoup
    # Open the file as utf-8
    try:
        file = open(path,'r',encoding="utf-8").read()
        print("Reading data from file.")
    except Exception:
        print("Input file not found.")
        return False

    return BeautifulSoup(file,'xml')

def get_event_ID(infile):
    # Split the file name at spaces and full stops to get the eventor eventID
    filename = infile.split()
    return filename[2].split('.')[0]

def get_file_type(infile):
    # Split the file name at the last full stop to get file type
    filename = infile.split('.')
    return filename[-1]

def create_dataframe(data,event_ID):
    # Create a dataframe from the data

    if data == False:
        # Return if the input is empty
        return
    
    print("Converting data and writing .csv file")

    # The first row contains the column headers.
    df = pd.DataFrame(data[1:],columns = data[0])

    number_of_entries = len(df.index)

    # Concatenate first and last names
    df["Navn"] = df["Fornavn"] + " " + df["Etternavn"]

    # 'x' is a status flag in ttime which says the participant is entered.
    df["Status"] = "x"

    # Creating a column for ttime special data. Format string in next line.
    # A : Eventor ID's: event ID, organization ID, club ID, runner ID, class ID
    df["Spesial"] = "A:" + event_ID + ",18," + df["Organisasjons-id"] + "," + df["Person-id"]

    # Reordering and dropping columns to conform with ttime database format
    df = df[["Person-id","Status","Navn","Klasse","Klubb","Spesial","Emit"]]

    # Choosing file name for csv database
    outfile = "EntryDB " + str(event_ID) + ".csv"

    # Writing dataframe to csv file with correct encoding for ttime.
    # If you want to change the csv separator, you also need to change the
    # commas in the ttime special data field above.
    df.to_csv(outfile, sep=";", header=False, index=False, encoding="iso-8859-1")

    print("Done." , number_of_entries , "entries written to file" , outfile)


def gui_open_file():
    # open a GUI file picker
    import tkinter
    from tkinter import filedialog
    tkinter.Tk().withdraw()
    filename = filedialog.askopenfilename(initialdir = "~",title = "Select file",filetypes = [("xml or xls files","*.xml *.xls")])
    return filename

if __name__ == "__main__":

    # open GUI file picker if no file or multiple files are specified
    filename = sys.argv[1] if len(sys.argv) == 2 else gui_open_file()
    filetype = get_file_type(filename)

    if filetype == "xml":
        data, event_ID = read_IOF_xml(filename)
    elif filetype == "xls":
        data = read_excel_xml(filename)
        event_ID = get_event_ID(filename)
    
    create_dataframe(data,event_ID)
