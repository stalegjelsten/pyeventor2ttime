# pyeventor2ttime

## About

A script to convert entries from [Eventor](https://eventor.orientering.no) to the [ttime](http://ttime.no/) database format.

The [ttime](http://ttime.no/) application for Windows is used for timekeeping and entries during orienteering events with [emit](https://emit.no) eCards. At the moment, there is no easy way to pull entries from the [Eventor](https://eventor.orientering.no) platform to ttime. This program is written to make that transition easier.

Previously, ttime could pull data from Eventor directly, but this is unfortunately not possible anymore.

## Downloading entries from Eventor

Go to [eventor.orientering.no](https://eventor.orientering.no) and log in. Find your race and click `Edit` to change the race preferences. You need to have at least Event organizer privileges in your club to be able to access these preferences.

You now have two options for exporting your data: as an [IOF XML 3.0](https://orienteering.sport/iof/it/data-standard-3-0/) file or as an [Excel 2003 XML format](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa140066(v=office.10)?redirectedfrom=MSDN) format file. The IOF XML format is widely supported by different kinds of orienteering software, while the Excel 2003 XML format has the advantage that you can easily manipulate the data by using the Microsoft Excel.

The procedures for downloading both kinds of file formats are given below.

### Download IOF XML 3.0 data from Eventor

- Go to `Data exchange`
- Choose the `xml` option by the text "Export entries"

### Download Eventor Excel 2003 XML

Go to entry overview. See the image below.
![Entry overview](img/ss1.jpg)

Download the entries as an Excel document.
![Excel download](img/ss2.jpg)

## Usage

The program can be run as a python script or as a standalone Windows program. The Windows executable file is easier to use, but it is quite large (approx. 225 MB), and also quite a bit slower than the python script.

The python script can be downloaded from the repository above. The Windows .exe file can be found under [releases](https://github.com/stalegjelsten/pyeventor2ttime/releases).

You can either run the python file `pyeventor2ttime.py` or the Windows executable file `pyeventor2ttime.exe`. The python file requires the packages `pandas`, `beautifulsoup4` and possibly `lxml` to run. Usage:

```bash
python pyeventor2ttime.py "Entry overview XXXXX.xls"
```

The Windows executable can be downloaded and should be put in the same folder as the .xls or .xml file you downloaded from Eventor. You need to open a command prompt in the same folder (tips: in Windows Explorer, hold the Shift key and right-click inside the folder to launch a command prompt from there). Use the command:

```bash
pyeventor2ttime.exe "entries_Random_event.xml"
```

The Windows executable program unpacks all the required modules to a temporary directory, and this will take some time, especially on computers with slower hard drives. It can take anywhere from 10 seconds to a few minutes. Please be patient. You can cancel the program by pressing Ctrl + C during execution.

## FAQ

Q: What about entries without emit ecards?  
A: Those entries are given no ecard number or number 999

Q: The script has a bug.  
A: Please open an issue here on GitHub

Q: I have a problem with the script.  
A: Send me a personal message here on GitHub or by email
