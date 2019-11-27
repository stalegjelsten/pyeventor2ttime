# About
A script to convert Eventor (eventor.orientering.no) to ttime database format

# Usage
You can either run the python file *pyeventor2ttime.py* or the Windows executable file *pyeventor2ttime.exe*. The python file requires the packages *pandas*, *beautifulsoup4* and possibly *lxml* to run. Usage:
```
python pyeventor2ttime.py "Entry overview XXXXX.xls"
```

The Windows executable can be downloaded and should be put in the same folder as your Entry overview XXXXX.xls file from Eventor. You need to open a command prompt in the same folder (tips: in Windows Explorer, hold the Shift key and right-click inside the folder to launch a command prompt from there). Use the command

```
pyeventor2ttime.exe "Entry overview XXXXX.xls"
```
It takes about 30 seconds to complete.
