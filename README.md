# About
A script to convert Eventor (eventor.orientering.no) to ttime database format.

The [ttime](http://ttime.no/) application for Windows is used for timekeeping and entries during orienteering events. At the moment, there is no easy way to pull entries from the [Eventor](https://eventor.orientering.no) platform to ttime. This program is written to make that transition easier.

Previously, ttime could pull data from Eventor directly, but is unfortunately not possible anymore.

# Downloading entries from Eventor

Go to [eventor.orientering.no](https://eventor.orientering.no) and log in. Find your race and click edit to change the race preferences.

Go to entry overview. See the image below.
![Entry overview](img/ss1.jpg)

Download the entries as an Excel document.
![Excel download](img/ss2.jpg)

# Installation

The program can be run as a python script or as a standalone Windows program. The Windows executable file is easier to use, but it is quite large (approx. 225 MB), and also quite a bit slower than the python script.

The python script can be downloaded from the repository above. The Windows .exe file can be found under [releases](https://github.com/stalegjelsten/pyeventor2ttime/releases).

# Usage
You can either run the python file `pyeventor2ttime.py` or the Windows executable file `pyeventor2ttime.exe`. The python file requires the packages `pandas`, `beautifulsoup4` and possibly `lxml` to run. Usage:
```
python pyeventor2ttime.py "Entry overview XXXXX.xls"
```

The Windows executable can be downloaded and should be put in the same folder as your Entry overview XXXXX.xls file from Eventor. You need to open a command prompt in the same folder (tips: in Windows Explorer, hold the Shift key and right-click inside the folder to launch a command prompt from there). Use the command:

```
pyeventor2ttime.exe "Entry overview XXXXX.xls"
```
The Windows executable program unpacks all the required modules to a temporary directory, and this will take some time, especially on computers with slower hard drives. It can take anywhere from 10 seconds to a few minutes. Please be patient. You can cancel the program by pressing Ctrl + C during execution.
