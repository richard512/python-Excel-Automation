# Python-Excel-Automation

Simple examples of automation of excel via python

## Compatibility

| script | platform(s) it works on | what it does |
| ------------- | ------------- | ------------- |
| autoclicker.py | linux | checks a pixel on the screen. if it's a specified color, the script does a left mouse click |
| copyPaste.py | windows | appends one xlsx to another, then saves to csv |
| dispatch.win32com.py | windows | |
| findPhotoWithinPhoto.py | both | finds the location of an image within another image. useful for gui automation |
| openpyxlTest.py | both | uses openpyxl to generate a simple spreadsheet |
| runMacro.py | windows | uses win32com and excel com to run a VBA function inside of time.xlsm |
| time.py | both | outputs the current date and time |
| time.xlsm | windows | contains a button which executes a VBA function which runs time.py 3 times sequentially without clobbering |
| worksheet1.xlsx | both | simple spreadsheet, with functions in the "total" row |
| worksheet2.xlsx | both | simple spreadsheet, with functions in the "total" row |
| xlrdTest.py | both | uses xlrd to enumerate the contents of worksheet1.xlsx |

| Python Library Name  | Linux support | Windows support |
| ------------- | ------------- | ------------- |
| GTK  | strong  | weak  |
| cv2 (opencv)  | strong  | wonky install, but works  |
| pymouse  | strong  | weak  |
| win32com  | none  | strong  |
| numpy  | strong  | strong  |

## Install Prerequisites (Ubuntu Linux)

```
sudo apt install python pip python-opencv python-xlrd python-gtk2-dev
```

## Install Prerequisites (Windows 10)

Install [Python for Windows](https://www.python.org/downloads/windows/)

Add the location of python.exe to your PATH environment variable

Download [get-pip.py](https://bootstrap.pypa.io/get-pip.py) and run it in command prompt:

```
python get-pip.py
```

Add the location of pip.exe to your PATH environment variable

Run this in a command prompt

```
pip install pypiwin32
pip install xlrd
pip install openpyxl
```

Download the latest OpenCV whl (wheel) file from [here](http://www.lfd.uci.edu/~gohlke/pythonlibs/#opencv) then install it like this:

```
pip install opencv_python-3.2.0-cp36-cp36m-win32.whl
```

Might need [Visual C++ 2015 redistributable (vc_redist.x86.exe)](https://www.microsoft.com/en-us/download/details.aspx?id=48145) for OpenCV

If you need GTK, follow [this guide](https://www.gtk.org/download/windows.php)
