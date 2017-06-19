# Python-Excel-Automation

Simple examples of automation of excel via python

## Compatibility

Most of these scripts were built to run on linux. Some were built to run on windows. Some work on both.

| Library  | Linux support | Windows support |
| ------------- | ------------- | ------------- |
| GTK  | strong  | weak  |
| cv2 (opencv)  | strong  | weak  |
| pymouse  | strong  | weak  |
| win32com  | none  | strong  |


## Install Prerequisites (Ubuntu Linux)

```
sudo apt install python pip python-opencv python-xlrd python-gtk2-dev
```

## Install Prerequisites (Windows 10)

Install [Python for Windows](https://www.python.org/downloads/windows/)

Add the location of python.exe to your PATH environment variable

Download [Pip](https://bootstrap.pypa.io/get-pip.py)

Go to the location of get-pip.py in command prompt and run:

```
python get-pip.py
```

Add the location of pip.exe to your PATH environment variable

Run this in a command prompt

```
pip install pypiwin32
```

If you need GTK, follow [this guide](https://www.gtk.org/download/windows.php)
