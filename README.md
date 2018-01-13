# Wav-To-Excel
Script to transfer .wav files name and duration to a new Excel document using the Openpyxl module.

This script was made for a friend and colleague who needed to manually enter multiple .wav filenames and durations in an Excel document.

## Update
- I've added a Tkinter GUI version of the tool in the GUI folder. At this stage, the GUI version is working just fine but the code will need a bit of cleanup/refactor which I will be doing in the near future.

----

## Please Note
- Compared to other scripts on my profile, this one was made in Python 2 since it is what we use at work.

## Basic Usage
The user is prompted to enter a path where .wav files are located. If the path is valid ans contains .wav files, the script will loop through
the folder's content and for each file, will extract the filename, calculate it's duration using the number of frames and framerate and format
everything inside a new Excel workbook. Once the script is done, a file called "wav_to_excel.xlsx" is created at the same place where the python
script is.

