# Wav-To-Excel
Script to transfer .wav files name and duration to a new Excel document using the Openpyxl module.

This script was made for a friend and colleague who needed to manually enter multiple .wav filenames and durations in an Excel document.

## Please Note
- Compared to other scripts on my profile, this one was made in Python 2 since it is what we use at work.
- I'm currently working on a GUI version of this with PyQt to make it easier for users to enter paths and visualize what is happening.

## Basic Usage
The user is prompted to enter a path where .wav files are located. If the path is valid ans contains .wav files, the script will loop through
the folder's content and for each file, will extract the filename, calculate it's duration using the number of frames and framerate and format
everything inside a new Excel workbook. Once the script is done, a file called "wav_to_excel.xlsx" is created at the same place where the python
script is.

