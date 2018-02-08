# -*- coding: utf-8 -*-
"""
Created on Mon Aug 07 15:23:12 2017

This script extracts the values from all sheets in a spreadsheet and
copies them in to another spreadsheet. No other information (formulas,
formatting, etc.) is copied over, unfortunately I do not know if that is
possible.

It was created to mine the hidden data from the EFT spredsheets.

@author: edward.barratt
"""

import os
import subprocess
from shutil import copyfile
import win32com.client as win32
import numpy as np

import EFT_Tools as tools

# The auto hot key paths, so that the initial warning window on the EFT can be
# closed automatically.
ahk_exepath = 'C:\Program Files\AutoHotkey\AutoHotkey.exe'
ahk_ahkpath = 'closeWarning.ahk'


InputFile = 'original\EFT2017_v8.0_original.xlsb'
OutputFile = 'output\EFT2017_v8.0_Extracted.xlsx'

# Create a temporary copy so that we don't risk changing the original.
infn, inex = os.path.splitext(InputFile)
InputFileTemp = infn + 'TEMP' + inex
copyfile(InputFile, InputFileTemp)
InputFileTemp = os.path.abspath(InputFileTemp) # Neccesary because win32 seems
OutputFile = os.path.abspath(OutputFile)       # to struggle with relative
                                               # paths.

# Open a excel application instance, and open both workbooks.
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
subprocess.Popen([ahk_exepath, ahk_ahkpath])
wb1 = excel.Workbooks.Open(InputFileTemp)
wb2 = excel.Workbooks.Add()
wb2.SaveAs(OutputFile)

shCount = wb1.Sheets.Count
print("Number of sheets: {}".format(shCount))

for shi in reversed(range(shCount)):
  sh = wb1.Sheets[shi+1]
  print('Copying the contents of sheet {}: {}'.format(shi+1, sh.Name))
  ws1 = wb1.Worksheets(sh.Name)
  ws2 = wb2.Worksheets.Add()
  ws2.Name = sh.Name
  # Check chunks that are 100 by 100 at a time, until we find the extent of the data.
  b = 100
  X1, X2 = 1, b
  Y1, Y2 = 1, b
  FoundEndX, FoundEndY = False, False
  while not FoundEndY:
    dataCheck = ws1.Range('{}{}:{}{}'.format(tools.numToLetter(X1), Y1, tools.numToLetter(X2), Y2)).Value
    dataCheck = np.array(dataCheck)
    if np.sum(np.equal(dataCheck, None)) == np.size(dataCheck):
      YMin = 1
      YMax = Y2
      FoundEndY = True
    elif Y2 > 10000:
      FoundEndY = True
    else:
      Y1 = Y2 + 1
      Y2 = Y2 + b
  while not FoundEndX:
    dataCheck = ws1.Range('{}{}:{}{}'.format(tools.numToLetter(X1), YMin, tools.numToLetter(X2), YMax)).Value
    dataCheck = np.array(dataCheck)
    if np.sum(np.equal(dataCheck, None)) == np.size(dataCheck):
      XMin = 1
      XMax = X2
      FoundEndX = True
    elif X2 > 10000:
      FoundEndX = True
    else:
      X1 = X2 + 1
      X2 = X2 + b
  # Now we've found the full extent, copy the values to the new workbook.
  RangeExtent = '{}{}:{}{}'.format(tools.numToLetter(XMin), YMin, tools.numToLetter(XMax), YMax)
  print('    Extent: {}'.format(RangeExtent))
  ws2.Range(RangeExtent).Value = ws1.Range(RangeExtent).Value

wb1.Save()
wb1.Close()
os.remove(InputFileTemp)
wb2.Save()
wb2.Close()
excel.Quit()
