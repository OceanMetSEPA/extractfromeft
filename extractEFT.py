 # -*- coding: utf-8 -*-
"""
Created on Wed Oct 18 12:25:39 2017

@author: edward.barratt
"""

import os
from os import path
import time
import argparse
import shutil
import win32com.client as win32

#import processEFT
import EFT_Tools as tools


availableBreakdowns = {'Basic Split', 'Detailed Option 1',
                    'Detailed Option 2', 'Detailed Option 3'}

def extractEFT(fileName, location, roadtype=tools.availableRoadTypes,
               roadTypes=tools.availableRoadTypes, keepTempFiles=False,
               vBreakdown='Detailed Option 2'):
  tic = time.clock()

  # Get the files are ready for processing.
  ahk_ahkpathG, fileNames, versions, versionsForOutput = tools.prepareToExtract(fileName, location)
  fileName = fileNames[0]
  version = versions[0]
  versionForOutPut = versionsForOutput[0]

  details = tools.versionDetails[version]

  # Make a temporary copy of the filename, so that we do no processing on the
  # original. Just in case we brake it. Also define temporary file names and
  # output save locations, etc.
  [FN, FE] =  path.splitext(fileName)
  fileNameT = FN + '_TEMP_' + tools.randomString() + FE
  fileNameCSV_ = fileName.replace(FE, '.csv')
  fileNameCSVNotComplete = fileNameCSV_.replace('empty', 'extracted')
  fileNameCSV = fileNameCSV_.replace('empty', 'extracted')

  vi = 1
  while path.isfile(fileNameCSV):
    vi += 1
    fileNameCSV = fileNameCSV_.replace('empty', 'extracted({})'.format(vi))
  shutil.copyfile(fileName, fileNameT)

  # Create the Excel Application object.
  excel = win32.gencache.EnsureDispatch('Excel.Application')

  # And now start the processing!
  first = True
  tempFilesCreated = [fileNameT]
  NOxNO2F = tools.readNO2Factors(mode='ByRoadType')
  inputData = tools.createEFTInput(vBreakdown=vBreakdown, roadTypes=roadTypes)
  inputData = inputData.as_matrix()
  for year in years:
    ticyear = time.clock()
    print('Year: {}'.format(year))

    # Run the ET and get the output.
    excel, newSavedFile, defaultProportions, k = tools.runAndExtract(fileNameT,
                          location, year, -9, tools.ahk_exepath, ahk_ahkpathG,
                          vBreakdown, details, versionForOutPut, excel=excel, inputData=inputData)
    tempFilesCreated.append(newSavedFile)
    # Now get the output values as a dataframe.
    print('  Done, reading output values.')
    output = tools.extractOutput(newSavedFile, versionForOutPut, year, location, -9, details)
    output = output.sort_values(['year', 'area', 'type', 'euro', 'speed', 'vehicle'])
    output = tools.addNO2(output, Factors=NOxNO2F, mode='ByRoadType')
    output = output.drop('euro', 1)
    print('  Writing to file')
    if first:
      # Save to a new csv file.
      output.to_csv(fileNameCSVNotComplete, index=False)
      first = False
    else:
      # Append to the csv file.
      output.to_csv(fileNameCSVNotComplete, mode='a', header=False, index=False)
    tocyear = time.clock()
    print('  Processing for year {} complete in {}.'.format(year, tools.secondsToString(tocyear-ticyear, form='long')))
  shutil.move(fileNameCSVNotComplete, fileNameCSV)
  print('Processing complete. Output saved in the following files.')
  print('  {}'.format(fileNameCSV))
  if not keepTempFiles:
    print('Deleting temporary files.')
    for tf in tempFilesCreated:
      os.remove(tf)
  toc = time.clock()
  excel.Quit()
  print('Process complete in {}.'.format(tools.secondsToString(toc-tic, form='long')))


if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Extract emission values from the EFT')
  parser.add_argument('--version', '-v', metavar='version number',
                      type=float, nargs='?', default=7.0,
                      choices=tools.availableVersions,
                      help="The EFT version number. One of {}. Default 7.0.".format(", ".join(str(v) for v in tools.availableVersions)))
  parser.add_argument('-a', '--area', metavar='areas',
                      type=str, nargs='?', default='Scotland',
                      choices=tools.availableAreas,
                      help="The areas to be processed. One of '{}'. Default 'Scotland'.".format("', '".join(tools.availableAreas)))
  rtop = list(tools.availableRoadTypes)
  rtop.append('all')
  parser.add_argument('-t', '--roadtype', metavar='road type',
                      type=str, nargs='?', default='all',
                      choices=rtop,
                      help="The road type to be processed. One of '{}', or 'all'. Default 'all'.".format("', '".join(tools.availableRoadTypes)))
  parser.add_argument('-y', '--years', metavar='year',
                      type=int, nargs='*', default=-9999,
                      choices=range(2008, 2031),
                      help="The year or years to be processed. Default 'all'")
  parser.add_argument('-b', '--breakdown', metavar='vehicle breakdown',
                      type=str, nargs='*', default='Detailed Option 2',
                      choices=availableBreakdowns,
                      help="Vehicle breakdown. One of '{}'. Default 'Detailed Option 2'.".format("', '".join(availableBreakdowns)))
  parser.add_argument('--keeptemp', metavar='keeptemp',
                      type=bool,  nargs='?', default=False,
                      help="Whether to keep or delete temporary files. Boolean. Default False (delete).")
  parser.add_argument('-i', '--inputfile', metavar='input file',
                      type=str,   nargs='?', default=None,
                      help="The file to process. If set then version will be ignored.")
  args = parser.parse_args()
  breakdown = args.breakdown
  version = args.version
  inputfile = args.inputfile
  if inputfile is not None:
    version = tools.extractVersion(inputfile)
  else:
    inputfile = tools.getInputFile(version)
  if version == 6.0:
    availableYears = range(2008, 2031)
  else:
    availableYears = range(2013, 2031)
  area = args.area
  years = args.years
  roadtype = args.roadtype
  if roadtype in ['all', 'All', 'ALL']:
    roadtype = tools.availableRoadTypes
  if years == -9999:
    years = availableYears
  keepTempFiles = args.keeptemp

  if not all(y in availableYears for y in years):
    raise ValueError('One or more years are not allowed for the specified EFT version.')

  extractEFT(inputfile, area, roadTypes=roadtype, keepTempFiles=False, vBreakdown=breakdown)

