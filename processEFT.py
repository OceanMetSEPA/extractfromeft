# -*- coding: utf-8 -*-
"""
Created on Wed Jun 07 10:37:25 2017

@author: edward.barratt

Routines to process Emission Factor Toolbox spreadsheets and extract emission
rates for vehicle classes against year and euro class.
"""
from __future__ import print_function

import os
from os import path
import argparse
import subprocess
import time
from datetime import datetime
import shutil
import random
import string
import numpy as np
import pandas as pd
import win32com.client as win32
import pywintypes

from tools import secondsToString, extractVersion

# Define some global variables. These may need to be augmented if a new EFT
# version is released.
workingDir = os.getcwd()

ahk_exepath = 'C:\Program Files\AutoHotkey\AutoHotkey.exe'
ahk_ahkpath = 'closeWarning.ahk'

versionDetails = {}
versionDetails[7.4] = {}
versionDetails[7.4]['vehRowStarts'] = [69, 79, 91, 101, 114, 130, 146, 161]
versionDetails[7.4]['vehRowEnds']   = [76, 87, 98, 109, 125, 141, 157, 172]
versionDetails[7.4]['vehRowStartsMC'] = [177, 183, 189, 195, 201, 207]
versionDetails[7.4]['vehRowEndsMC']   = [182, 188, 194, 200, 206, 212]
versionDetails[7.4]['busCoachRow']   = [429, 430]
versionDetails[7.4]['SourceNameName'] = 'Source Name'
versionDetails[7.4]['AllLDVName'] = 'All LDVs (g/km/s)'
versionDetails[7.4]['AllHDVName'] = 'All HDVs (g/km/s)'
versionDetails[7.4]['AllVehName'] = 'All Vehicles (g/km/s)'
versionDetails[7.4]['PolName'] = 'Pollutant Name'
versionDetails[7.0] = {}
versionDetails[7.0]['vehRowStarts'] = [69, 79, 100, 110, 123, 139, 155, 170]
versionDetails[7.0]['vehRowEnds']   = [75, 87, 106, 119, 134, 150, 166, 181]
versionDetails[7.0]['vehRowStartsMC'] = [186, 192, 198, 204, 210, 216]
versionDetails[7.0]['vehRowEndsMC']   = [191, 197, 203, 209, 215, 221]
versionDetails[7.0]['busCoachRow']   = [482, 483]
versionDetails[7.0]['SourceNameName'] = 'Source Name'
versionDetails[7.0]['AllLDVName'] = 'All LDVs (g/km/s)'
versionDetails[7.0]['AllHDVName'] = 'All HDVs (g/km/s)'
versionDetails[7.0]['AllVehName'] = 'All Vehicles (g/km/s)'
versionDetails[7.0]['PolName'] = 'Pollutant Name'
versionDetails[6.0] = {}
versionDetails[6.0]['vehRowStarts'] = [69, 79, 100, 110, 123, 139, 155, 170]
versionDetails[6.0]['vehRowEnds']   = [75, 87, 106, 119, 134, 150, 166, 181]
versionDetails[6.0]['vehRowStartsMC'] = [186, 192, 198, 204, 210, 216]
versionDetails[6.0]['vehRowEndsMC']   = [191, 197, 203, 209, 215, 221]
versionDetails[6.0]['busCoachRow']   = [482, 483]
versionDetails[6.0]['SourceNameName'] = 'Source_Name'
versionDetails[6.0]['AllLDVName'] = 'All LDV (g/km/s)'
versionDetails[6.0]['AllHDVName'] = 'All HDV (g/km/s)'
versionDetails[6.0]['AllVehName'] = 'All Vehicle (g/km/s)'
versionDetails[6.0]['PolName'] = 'Pollutant_Name'

euroClassNameVariations = dict()
euroClassNameVariations[0] = ['1Pre-Euro 1', '1Pre-Euro I', '1_Pre-Euro 1', '2Pre-Euro 1',
          '4Pre-Euro 1', '5Pre-Euro 1', '6Pre-Euro 1', '7Pre-Euro 1',
          '1_Pre-Euro 1']
euroClassNameVariations[1] = ['2Euro 1', '2Euro I', '1Euro 1', '2Euro 1', '2Euro 1',
          '4Euro 1', '5Euro 1', '6Euro 1', '7Euro 1', '9 Euro I DPFRF',
          '8Euro 1 DPFRF', '9Euro I DPFRF']
euroClassNameVariations[2] = ['3Euro 2', '3Euro II', '1Euro 2', '2Euro 2', '2Euro 2',
          '4Euro 2', '5Euro 2', '6Euro 2', '7Euro 2', '10 Euro II DPFRF',
          '9Euro II SCRRF', '9Euro 2 DPFRF']
euroClassNameVariations[3] = ['4Euro 3', '4Euro III', '1Euro 3', '2Euro 3', '2Euro 3',
          '4Euro 3', '5Euro 3', '6Euro 3', '7Euro 3', '11 Euro III DPFRF',
          '10Euro III SCRRF', '8Euro 3 DPF', '10Euro 3 DPFRF']
euroClassNameVariations[4] = ['5Euro 4', '5Euro IV', '1Euro 4', '2Euro 4', '2Euro 4',
          '4Euro 4', '5Euro 4', '6Euro 4', '7Euro 4', '12 Euro IV DPFRF',
          '11Euro IV SCRRF', '9Euro 4 DPF']
euroClassNameVariations[5] = ['6Euro 5', '6Euro V', '1Euro 5', '2Euro 5', '2Euro 5',
          '4Euro 5', '5Euro 5', '6Euro 5', '7Euro 5', '7Euro V_SCR',
          '6Euro V_EGR', '12Euro V EGR + SCRRF']
euroClassNameVariations[6] = ['7Euro 6', '6Euro VI', '1Euro 6', '2Euro 6', '2Euro 6',
          '4Euro 6', '5Euro 6', '6Euro 6', '7Euro 6', '8Euro VI',
          '7Euro 6c', '7Euro 6d']

vehSplit2 = "Detailed Option 2"
vehSplit3 = "Detailed Option 3"

euroClassNameVariationsAll = euroClassNameVariations[0][:]
for ei in range(1,7):
  euroClassNameVariationsAll.extend(euroClassNameVariations[ei])
euroClassNameVariationsAll = list(set(euroClassNameVariationsAll))




EuroClassNameColumns = ["A", "H"]
DefaultEuroColumns = ["B", "I"]
UserDefinedEuroColumns = ["D", "K"]
EuroClassNameColumnsMC = ["B", "H"]
DefaultEuroColumnsMC = ["C", "I"]
UserDefinedBusColumn = ["D"]
UserDefinedBusMWColumn = ["E"]
DefaultBusColumn = ["B"]
DefaultBusMWColumn = ["C"]

availableVersions = versionDetails.keys()
availableAreas = ['England (not London)', 'Northern Ireland', 'Scotland', 'Wales']
availableModes = ['ExtractAll', 'ExtractCarRatio', 'ExtractBus']
availableEuros = [0,1,2,3,4,5,6]

def randomString(N = 10):
  return ''.join(random.choice(string.ascii_uppercase + string.ascii_lowercase + string.digits) for x in range(N))

def romanNumeral(N):
  # Could write a function that deals with any, but I only need up to 10.
  RNs = [0, 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
  return RNs[N]

def euroSearchTerms(N):
  ES = euroClassNameVariations[N]
  return ES

def checkEuroClassesValid(workBook, vehRowStarts, vehRowEnds, EuroClassNameColumns, MC=0):
  """
  Check that all of the available euro classes are specified.
  """
  if MC == 1:
    print("      Checking all motorcycle euro class names are understood.")
  elif MC == -1:
    print("      Checking all non-motorcycle euro class names are understood.")
  else:
    print("      Checking all euro class names are understood.")
  ws_euro = workBook.Worksheets("UserEuro")
  for [vi, vehRowStart] in enumerate(vehRowStarts):
    vehRowEnd = vehRowEnds[vi]
    for [ci, euroNameCol] in enumerate(EuroClassNameColumns):
      euroClassRange = "{col}{rstart}:{col}{rend}".format(col=euroNameCol, rstart=vehRowStart, rend=vehRowEnd)
      euroClassesAvailable = ws_euro.Range(euroClassRange).Value

      for ecn in euroClassesAvailable:
        ecn = ecn[0]
        if ecn is None:
          continue
        if ecn not in euroClassNameVariationsAll:
          raise ValueError('Unrecognized Euro Class Name: "{}".'.format(ecn))
  print("        All understood.")


def specifyBusCoach(wb, busCoach, busCoachRow, UserDefinedBusColumn,
                    UserDefinedBusMWColumn, DefaultBusColumn, DefaultBusMWColumn):
  defaultBusProps = {}
  ws_euro = wb.Worksheets("UserEuro")
  defaultBusProps['bus_non_mw'] = ws_euro.Range("{}{}".format(DefaultBusColumn[0], busCoachRow[0])).Value
  defaultBusProps['coach_non_mw'] = ws_euro.Range("{}{}".format(DefaultBusColumn[0], busCoachRow[1])).Value
  defaultBusProps['bus_mw'] = ws_euro.Range("{}{}".format(DefaultBusMWColumn[0], busCoachRow[0])).Value
  defaultBusProps['coach_mw'] = ws_euro.Range("{}{}".format(DefaultBusMWColumn[0], busCoachRow[1])).Value

  if busCoach != 'default':
    if busCoach == 'bus':
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[0])).Value = 1
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[1])).Value = 0
      try:
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[0])).Value = 1
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[1])).Value = 0
      except pywintypes.com_error:
        # Doesn't work in version 6, it doesn't let you specify the motorway proportion.
        pass
    elif busCoach == 'coach':
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[0])).Value = 0
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[1])).Value = 1
      try:
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[0])).Value = 0
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[1])).Value = 1
      except pywintypes.com_error:
        pass
    else:
      raise ValueError("busCoach should be either 'bus' or 'coach'. '{}' is not understood.".format(busCoach))
  return defaultBusProps

def specifyEuroProportions(euroClass, workBook, vehRowStarts, vehRowEnds,
                 EuroClassNameColumns, DefaultEuroColumns, UserDefinedEuroColumns, MC=False):
  """
  Specify the euro class proportions.
  Will return the defualt proportions.
  """
  defaultProps = {}
  #print("    Setting euro ratios to 100% for euro {}.".format(euroClass))
  ws_euro = workBook.Worksheets("UserEuro")
  for [vi, vehRowStart] in enumerate(vehRowStarts):
    if MC:
      vehNameA = ws_euro.Range("A{row}".format(row=vehRowStart)).Value
      vehNameB = ws_euro.Range("A{row}".format(row=vehRowStart+1)).Value
      if vehNameB is None:
        vehName = 'Motorcycle - {}'.format(vehNameA)
      else:
        vehName = 'Motorcycle - {} - {}'.format(vehNameA, vehNameB)
    else:
      vehName = ws_euro.Range("A{row}".format(row=vehRowStart-1)).Value
    #print("      Setting euro ratios for {}.".format(vehName))
    vehRowEnd = vehRowEnds[vi]
    for [ci, euroNameCol] in enumerate(EuroClassNameColumns):
      userDefinedCol = UserDefinedEuroColumns[ci]
      defaultEuroCol = DefaultEuroColumns[ci]
      euroClassRange = "{col}{rstart}:{col}{rend}".format(col=euroNameCol, rstart=vehRowStart, rend=vehRowEnd)
      euroClassesAvailable = ws_euro.Range(euroClassRange).Value
      # Make sure we don't include trailing 'None' rows.
      euroClassesAvailableR = list(reversed(euroClassesAvailable))
      for eca in euroClassesAvailableR:
        #print(eca)
        if eca[0] is None:
          vehRowEnd = vehRowEnd - 1
        else:
          break
      # See which columns contain a line that specifies the required euro class.
      rowsToDo = []
      euroClass_ = euroClass
      while len(rowsToDo) == 0:
        got = False
        euroSearchTerms_ = euroSearchTerms(euroClass_)
        for [ei, name] in enumerate(euroClassesAvailable):
          name = name[0]
          if name in euroSearchTerms_:
            rowsToDo.append(vehRowStart + ei)
            got = True
        if not got:
          # print('      No values available for euro {}, trying euro {}.'.format(euroClass_, euroClass_-1))
          euroClass_ -= 1
      ignoreForPropRecord = False
      if euroClass_ != euroClass:
        ignoreForPropRecord = True
      # Get the default proportions.
      defaultProportions = []
      for row in rowsToDo:
        propRange = "{col}{row}".format(col=defaultEuroCol, row=row)
        defaultProportion = ws_euro.Range(propRange).Value
        defaultProportions.append(defaultProportion)
        #print(propRange)
        #print(defaultProportions)
      defaultProportions = np.array(defaultProportions)
      if ci == 0:
        if ignoreForPropRecord:
          #print('        Default proportions taken as 0.00%.')
          defaultProps[vehName] = 0
        else:
          #print('        Default proportions are {:.2f}%.'.format(100*sum(defaultProportions)))
          defaultProps[vehName] = 100*sum(defaultProportions)
      # Normalize them.
      if sum(defaultProportions) < 0.00001:
        defaultProportions = defaultProportions + 1
      userProportions = defaultProportions/sum(defaultProportions)
      # And set the values in the sheet.
      # Set all to zero first.
      userRange = "{col}{rstart}:{col}{rend}".format(col=userDefinedCol, rstart=vehRowStart, rend=vehRowEnd)
      ws_euro.Range(userRange).Value = 0
      # Then set the specific values.
      #print(rowsToDo)
      for [ri, row] in enumerate(rowsToDo):
        userRange = "{col}{row}".format(col=userDefinedCol, row=row)
        value = userProportions[ri]
        ws_euro.Range(userRange).Value = value
  #print('    All complete')
  return defaultProps

def splitSourceNameS(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  row['vehicle'] = v
  return int(s[1:])

def splitSourceNameV(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  return v

def splitSourceNameT(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  return t

def getInputFile(mode, version, directory='input'):
  """
  Return the absolute path to the appropriate file for the selected mode and
  version. Will return an error if no file is available.
  """

  # First check that the directory exists.
  if not path.isdir(directory):
    raise ValueError('Cannot find directory {}.'.format(directory))

  # Now figure out the file name.
  if version == 6.0:
    vPart = 'EFT2014_v6.0.2'
    ext = '.xls'
  elif version == 7.0:
    vPart = 'EFT2016_v7.0'
    ext = '.xlsb'
  elif version == 7.4:
    vPart = 'EFT2017_v7.4'
    ext = '.xlsb'
  else:
    raise ValueError('Version {} is not recognised.'.format(version))

  if mode in ['ExtractAll', 'ExtractBus']:
    fname = ['{}/{}_prefilledValues{}'.format(directory, vPart, ext)]
  elif mode == 'ExtractCarRatio':
    fname = ['{}/{}_prefilled_CarsDetailed2{}'.format(directory, vPart, ext),
             '{}/{}_prefilled_CarsDetailed3{}'.format(directory, vPart, ext)]
  else:
    raise ValueError('Mode {} is not recognised.'.format(mode))

  # Check that file(s) exists.
  for f in fname:
    if not path.exists(f):
      raise ValueError('Cannot find file {}.'.format(f))

  # return the absolute paths.
  if mode in ['ExtractAll', 'ExtractBus']:
    return path.abspath(fname[0])
  else:
    return [path.abspath(f) for f in fname]


def prepareToExtract(fileNames, locations):
  """
  Extract the pre-processing information from the filenames and locations.
  """
  # Make sure location is a list that can be iterated through.
  if type(locations) is str:
    locations = [locations]
  # Make sure fileNames is a list that can be iterated through.
  if type(fileNames) is str:
    fileNames = [fileNames]

  # Check that the auto hot key executable, and control file, are available.
  if not path.isfile(ahk_exepath):
    raise ValueError('The Autohotkey executable file {} could not be found.'.format(ahk_exepath))
  if not path.isfile(ahk_ahkpath):
    ahk_ahkpath_ = workingDir + '\\' + ahk_ahkpath
    if not path.isfile(ahk_ahkpath_):
      raise ValueError('The Autohotkey file {} could not be found.'.format(ahk_ahkpath))
    else:
      ahk_ahkpathGot = ahk_ahkpath_
  else:
    ahk_ahkpathGot = ahk_ahkpath

  versionNos = []
  versionForOutputs = []
  for fNi, fN in enumerate(fileNames):
    # Extract the version number.
    version, versionForOutput = extractVersion(fN)
    versionNos.append(version)
    versionForOutputs.append(versionForOutput)

    # Get the absolute path to the file. The excel win32 stuff doesn't seem to
    # work with relative paths.
    fN_ = path.abspath(fN)
    if not path.isfile(fN):
      raise ValueError('Could not find {}.'.format(fN))
    fileNames[fNi] = fN_

  return ahk_ahkpathGot, fileNames, versionNos, versionForOutputs

def runAndExtract(excel, fileName, location, year, euroClass, ahk_exepath,
                  ahk_ahkpathG, vehSplit, details, versionForOutPut,
                  checkEuroClasses=False, DoMCycles=True, DoBusCoach=False,
                  busCoach='default'):
  """
  Prepare the file for running the macro.
  """
  # Start off the autohotkey script as a (parallel) subprocess. This will
  # continually check until the compatibility warning appears, and then
  # close the warning.
  subprocess.Popen([ahk_exepath, ahk_ahkpathG])

  # Open the document.
  wb = excel.Workbooks.Open(fileName)
  excel.Visible = True

  if checkEuroClasses:
    # Check that all of the euro class names within the document are as
    # we would expect. An error will be raised if there are any surprises
    # and this will mean that the global variables at the start of the
    # code will need to be edited.
    if DoMCycles:
      checkEuroClassesValid(wb, details['vehRowStartsMC'], details['vehRowEndsMC'], EuroClassNameColumnsMC, MC=1)
    checkEuroClassesValid(wb, details['vehRowStarts'], details['vehRowEnds'], EuroClassNameColumns, MC=-1)

  # Set the default values in the Input Data sheet.
  ws_input = wb.Worksheets("Input Data")
  ws_input.Range("B4").Value = location
  ws_input.Range("B5").Value = year
  # Ensure that the correct detailed split is specified. Setting it will
  # raise a popup and delete the traffic array, so we want to avoid that.
  if ws_input.Range("B6").Value != vehSplit:
    raise ValueError('Traffic Format should be "{}".'.format(vehSplit))

  # Now we need to populate the UserEuro table with the defaults. Probably
  # only need to do this once per year, per area, but will do it every time
  # just in case.
  excel.Application.Run("PasteDefaultEuroProportions")

  # Now specify that we only want the specified euro class, by turning the
  # proportions for that class to 1, (or a weighted value if there are more
  # than one row for the particular euro class). This function also reads
  # the default proportions.
  defaultProportions = pd.DataFrame(columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
  # Motorcycles first
  if DoMCycles:
    print('      Assigning fleet euro proportions for motorcycles.')
    defaultProportionsMC_ = specifyEuroProportions(euroClass, wb,
                                details['vehRowStartsMC'], details['vehRowEndsMC'],
                                EuroClassNameColumnsMC, DefaultEuroColumnsMC,
                                UserDefinedEuroColumns, MC=True)
    for key, value in defaultProportionsMC_.items():
      defaultProportionsRow = pd.DataFrame([[year, location, key, euroClass, value]],
                                           columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
      defaultProportions = defaultProportions.append(defaultProportionsRow)
    print('      Assigning fleet euro proportions for all other vehicle types.')
  else:
    print('      Assigning fleet euro proportions for all vehicle types except motorcycles.')
  # And all other vehicles
  defaultProportions_ = specifyEuroProportions(euroClass, wb,
                           details['vehRowStarts'], details['vehRowEnds'],
                           EuroClassNameColumns, DefaultEuroColumns,
                           UserDefinedEuroColumns)
  # Organise the default proportions.
  for key, value in defaultProportions_.items():
    defaultProportionsRow = pd.DataFrame([[year, location, key, euroClass, value]],
                                         columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
    defaultProportions = defaultProportions.append(defaultProportionsRow)
  defaultProportions['version'] = versionForOutPut

  busCoachProportions = 'NotMined'
  if DoBusCoach:
    # Set the bus - coach proportions.
    busCoachProportions = specifyBusCoach(wb, busCoach, details['busCoachRow'],
                                          UserDefinedBusColumn, UserDefinedBusMWColumn,
                                          DefaultBusColumn, DefaultBusMWColumn)

  # Now run the EFT tool.
  ws_input.Select() # Select the appropriate sheet, we can't run the macro
                    # from another sheet.
  print('      Running EFT routine. Ctrl+C will pause processing at the end of the routine...')
  alreadySaved = False
  try:
    excel.Application.Run("RunEfTRoutine")
    print('        Complete. Ctrl+C will now halt entire programme as usual.')
    time.sleep(0.5)
  except KeyboardInterrupt:
    print('Process paused at {}.'.format(datetime.strftime(datetime.now(), '%H:%M:%S on %d-%m-%Y')))
    # Save and Close. Saving as an xlsm, rather than a xlsb, file, so that it
    # can be opened by pandas.
    (FN, FE) =  path.splitext(fileName)
    if DoBusCoach:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}_{}).xlsm'.format(location, year, euroClass, busCoach))
    else:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}).xlsm'.format(location, year, euroClass))
    wb.SaveAs(tempSaveName, win32.constants.xlOpenXMLWorkbookMacroEnabled)
    wb.Close()
    excel.Quit()
    alreadySaved = True
    time.sleep(1)
    raw_input('Press enter to resume.')
    print('Resumed at {}.'.format(datetime.strftime(datetime.now(), '%H:%M:%S on %d-%m-%Y')))
    excel = win32.gencache.EnsureDispatch('Excel.Application')

  if not alreadySaved:
    # Save and Close. Saving as an xlsm, rather than a xlsb, file, so that it
    # can be opened by pandas.
    (FN, FE) =  path.splitext(fileName)
    if DoBusCoach:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}_{}).xlsm'.format(location, year, euroClass, busCoach))
    else:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}).xlsm'.format(location, year, euroClass))
    wb.SaveAs(tempSaveName, win32.constants.xlOpenXMLWorkbookMacroEnabled)
    wb.Close()

  time.sleep(1) # To allow all systems to catch up.
  return excel, tempSaveName, defaultProportions, busCoachProportions

def extractOutput(fileName, versionForOutPut, year, location, euroClass, details):
  ex = pd.ExcelFile(fileName)
  output = ex.parse("Output")
  # Add some other columns to the dataframe.
  output['version'] = versionForOutPut
  output['year'] = year
  output['area'] = location
  output['type'] = output.apply(splitSourceNameT, SourceName=details['SourceNameName'], axis=1)
  output['vehicle'] = output.apply(splitSourceNameV, SourceName=details['SourceNameName'], axis=1)
  output['euro'] = euroClass
  output['speed'] = output.apply(splitSourceNameS, SourceName=details['SourceNameName'], axis=1)
  # Drop columns that are not required anymore.
  output = output.drop(details['SourceNameName'], 1)
  output = output.drop(details['AllLDVName'], 1)
  output = output.drop(details['AllHDVName'], 1)
  # Pivot the table so each pollutant has a column.
  pollutants = list(output[details['PolName']].unique())
  # Rename, because after the pivot the 'column' name will become the
  # index name.
  output = output.rename(columns={details['PolName']: 'RowIndex'})
  output = output.pivot_table(index=['year', 'area', 'euro', 'version',
                                     'speed', 'vehicle', 'type'],
                                    columns='RowIndex',
                                    values=details['AllVehName'])
  output = output.reset_index()

  renames = {}
  # Rename the pollutant columns to include the units.
  for Pol in pollutants:
    if Pol == 'PM25':
      Pol_ = 'PM2.5'
    else:
      Pol_ = Pol
    renames[Pol] = '{} (g/km/s/veh)'.format(Pol_)
  output = output.rename(columns=renames)
  return output

def extractPetrolDieselCarProportions(fileName2, fileName3, locations, keepTempFiles=False):
  tic = time.clock()

  # get the files are ready for processing.
  ahk_ahkpathG, fileNames, versions, versionsForOutput = prepareToExtract([fileName2, fileName3], locations)
  fileName2 = fileNames[0]
  fileName3 = fileNames[1]
  version2 = versions[0]
  version3 = versions[1]
  if version2 != version3:
    raise ValueError('Input files should be of the same version.')
  version = version2
  versionForOutPut = versionsForOutput[0]

  # Now get the version dependent properties, mainly to do with which rows of
  # the spreadsheet contain which data.
  details = versionDetails[version]

  # Make a temporary copy of the filename, so that we do no processing on the
  # original. Just in case we brake it. Also define temporary file names and
  # output save locations, etc.
  [FN2, FE2] =  path.splitext(fileName2)
  [FN3, FE3] =  path.splitext(fileName3)
  fileName2T = FN2 + '_TEMP_' + randomString() + FE2
  fileName3T = FN3 + '_TEMP_' + randomString() + FE3
  shutil.copyfile(fileName2, fileName2T)
  shutil.copyfile(fileName3, fileName3T)
  fileNameCSVNotComplete = fileName2.replace('prefilled_CarsDetailed2', 'CarFuelRatios_InPreparation')
  fileNameCSVNotComplete = fileNameCSVNotComplete.replace(FE2, '.csv')
  fileNameCSV = fileNameCSVNotComplete.replace('_InPreparation', '')

  # Create the Excel Application object.
  excel = win32.gencache.EnsureDispatch('Excel.Application')

  # And now start the processing!
  first = True
  tempFilesCreated = [fileName2T, fileName3T]
  #defaultProportions = pd.DataFrame(columns=['year', 'area', 'vehicle', 'euroClass', 'proportion'])

  for location in locations:
    ticloc = time.clock()
    print('Location: {}'.format(location))
    for year in years:
      ticyear = time.clock()
      print('  Year: {}'.format(year))
      for euroClass in euroClasses:
        ticeuro = time.clock()
        carRatios = pd.DataFrame(columns=['year', 'area', 'euro', 'roadType', 'petrol', 'diesel', 'maximumFitResidual'])
        print('    Euro class: {}'.format(euroClass))
        if first:
          excel, newSavedFile2, defaultProportions2, k = runAndExtract(excel, fileName2T, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit2, details, versionForOutPut, checkEuroClasses=True, DoMCycles=False)
          excel, newSavedFile3, defaultProportions3, k = runAndExtract(excel, fileName3T, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut, checkEuroClasses=True, DoMCycles=False)
        else:
          excel, newSavedFile2, defaultProportions2, k = runAndExtract(excel, fileName2T, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit2, details, versionForOutPut, DoMCycles=False)
          excel, newSavedFile3, defaultProportions3, k = runAndExtract(excel, fileName3T, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut, DoMCycles=False)
        tempFilesCreated.extend([newSavedFile2, newSavedFile3])
        print('      Done, reading output values.')

        # Now get the output values as a dataframe.
        output2 = extractOutput(newSavedFile2, versionForOutPut, year, location, euroClass, details)
        output3 = extractOutput(newSavedFile3, versionForOutPut, year, location, euroClass, details)

        # Assume that the diesel/car ratio could depend on road type but not on
        # speed.
        roadTypes2 = set(output2['type'])
        roadTypes3 = set(output3['type'])
        if roadTypes2 != roadTypes3:
          raise ValueError('road types do not agree.')
        roadTypes = roadTypes2
        for roadType in roadTypes:
          output2_r = output2[output2['type'] == roadType]
          output3_r = output3[output3['type'] == roadType]
          output_allCars = output2_r[output2_r['vehicle'] == '2. Cars']
          output_petrolCars = output3_r[output3_r['vehicle'] == '2. Petrol Cars']
          output_dieselCars = output3_r[output3_r['vehicle'] == '3. Diesel Cars']
          Pols = ['NOx (g/km/s/veh)', 'PM10 (g/km/s/veh)', 'PM2.5 (g/km/s/veh)']
          EAll = []
          EPetrol = []
          EDiesel = []
          for Pol in Pols:
            EAll.extend(list(output_allCars[Pol]))
            EPetrol.extend(list(output_petrolCars[Pol]))
            EDiesel.extend(list(output_dieselCars[Pol]))
          # Now we have the emissions for all cars (EAll), for only diesel cars
          # (EDiesel), and for only petrol cars (EPetrol).
          # We assume that EAll = A*EPetrol + B*EDiesel and solve for A and B.
          # That is equivalent to EAll = C*P, where C = [[EPetrol EDiesel]] and
          # P = [[A], [B]]. We use numpy.linalg.lstsq to solve for P.
          EPetrol = np.array(EPetrol)
          EDiesel = np.array(EDiesel)
          EAll = np.array(EAll)
          C = np.vstack([EPetrol, EDiesel]).T
          solution, sumResSquared, _notneeded_, _notneeded2_ = np.linalg.lstsq(C, EAll)
          A, B = solution
          # Get the residuals.
          res = abs(EAll - np.dot(C, solution))
          maxres = np.max(res)

          ratioSingle = pd.DataFrame([[year, location, euroClass, roadType, A, B, maxres]], columns=['year', 'area', 'euro', 'roadType', 'petrol', 'diesel', 'maximumFitResidual'])
          carRatios = carRatios.append(ratioSingle)
          carRatios['version'] = versionForOutPut

        print('      Writing to file')
        if first:
          # Save to a new csv file.
          carRatios.to_csv(fileNameCSVNotComplete, index=False)
          first = False
        else:
          # Append to the csv file.
          carRatios.to_csv(fileNameCSVNotComplete, mode='a', header=False, index=False)
        toceuro = time.clock()
        print('      Processing for euro {} complete in {}.'.format(euroClass, secondsToString(toceuro-ticeuro, form='long')))
      tocyear = time.clock()
      print('      Processing for year {} complete in {}.'.format(year, secondsToString(tocyear-ticyear, form='long')))
    tocloc = time.clock()
    print('      Processing for area {} complete in {}.'.format(location, secondsToString(tocloc-ticloc, form='long')))

  shutil.move(fileNameCSVNotComplete, fileNameCSV)
  print('Processing complete. Output saved in the following files.')
  print('  {}'.format(fileNameCSV))
  if not keepTempFiles:
    print('Deleting temporary files.')
    for tf in tempFilesCreated:
      os.remove(tf)

  toc = time.clock()
  print('Process complete in {}.'.format(secondsToString(toc-tic,  form='long')))


def processEFT(fileName, locations, splitBusCoach=False, keepTempFiles=False):
  tic = time.clock()

  # Get the files are ready for processing.
  ahk_ahkpathG, fileNames, versions, versionsForOutput = prepareToExtract(fileName, locations)
  fileName = fileNames[0]
  version = versions[0]
  versionForOutPut = versionsForOutput[0]


  details = versionDetails[version]

  # Make a temporary copy of the filename, so that we do no processing on the
  # original. Just in case we brake it. Also define temporary file names and
  # output save locations, etc.
  [FN, FE] =  path.splitext(fileName)
  fileNameT = FN + '_TEMP_' + randomString() + FE
  fileNameCSV_ = fileName.replace(FE, '.csv')
  fileNameCSVNotComplete = fileNameCSV_.replace('prefilledValues', 'inProduction')
  fileNameCSV = fileNameCSV_.replace('prefilledValues', 'processedValues')
  fileNameDefaultProportions = fileNameCSV_.replace('prefilledValues', 'defaultProportions')
  fileNameDefaultProportionsNotComplete = fileNameCSV_.replace('prefilledValues', 'defaultProportionsInProduction')
  fileNameBusSplit = fileNameCSV_.replace('prefilledValues', 'BusCoachProportions')
  fileNameBusSplitNotComplete = fileNameCSV_.replace('prefilledValues', 'BusCoachProportionsInProduction')
  vi = 1
  while path.isfile(fileNameCSV):
    vi += 1
    fileNameCSV = fileNameCSV_.replace('prefilledValues', 'processedValues({})'.format(vi))
    fileNameDefaultProportions  = fileNameCSV_.replace('prefilledValues', 'defaultProportions({})'.format(vi))
  shutil.copyfile(fileName, fileNameT)

  # Create the Excel Application object.
  excel = win32.gencache.EnsureDispatch('Excel.Application')

  # And now start the processing!
  first = True
  tempFilesCreated = [fileNameT]
  for location in locations:
    ticloc = time.clock()
    print('Location: {}'.format(location))
    for year in years:
      ticyear = time.clock()
      print('  Year: {}'.format(year))
      busCoachSplit = pd.DataFrame(columns=['year', 'area', 'version', 'type', 'Buses', 'Coaches'])
      for euroClass in euroClasses:
        ticeuro = time.clock()
        print('    Euro class: {}'.format(euroClass))
        if first:
          excel, newSavedFile, defaultProportions, k = runAndExtract(excel, fileNameT, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut, checkEuroClasses=True)
        else:
          excel, newSavedFile, defaultProportions, k = runAndExtract(excel, fileNameT, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut)
        tempFilesCreated.append(newSavedFile)
        # Now get the output values as a dataframe.
        print('      Done, reading output values.')
        output = extractOutput(newSavedFile, versionForOutPut, year, location, euroClass, details)
        if splitBusCoach:
          #if euroClass == euroClasses[-1]:
          print('      Done, splitting buses from coaches.')
          excel, newSavedFileBus, k, busCoachRatio = runAndExtract(excel, fileNameT, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut, DoMCycles=False, DoBusCoach=True, busCoach='bus')
          excel, newSavedFileCoa, k, busCoachRatio = runAndExtract(excel, fileNameT, location, year, euroClass, ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut, DoMCycles=False, DoBusCoach=True, busCoach='coach')
          tempFilesCreated.append(newSavedFileBus)
          tempFilesCreated.append(newSavedFileCoa)
          outputBus = extractOutput(newSavedFileBus, versionForOutPut, year, location, euroClass, details)
          outputCoa = extractOutput(newSavedFileCoa, versionForOutPut, year, location, euroClass, details)
          outputBus = outputBus[outputBus['vehicle'] == '5. Buses and Coaches']
          outputCoa = outputCoa[outputCoa['vehicle'] == '5. Buses and Coaches']
          outputBus['vehicle'] = '5a. Buses'
          outputCoa['vehicle'] = '5b. Coaches'
          # Remove the bus and Coach rows from the output.
          output= output[output['vehicle'] != '5. Buses and Coaches']
          # And append the 'Bus' and 'Coach' only rows.
          output = output.append(outputBus)
          output = output.append(outputCoa)
        output = output.sort_values(['year', 'area', 'type', 'euro', 'speed', 'vehicle'])
        print('      Writing to file')
        if first:
          # Save to a new csv file.
          output.to_csv(fileNameCSVNotComplete, index=False)
          defaultProportions.to_csv(fileNameDefaultProportionsNotComplete, index=False)
          first = False
        else:
          # Append to the csv file.
          output.to_csv(fileNameCSVNotComplete, mode='a', header=False, index=False)
          defaultProportions.to_csv(fileNameDefaultProportionsNotComplete, mode='a', header=False, index=False)
        toceuro = time.clock()
        print('      Processing for euro {} complete in {}.'.format(euroClass, secondsToString(toceuro-ticeuro, form='long')))
      if splitBusCoach:
        busCoachSplitRow = pd.DataFrame([[year, location, versionForOutPut,
                                       'Motorway', busCoachRatio['bus_mw'],
                                       busCoachRatio['coach_mw']]],
                                     columns=['year', 'area', 'version',
                                              'type', 'Buses', 'Coaches'])
        busCoachSplit = busCoachSplit.append(busCoachSplitRow)
        busCoachSplitRow = pd.DataFrame([[year, location, versionForOutPut,
                                       'Non-Motorway', busCoachRatio['bus_non_mw'],
                                       busCoachRatio['coach_non_mw']]],
                                     columns=['year', 'area', 'version',
                                              'type', 'Buses', 'Coaches'])
        busCoachSplit = busCoachSplit.append(busCoachSplitRow)
        if location == locations[0]:
          busCoachSplit.to_csv(fileNameBusSplitNotComplete, index=False)
        else:
          busCoachSplit.to_csv(fileNameBusSplitNotComplete, mode='a', header=False, index=False)
      tocyear = time.clock()
      print('      Processing for year {} complete in {}.'.format(year, secondsToString(tocyear-ticyear, form='long')))
    tocloc = time.clock()
    print('      Processing for area {} complete in {}.'.format(location, secondsToString(tocloc-ticloc, form='long')))
  shutil.move(fileNameCSVNotComplete, fileNameCSV)
  shutil.move(fileNameDefaultProportionsNotComplete, fileNameDefaultProportions)
  shutil.move(fileNameBusSplitNotComplete, fileNameBusSplit)
  print('Processing complete. Output saved in the following files.')
  print('  {}'.format(fileNameCSV))
  print('  {}'.format(fileNameDefaultProportions))
  if not keepTempFiles:
    print('Deleting temporary files.')
    for tf in tempFilesCreated:
      os.remove(tf)
  toc = time.clock()
  excel.Quit()
  print('Process complete in {}.'.format(secondsToString(toc-tic, form='long')))

if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Extract emission values from the EFT')
  parser.add_argument('--version', '-v', metavar='version number',
                      type=float, nargs='?', default=7.0,
                      choices=availableVersions,
                      help="The EFT version number. One of {}. Default 7.0.".format(", ".join(str(v) for v in availableVersions)))
  parser.add_argument('--area', '-a', metavar='areas',
                      type=str, nargs='*', default='all',
                      help="The areas to be processed. One or more of '{}'. Default 'all'.".format("', '".join(availableAreas)))
  parser.add_argument('--years', '-y', metavar='year',
                      type=int, nargs='*', default=-9999,
                      choices=range(2008, 2031),
                      help="The year or years to be processed. Default 'all'")
  parser.add_argument('--euros', '-e', metavar='euro classes',
                      type=int, nargs='*', default=-9999,
                      choices = availableEuros,
                      help="The euro class or classes to be processed. One of more number between 0 and 6. Default 0 1 2 3 4 5 6.")
  parser.add_argument('--mode', '-m', metavar='mode',
                      type=str, nargs='?', default=availableModes[0],
                      choices=availableModes,
                      help="The mode. One of '{}'. Default '{}'.".format("', '".join(availableModes), availableModes[0]))
  parser.add_argument('--keeptemp', metavar='keeptemp',
                      type=bool,  nargs='?', default=False,
                      help="Whether to keep or delete temporary files. Boolean. Default False (delete).")
  parser.add_argument('--inputfile', '-i', metavar='input file',
                      type=str,   nargs='?', default=None,
                      help="The file to process. If set then version will be ignored.")
  args = parser.parse_args()

  version = args.version
  mode = args.mode
  inputfile = args.inputfile
  if inputfile is not None:
    version = extractVersion(inputfile)
  else:
    inputfile = getInputFile(mode, version)
  if version == 6.0:
    allowedYears = range(2008, 2031)
  else:
    allowedYears = range(2013, 2031)
  area = args.area
  if area == 'all':
    area = availableAreas
  euroClasses = args.euros
  if euroClasses == -9999:
    euroClasses = range(7)
  years = args.years
  if years == -9999:
    years = allowedYears
  keepTempFiles = args.keeptemp

  if not all(y in allowedYears for y in years):
    raise ValueError('One or more years are not allowed for the specified EFT version.')

  if mode == 'ExtractAll':
    processEFT(inputfile, area, keepTempFiles=keepTempFiles)
  elif mode == 'ExtractCarRatio':
    extractPetrolDieselCarProportions(inputfile[0], inputfile[1], area, keepTempFiles=keepTempFiles)
  elif mode == 'ExtractBus':
    processEFT(inputfile, area, splitBusCoach=True, keepTempFiles=keepTempFiles)
