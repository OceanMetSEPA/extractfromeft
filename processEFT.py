# -*- coding: utf-8 -*-
"""
Created on Wed Jun 07 10:37:25 2017

@author: edward.barratt

Routines to process Emission Factor Toolbox spreadsheets and extract emission
rates for vehicle classes against year and euro class.
"""
from __future__ import print_function

import os
import sys
from os import path
from time import sleep
import subprocess
import shutil
import random
import string
import numpy as np
import pandas as pd
import win32com.client as win32

workingDir = os.getcwd()

ahkexe = 'C:\Program Files\AutoHotkey\AutoHotkey.exe'
ahkahk = 'closeWarning.ahk'

versionDetails = {}
versionDetails[7.4] = {}
versionDetails[7.4]['vehRowStarts'] = [69, 79, 91, 101, 114, 130, 146, 161]
versionDetails[7.4]['vehRowEnds']   = [76, 87, 98, 109, 125, 141, 157, 172]
versionDetails[7.4]['vehRowStartsMC'] = [177, 183, 189, 195, 201, 207]
versionDetails[7.4]['vehRowEndsMC']   = [182, 188, 194, 200, 206, 212]
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

location = 'Scotland'
vehSplit = "Detailed Option 3"
years = range(2013, 2031)
euroClasses = range(7)

euroClassNameVariationsAll = euroClassNameVariations[0]
for ei in range(1,7):
  euroClassNameVariationsAll.extend(euroClassNameVariations[ei])
euroClassNameVariationsAll = list(set(euroClassNameVariationsAll))

availableVersions = versionDetails.keys()

EuroClassNameColumns = ["A", "H"]
DefaultEuroColumns = ["B", "I"]
UserDefinedEuroColumns = ["D", "K"]
EuroClassNameColumnsMC = ["B", "H"]
DefaultEuroColumnsMC = ["C", "I"]

def randomString(N = 10):
  return ''.join(random.choice(string.ascii_uppercase + string.ascii_lowercase + string.digits) for x in range(N))

def romanNumeral(N):
  # Could write a function that deals with any, but I only need up to 10.
  RNs = [0, 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
  return RNs[N]

def euroSearchTerms(N):
  ES = euroClassNameVariations[N]
  return ES

def checkEuroClasses(workBook, vehRowStarts, vehRowEnds, EuroClassNameColumns):
  print("Checking all euro class names are understood.")
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
  print("  All understood.")


def specifyEuros(euroClass, workBook, vehRowStarts, vehRowEnds,
                 EuroClassNameColumns, DefaultEuroColumns, UserDefinedEuroColumns):
  # Define some euro class search terms.
  #euroSearchTerms_ = euroSearchTerms(euroClass)
  print("    Setting euro ratios to 100% for euro {}.".format(euroClass))
  ws_euro = workBook.Worksheets("UserEuro")
  for [vi, vehRowStart] in enumerate(vehRowStarts):
    print("      Setting euro ratios for {}.".format(ws_euro.Range("A{row}".format(row=vehRowStart-1)).Value))
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
            break
        if not got:
          print('      No values available for euro {}, trying euro {}.'.format(euroClass_, euroClass_-1))
          euroClass_ -= 1

      # Get the default proportions.
      defaultProportions = []
      for row in rowsToDo:
        propRange = "{col}{row}".format(col=defaultEuroCol, row=row)
        defaultProportion = ws_euro.Range(propRange).Value
        defaultProportions.append(defaultProportion)
      defaultProportions = np.array(defaultProportions)
      # Normalize them.
      if sum(defaultProportions) < 0.00001:
        defaultProportions = defaultProportions + 1
      userProportions = defaultProportions/sum(defaultProportions)
      # And set the values in the sheet.
      # Set all to zero first.
      userRange = "{col}{rstart}:{col}{rend}".format(col=userDefinedCol, rstart=vehRowStart, rend=vehRowEnd)
      ws_euro.Range(userRange).Value = 0
      # Then set the specific values.
      for [ri, row] in enumerate(rowsToDo):
        userRange = "{col}{row}".format(col=userDefinedCol, row=row)
        value = userProportions[ri]
        ws_euro.Range(userRange).Value = value
  print('    All complete')

def splitSourceNameS(row):
  s = row[SourceNameName]
  s, v = s.split(' - ')
  row['vehicle'] = v
  return int(s[1:])

def splitSourceNameV(row):
  s = row[SourceNameName]
  s, v = s.split(' - ')
  return v

def processEFT(fileName):
  global SourceNameName, AllLDVName, AllHDVName, ahkahk

  if not path.isfile(ahkexe):
    raise ValueError('The Autohotkey executable file {} could not be found.'.format(ahkexe))
  if not path.isfile(ahkahk):
    ahkahk_ = workingDir + '\\' + ahkahk
    if not path.isfile(ahkahk_):
      raise ValueError('The Autohotkey file {} could not be found.'.format(ahkahk))
    else:
      ahkahk = ahkahk_

  fileName = path.abspath(fileName)
  if not path.isfile(fileName):
    raise ValueError('Could not find {}.'.format(fileName))

  # See what version we're looking at.
  version = False
  for versiono in availableVersions:
    if fileName.find('v{:.1f}'.format(versiono)) >= 0:
      version = versiono
      version_ = versiono
      break
  if version:
    print('Processing EFT of version {}.'.format(version))
  else:
    # Not one that is predefined, see if we can get the version number.
    fv = fileName.find('v')
    fp = fileName.find('_prefilled')
    if (fv >= 0) and (fp >= 0):
      fl = fileName.find('.', fv, fp)
      if fl >= 0:
        verTry = fileName[fv+1:fl+2]
        try:
          version = float(verTry)
        except:
          pass
    if version:
      # Get closest version number
      versioncloseI = np.argmin(abs(np.array(availableVersions) - version))
      version_ = version
      versionp = availableVersions[versioncloseI]
      print('Unknown version {}, will process as version {}.'.format(version, versionp))

      version = versionp
    else:
      maxAvailableVersions = max(availableVersions)
      print('Cannot parse version number from filename, will attempt to process as version {}.'.format(maxAvailableVersions))
      version = maxAvailableVersions
      version_ = 'Unknown Version as {}'.format(maxAvailableVersions)
    print('You may wish to edit the versionDetails global variables to account for the new version.')

  vehRowStarts = versionDetails[version]['vehRowStarts']
  vehRowEnds = versionDetails[version]['vehRowEnds']
  vehRowStartsMC = versionDetails[version]['vehRowStartsMC']
  vehRowEndsMC = versionDetails[version]['vehRowEndsMC']
  SourceNameName = versionDetails[version]['SourceNameName']
  AllLDVName = versionDetails[version]['AllLDVName']
  AllHDVName = versionDetails[version]['AllHDVName']
  AllVehName = versionDetails[version]['AllVehName']
  PolName = versionDetails[version]['PolName']
  # Make a temporary copy of the filename.
  [FN, FE] =  path.splitext(fileName)
  fileNameT = FN + '_TEMP_' + randomString() + FE
  fileNameTm = fileNameT.replace(FE, '_.xlsm')
  fileNameCSV_ = fileName.replace(FE, '.csv')
  fileNameCSV = fileNameCSV_.replace('prefilledValues', 'processedValues')
  vi = 1
  while path.isfile(fileNameCSV):
    vi += 1
    fileNameCSV = fileNameCSV_.replace('prefilledValues', 'processedValues({})'.format(vi))
  shutil.copyfile(fileName, fileNameT)

  # Create the Excel Application object.
  excel = win32.gencache.EnsureDispatch('Excel.Application')
  first = True
  tempFilesCreated = [fileNameT]

  for year in years:
    print('Year: {}'.format(year))
    # Ensure that the correct detailed split is specified. Setting it will raise a
    # popup and delete the traffic array, so we want to avoid that.

    for euroClass in euroClasses:
      print('  Euro class: {}'.format(euroClass))
      # Open the document.
      subprocess.Popen([ahkexe, ahkahk])  # Closes the warning, when it appears.
      wb = excel.Workbooks.Open(fileNameT)
      excel.Visible = True

      # Set the default values in the Input Data sheet.
      ws_input = wb.Worksheets("Input Data")
      ws_input.Range("B4").Value = location
      ws_input.Range("B5").Value = year
      if ws_input.Range("B6").Value != vehSplit:
        raise ValueError('Traffic Format should be "{}".'.format(vehSplit))

      if first:
        checkEuroClasses(wb, vehRowStartsMC, vehRowEndsMC, EuroClassNameColumnsMC)
        checkEuroClasses(wb, vehRowStarts, vehRowEnds, EuroClassNameColumns)
      # Now we need to populate the UserEuro table with the defaults.
      excel.Application.Run("PasteDefaultEuroProportions")

      # Now specify that we only want the specified euro class.
      # Motorcycles first
      specifyEuros(euroClass, wb, vehRowStartsMC, vehRowEndsMC,
                   EuroClassNameColumnsMC, DefaultEuroColumnsMC, UserDefinedEuroColumns)
      # And all other vehicles
      specifyEuros(euroClass, wb, vehRowStarts, vehRowEnds,
                   EuroClassNameColumns, DefaultEuroColumns, UserDefinedEuroColumns)

      # Now run the tool.
      ws_input.Select()# = wb.Worksheets("Input Data")
      print('    Running EFT routine.')
      excel.Application.Run("RunEfTRoutine")
      # Save and Close. Saving as an xlsm, rather than a xlsb, file, so that it
      # can be opened by pandas.
      fsave = fileNameTm.replace('.xlsm', '({}E{}).xlsm'.format(year, euroClass))
      wb.SaveAs(fsave, win32.constants.xlOpenXMLWorkbookMacroEnabled)
      tempFilesCreated.append(fsave)
      wb.Close()
      sleep(1)
      print('    Done')

      # Now get the new values.
      ex = pd.ExcelFile(fsave)
      output = ex.parse("Output")
      output['year'] = year
      output['euro'] = euroClass
      output['version'] = version_
      output['speed'] = output.apply(splitSourceNameS, axis=1)
      output['vehicle'] = output.apply(splitSourceNameV, axis=1)
      output = output.drop(SourceNameName, 1)
      output = output.drop(AllLDVName, 1)
      output = output.drop(AllHDVName, 1)
      # pivot the table so each pollutant has a column.
      Pollutants = list(output[PolName].unique())
      output = output.rename(columns={PolName: 'RowIndex'}) # Because after the
      # pivot the 'column' name will become the index name.
      output = output.pivot_table(index=['year',
                                         'euro',
                                         'version',
                                         'speed',
                                         'vehicle'],
                                  columns='RowIndex',
                                  values=AllVehName)
      output = output.reset_index()
      renames = {}
      for Pol in Pollutants:
        if Pol == 'PM25':
          Pol_ = 'PM2.5'
        else:
          Pol_ = Pol
        renames[Pol] = '{} (g/km/s/veh)'.format(Pol_)
      output = output.rename(columns=renames)

      if first:
        allOutput = output
        first = False
      else:
        allOutput = allOutput.append(output)
  print('Saving to {}.'.format(fileNameCSV))
  allOutput.to_csv(fileNameCSV)
  print('Deleting temporary files.')
  for tf in tempFilesCreated:
    os.remove(tf)

if __name__ == '__main__':
  args = sys.argv
  if len(args) > 1:
    fNames = args[1:]
    for fName in fNames:
      processEFT(fName)

  #processEFT('C:\Users\edward.barratt\Documents\Modelling\EmissionFactors\EmissionsFactorsToolkit\EFT2014_v6.0.2_prefilledValues.xls')
  #processEFT('C:\Users\edward.barratt\Documents\Modelling\EmissionFactors\EmissionsFactorsToolkit\EFT2016_v7.0_prefilledValues.xlsb')
  #processEFT('C:\Users\edward.barratt\Documents\Modelling\EmissionFactors\EmissionsFactorsToolkit\EFT2017_v7.4_prefilledValues.xlsb')