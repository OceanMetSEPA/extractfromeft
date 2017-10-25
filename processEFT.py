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
import time
import shutil
import numpy as np
import pandas as pd
import win32com.client as win32


import EFT_Tools as tools

vehSplit2 = "Detailed Option 2"
vehSplit3 = "Detailed Option 3"

def extractPetrolDieselCarProportions(fileName, locations, keepTempFiles=False):
  tic = time.clock()

  # get the files are ready for processing.
  ahk_ahkpathG, fileNames, versions, versionsForOutput = tools.prepareToExtract(fileName, locations)
  fileName = fileNames[0]
  version = versions[0]
  versionForOutPut = versionsForOutput[0]

  # Now get the version dependent properties, mainly to do with which rows of
  # the spreadsheet contain which data.
  details = tools.versionDetails[version]

  # Make a temporary copy of the filename, so that we do no processing on the
  # original. Just in case we brake it. Also define temporary file names and
  # output save locations, etc.
  [FN, FE] =  path.splitext(fileName)
  fileName2T = FN + '_TEMP2_' + tools.randomString() + FE
  fileName3T = FN + '_TEMP3_' + tools.randomString() + FE
  shutil.copyfile(fileName, fileName2T)
  shutil.copyfile(fileName, fileName3T)
  fileNameCSVNotComplete = fileName.replace('empty', 'CarFuelRatios_InPreparation')
  fileNameCSVNotComplete = fileNameCSVNotComplete.replace(FE, '.csv')
  fileNameCSV = fileNameCSVNotComplete.replace('_InPreparation', '')

  # Create the Excel Application object.
  excel = win32.gencache.EnsureDispatch('Excel.Application')

  # And now start the processing!
  first = True
  tempFilesCreated = [fileName2T, fileName3T]

  inputData2 = tools.createEFTInput(vBreakdown=vehSplit2, roadTypes='all')
  inputData2 = inputData2.as_matrix()
  inputData3 = tools.createEFTInput(vBreakdown=vehSplit3, roadTypes='all')
  inputData3 = inputData3.as_matrix()
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
          excel, newSavedFile2, defaultProportions2, k = tools.runAndExtract(
                      fileName2T, location, year, euroClass, tools.ahk_exepath,
                      ahk_ahkpathG, vehSplit2, details, versionForOutPut,
                      checkEuroClasses=True, DoMCycles=False, inputData=inputData2,
                      excel=excel)
          excel, newSavedFile3, defaultProportions3, k = tools.runAndExtract(
                      fileName3T, location, year, euroClass, tools.ahk_exepath,
                      ahk_ahkpathG, vehSplit3, details, versionForOutPut,
                      checkEuroClasses=True, DoMCycles=False, inputData=inputData3,
                      excel=excel)
        else:
          excel, newSavedFile2, defaultProportions2, k = tools.runAndExtract(
                      fileName2T, location, year, euroClass, tools.ahk_exepath,
                      ahk_ahkpathG, vehSplit2, details, versionForOutPut,
                      DoMCycles=False, inputData=inputData2, excel=excel)
          excel, newSavedFile3, defaultProportions3, k = tools.runAndExtract(
                      fileName3T, location, year, euroClass, tools.ahk_exepath,
                      ahk_ahkpathG, vehSplit3, details, versionForOutPut,
                      DoMCycles=False, inputData=inputData3, excel=excel)
        tempFilesCreated.extend([newSavedFile2, newSavedFile3])
        print('      Done, reading output values.')

        # Now get the output values as a dataframe.
        output2 = tools.extractOutput(newSavedFile2, versionForOutPut, year, location, euroClass, details)
        output3 = tools.extractOutput(newSavedFile3, versionForOutPut, year, location, euroClass, details)
        print(set(output3['vehicle']))
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
          output_allCars = output2_r[output2_r['vehicle'] == 'Car']
          output_petrolCars = output3_r[output3_r['vehicle'] == 'Petrol Car']
          output_dieselCars = output3_r[output3_r['vehicle'] == 'Diesel Car']
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
        print('      Processing for euro {} complete in {}.'.format(euroClass, tools.secondsToString(toceuro-ticeuro, form='long')))
      tocyear = time.clock()
      print('      Processing for year {} complete in {}.'.format(year, tools.secondsToString(tocyear-ticyear, form='long')))
    tocloc = time.clock()
    print('      Processing for area {} complete in {}.'.format(location, tools.secondsToString(tocloc-ticloc, form='long')))

  shutil.move(fileNameCSVNotComplete, fileNameCSV)
  print('Processing complete. Output saved in the following files.')
  print('  {}'.format(fileNameCSV))
  if not keepTempFiles:
    print('Deleting temporary files.')
    for tf in tempFilesCreated:
      os.remove(tf)

  toc = time.clock()
  print('Process complete in {}.'.format(tools.secondsToString(toc-tic,  form='long')))


def processEFT(fileName, locations, splitBusCoach=False, keepTempFiles=False):
  tic = time.clock()

  # Get the files are ready for processing.
  ahk_ahkpathG, fileNames, versions, versionsForOutput = tools.prepareToExtract(fileName, locations)
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
  fileNameCSVNotComplete = fileNameCSV_.replace('empty', 'inProduction')
  fileNameCSV = fileNameCSV_.replace('empty', 'processedValues')
  fileNameDefaultProportions = fileNameCSV_.replace('empty', 'defaultProportions')
  fileNameDefaultProportionsNotComplete = fileNameCSV_.replace('empty', 'defaultProportionsInProduction')
  fileNameBusSplit = fileNameCSV_.replace('empty', 'BusCoachProportions')
  fileNameBusSplitNotComplete = fileNameCSV_.replace('empty', 'BusCoachProportionsInProduction')
  vi = 1
  while path.isfile(fileNameCSV):
    vi += 1
    fileNameCSV = fileNameCSV_.replace('empty', 'processedValues({})'.format(vi))
    fileNameDefaultProportions  = fileNameCSV_.replace('empty', 'defaultProportions({})'.format(vi))
  shutil.copyfile(fileName, fileNameT)

  # Create the Excel Application object.
  excel = win32.gencache.EnsureDispatch('Excel.Application')

  # And now start the processing!
  first = True
  tempFilesCreated = [fileNameT]
  inputData3 = tools.createEFTInput(vBreakdown=vehSplit3, roadTypes='all')
  inputData3 = inputData3.as_matrix()
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
          excel, newSavedFile, defaultProportions, k = tools.runAndExtract(
                  fileNameT, location, year, euroClass, tools.ahk_exepath,
                  ahk_ahkpathG, vehSplit3, details, versionForOutPut,
                  checkEuroClasses=True, inputData=inputData3, excel=excel)
        else:
          excel, newSavedFile, defaultProportions, k = tools.runAndExtract(
                  fileNameT, location, year, euroClass, tools.ahk_exepath,
                  ahk_ahkpathG, vehSplit3, details, versionForOutPut,
                  inputData=inputData3, excel=excel)
        tempFilesCreated.append(newSavedFile)
        # Now get the output values as a dataframe.
        print('      Done, reading output values.')
        output = tools.extractOutput(newSavedFile, versionForOutPut, year, location, euroClass, details)
        if splitBusCoach:
          #if euroClass == euroClasses[-1]:
          print('      Done, splitting buses from coaches.')
          excel, newSavedFileBus, k, busCoachRatio = tools.runAndExtract(excel, fileNameT, location, year, euroClass, tools.ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut, DoMCycles=False, DoBusCoach=True, busCoach='bus')
          excel, newSavedFileCoa, k, busCoachRatio = tools.runAndExtract(excel, fileNameT, location, year, euroClass, tools.ahk_exepath, ahk_ahkpathG, vehSplit3, details, versionForOutPut, DoMCycles=False, DoBusCoach=True, busCoach='coach')
          tempFilesCreated.append(newSavedFileBus)
          tempFilesCreated.append(newSavedFileCoa)
          outputBus = tools.extractOutput(newSavedFileBus, versionForOutPut, year, location, euroClass, details)
          outputCoa = tools.extractOutput(newSavedFileCoa, versionForOutPut, year, location, euroClass, details)
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
        print('      Processing for euro {} complete in {}.'.format(euroClass, tools.secondsToString(toceuro-ticeuro, form='long')))
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
      print('      Processing for year {} complete in {}.'.format(year, tools.secondsToString(tocyear-ticyear, form='long')))
    tocloc = time.clock()
    print('      Processing for area {} complete in {}.'.format(location, tools.secondsToString(tocloc-ticloc, form='long')))
  shutil.move(fileNameCSVNotComplete, fileNameCSV)
  shutil.move(fileNameDefaultProportionsNotComplete, fileNameDefaultProportions)
  if splitBusCoach:
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
  print('Process complete in {}.'.format(tools.secondsToString(toc-tic, form='long')))

if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Extract emission values from the EFT, broken down by euro class.')
  parser.add_argument('--version', '-v', metavar='version number',
                      type=float, nargs='?', default=7.0,
                      choices=tools.availableVersions,
                      help="The EFT version number. One of {}. Default 7.0.".format(", ".join(str(v) for v in tools.availableVersions)))
  parser.add_argument('--area', '-a', metavar='areas',
                      type=str, nargs='*', default='all',
                      choices=tools.availableAreas.append('all'),
                      help="The areas to be processed. One or more of '{}'. Default 'all'.".format("', '".join(tools.availableAreas)))
  parser.add_argument('--years', '-y', metavar='year',
                      type=int, nargs='*', default=-9999,
                      choices=range(2008, 2031),
                      help="The year or years to be processed. Default 'all'")
  parser.add_argument('--euros', '-e', metavar='euro classes',
                      type=int, nargs='*', default=-9999,
                      choices = tools.availableEuros,
                      help="The euro class or classes to be processed. One of more number between 0 and 6. Default 0 1 2 3 4 5 6.")
  parser.add_argument('--mode', '-m', metavar='mode',
                      type=str, nargs='?', default=tools.availableModes[0],
                      choices=tools.availableModes,
                      help="The mode. One of '{}'. Default '{}'.".format("', '".join(tools.availableModes), tools.availableModes[0]))
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
    version = tools.extractVersion(inputfile)
  else:
    inputfile = tools.getInputFile(version)
  if version == 6.0:
    availableYears = range(2008, 2031)
  else:
    availableYears = range(2013, 2031)
  area = args.area
  if area == 'all':
    area = tools.availableAreas
  euroClasses = args.euros
  if euroClasses == -9999:
    euroClasses = range(7)
  years = args.years
  if years == -9999:
    years = availableYears
  keepTempFiles = args.keeptemp

  if not all(y in availableYears for y in years):
    raise ValueError('One or more years are not allowed for the specified EFT version.')

  if mode == 'ExtractAll':
    processEFT(inputfile, area, keepTempFiles=keepTempFiles)
  elif mode == 'ExtractCarRatio':
    extractPetrolDieselCarProportions(inputfile, area, keepTempFiles=keepTempFiles)
  elif mode == 'ExtractBus':
    processEFT(inputfile, area, splitBusCoach=True, keepTempFiles=keepTempFiles)