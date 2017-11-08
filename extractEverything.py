
import os
from os import path
#import argparse
import time
import shutil
#import numpy as np
import pandas as pd
import win32com.client as win32


import EFT_Tools as tools


vehsplit = "Alternative Technologies"

def processEFT(fileName, locations, years, euroClasses=[0,1,2,3,4,5,6],
               splitBusCoach=False, splitSize=False, keepTempFiles=False,
               saveFile=None):
  tic = time.clock()

  if type(years) is not list:
    years = [years]
  if type(locations) is not list:
    locations = [locations]

  # Get the files are ready for processing.
  ahk_ahkpathG, fileNames, versions, versionsForOutput = tools.prepareToExtract(fileName, locations)
  fileName = fileNames[0]
  version = versions[0]
  versionForOutPut = versionsForOutput[0]

  details = tools.versionDetails[version]

  # Make a temporary copy of the filename, so that we do no processing on the
  # original. Just in case we break it. Also define temporary file names and
  # output save locations, etc.
  [FN, FE] =  path.splitext(fileName)
  fileNameT = FN + '_TEMP_' + tools.randomString(N=2) + FE
  fileNameCSV_ = fileName.replace(FE, '.csv')
  fileNameCSVNotComplete = fileNameCSV_.replace('empty', 'inProduction')
  fileNameCSV = fileNameCSV_.replace('empty', 'allExtracted')
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

  if saveFile is not None:
    fileNameCSV = saveFile

  # Create the Excel Application object.
  excel = win32.gencache.EnsureDispatch('Excel.Application')

  NO2F = tools.readNO2Factors(mode='ByEuro')

  # And now start the processing!
  first = True
  tempFilesCreated = [fileNameT]

  vehiclesToSkip=['Taxi (black cab)']

  if splitBusCoach:
    vehiclesToSkip.append('Bus and Coach')
    inputDataBusCoach = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToInclude=['Bus and Coach'])
    inputDataBusCoach = inputDataBusCoach.as_matrix()
  if splitSize:
    sizeVehs = details['sizeRowEnds'].keys()
    #vehiclesToSkip.extend(sizeVehs)
    inputsSize = {}
    for sV in sizeVehs:
      inputSize = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToInclude=[sV])
      inputsSize[sV] = inputSize.as_matrix()

  inputData = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToSkip=vehiclesToSkip)
  inputData = inputData.as_matrix()

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
          excel, newSavedFile, defaultProportions, k, w = tools.runAndExtract(
                  fileNameT, location, year, euroClass, tools.ahk_exepath,
                  ahk_ahkpathG, vehsplit, details, versionForOutPut,
                  checkEuroClasses=True, inputData=inputData, excel=excel)
        else:
          excel, newSavedFile, defaultProportions, k, w = tools.runAndExtract(
                  fileNameT, location, year, euroClass, tools.ahk_exepath,
                  ahk_ahkpathG, vehsplit, details, versionForOutPut,
                  inputData=inputData, excel=excel)
        tempFilesCreated.append(newSavedFile)
        # Now get the output values as a dataframe.
        print('      Done, reading output values.')
        output = tools.extractOutput(newSavedFile, versionForOutPut, year, location, euroClass, details)
        if splitBusCoach:
          print('      Done, splitting buses from coaches.')
          excel, newSavedFileBus, k, busCoachRatio, w = tools.runAndExtract(
                  fileNameT, location, year, euroClass, tools.ahk_exepath,
                  ahk_ahkpathG, vehsplit, details, versionForOutPut,
                  DoMCycles=False, DoBusCoach=True, inputData=inputDataBusCoach,
                  busCoach='bus', excel=excel)
          excel, newSavedFileCoa, k, busCoachRatio, w = tools.runAndExtract(
                  fileNameT, location, year, euroClass, tools.ahk_exepath,
                  ahk_ahkpathG, vehsplit, details, versionForOutPut,
                  DoMCycles=False, DoBusCoach=True, inputData=inputDataBusCoach,
                  busCoach='coach', excel=excel)
          tempFilesCreated.append(newSavedFileBus)
          tempFilesCreated.append(newSavedFileCoa)
          outputBus = tools.extractOutput(newSavedFileBus, versionForOutPut, year, location, euroClass, details)
          outputCoa = tools.extractOutput(newSavedFileCoa, versionForOutPut, year, location, euroClass, details)
          outputBus = outputBus[outputBus['vehicle'] == 'Bus and Coach']
          outputCoa = outputCoa[outputCoa['vehicle'] == 'Bus and Coach']
          outputBus['vehicle'] = 'Bus'
          # Remove the bus and Coach rows from the output.
          output= output[output['vehicle'] != 'Bus and Coach']
          # And append the 'Bus' and 'Coach' only rows.
          output = output.append(outputBus)
          output = output.append(outputCoa)
        output['weight'] = 'Default Split'
        if splitSize:
          print('      Done, splitting by vehicle size for the following vehicles: {}'.format(', '.join(sizeVehs)))
          for sV in sizeVehs:
            print('        {}...'.format(sV))
            ssta = details['sizeRowStarts'][sV]
            send = details['sizeRowEnds'][sV]
            for si, sizerow in enumerate(range(ssta, send+1)):
              print('          {} weight class {} of {}...'.format(sV, si+1, send-ssta+1))
              sizeConts = {'start': ssta, 'end': send, 'do': sizerow, 'name': sV}
              excel, newSavedFile, defaultProportions, k, weightname = tools.runAndExtract(
                      fileNameT, location, year, euroClass, tools.ahk_exepath,
                      ahk_ahkpathG, vehsplit, details, versionForOutPut,
                      inputData=inputsSize[sV], DoMCycles=False, excel=excel, sizeRow=sizeConts)
              print('            [{}]'.format(weightname))
              tempFilesCreated.append(newSavedFile)
              outputW = tools.extractOutput(newSavedFile, versionForOutPut, year, location, euroClass, details)
              outputW = outputW[outputW['vehicle'] == sV]
              outputW['weight'] = weightname
              output = output.append(outputW)

        output['fuel'] = output.apply(lambda row: tools.VehDetails[row['vehicle']]['Fuel'], axis=1)
        output['vehicle type'] = output.apply(lambda row: tools.VehDetails[row['vehicle']]['Veh'], axis=1)
        output['tech'] = output.apply(lambda row: tools.VehDetails[row['vehicle']]['Tech'], axis=1)
        output['NOx2NO2'] = output.apply(lambda row: NO2F[tools.VehDetails[row['vehicle']]['NOxVeh']][row['euro']], axis=1)
        output['NO2 (g/km/s/veh)'] = output['NOx (g/km/s/veh)']*output['NOx2NO2']
        output = output.sort_values(['year', 'area', 'type', 'euro', 'speed', 'vehicle type', 'fuel', 'vehicle', 'weight'])
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
  del(excel)
  print('Process complete in {}.'.format(tools.secondsToString(toc-tic, form='long')))


if __name__ == '__main__':
  for Y in range(2028, 2031):
    for E in range(0,7):
      saveFile = 'C:\Users\edward.barratt\Documents\Development\Python\ExtractFromEFT\input\EFT2017_v7.4_{}_E{}.csv'.format(Y, E)
      processEFT('C:\Users\edward.barratt\Documents\Development\Python\ExtractFromEFT\input\EFT2017_v7.4_empty.xlsb', 'Scotland', Y, euroClasses=[E], splitSize=True, splitBusCoach=True, saveFile=saveFile)