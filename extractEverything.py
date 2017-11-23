
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
               saveFile=None, splitEuroTech=False):
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
  #fileNameBusSplit = fileNameCSV_.replace('empty', 'BusCoachProportions')
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
  try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
  except TypeError:
    time.sleep(5)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
  excel.DisplayAlerts = False
  NO2F = tools.readNO2Factors(mode='ByEuro')

  # And now start the processing!
  first = True
  tempFilesCreated = [fileNameT]

  vehiclesToSkip=['Taxi (black cab)']

  if splitBusCoach:
    vehiclesToSkip.append('Bus and Coach')
    #inputDataBusCoach = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToInclude=['Bus and Coach'])

  if splitSize:
    sizeVehs = details['sizeRowEnds'].keys()
    #vehiclesToSkip.extend(sizeVehs)
  if splitEuroTech:
    techs = tools.euroClassTechnologies
  else:
    techs = ['All']

  inputData = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToSkip=vehiclesToSkip)
  #techs = ['Standard', 'EGR']
  FirstLoc = True
  for location in locations:
    ticloc = time.clock()
    print('Location: {}'.format(location))
    for year in years:
      ticyear = time.clock()
      print('  Year: {}'.format(year))
      busCoachSplit = pd.DataFrame(columns=['year', 'area', 'version', 'type', 'Buses', 'Coaches'])
      busCoachDone = False
      for euroClass in euroClasses:
        ticeuro = time.clock()
        print('    Euro class: {}'.format(euroClass))
        for tech in techs:
          tictech = time.clock()
          print('      Technology: {}'.format(tech))
          # Check to see if this technology is available for this euro class.
          if tech not in tools.euroClassNameVariations[euroClass].keys():
            print('        Not available for this euro class.')
            continue

          print('        Processing...')
          inputData = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all',
                                           vehiclesToSkip=vehiclesToSkip,
                                           tech=tech)

          if first:
            excel, newSavedFile, defaultProportions, k, w, gotTechs = tools.runAndExtract(
                    fileNameT, location, year, euroClass, tools.ahk_exepath,
                    ahk_ahkpathG, vehsplit, details, versionForOutPut,
                    inputData=inputData, excel=excel, tech=tech,
                    checkEuroClasses=True)
          else:
            excel, newSavedFile, defaultProportions, k, w, gotTechs = tools.runAndExtract(
                    fileNameT, location, year, euroClass, tools.ahk_exepath,
                    ahk_ahkpathG, vehsplit, details, versionForOutPut,
                    inputData=inputData, excel=excel, tech=tech)
          tempFilesCreated.append(newSavedFile)
          # Now get the output values as a dataframe.
          print('          Done, reading output values.')
          output = tools.extractOutput(newSavedFile, versionForOutPut, year, location, euroClass, details, techDetails=[tech, gotTechs])
          if splitBusCoach:
            print('          Done, splitting buses from coaches.')
            if tech in ['c', 'd']:
              print('            Not applicable for technology {}.'.format(tech))
            elif (euroClass in [5]) and (tech == 'Standard'):
              print('            Not applicable for technology {} for euro class {}.'.format(tech, euroClass))
            else:
              inputDataBusCoach = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all',
                                                       vehiclesToInclude=['Bus and Coach'],
                                                       tech=tech)
              excel, newSavedFileBus, k, busCoachRatio, w, gotTechsB = tools.runAndExtract(
                      fileNameT, location, year, euroClass, tools.ahk_exepath,
                      ahk_ahkpathG, vehsplit, details, versionForOutPut,
                      DoMCycles=False, DoBusCoach=True, inputData=inputDataBusCoach,
                      busCoach='bus', excel=excel, tech=tech)
              excel, newSavedFileCoa, k, busCoachRatio, w, gotTechsC = tools.runAndExtract(
                      fileNameT, location, year, euroClass, tools.ahk_exepath,
                      ahk_ahkpathG, vehsplit, details, versionForOutPut,
                      DoMCycles=False, DoBusCoach=True, inputData=inputDataBusCoach,
                      busCoach='coach', excel=excel, tech=tech)
              tempFilesCreated.append(newSavedFileBus)
              tempFilesCreated.append(newSavedFileCoa)
              outputBus = tools.extractOutput(newSavedFileBus, versionForOutPut, year, location, euroClass, details, techDetails=[tech, gotTechsB])
              outputCoa = tools.extractOutput(newSavedFileCoa, versionForOutPut, year, location, euroClass, details, techDetails=[tech, gotTechsC])
              outputBus = outputBus.loc[outputBus['vehicle'] == 'Bus and Coach']
              outputCoa = outputCoa.loc[outputCoa['vehicle'] == 'Bus and Coach']
              outputBus['vehicle'] = 'Bus'
              outputCoa['vehicle'] = 'Coach'
              # Remove the bus and Coach rows from the existing output.
              output= output[output['vehicle'] != 'Bus and Coach']
              # And append the 'Bus' and 'Coach' only rows.
              output = output.append(outputBus)
              output = output.append(outputCoa)
              busCoachDone = True

          output['weight'] = 'Default Split'
          if splitSize:
            print('          Done, splitting by vehicle size for the following vehicles: {}'.format(', '.join(sizeVehs)))
            if tech in ['c', 'd']:
              print('            Not applicable for technology {}.'.format(tech))
            elif (euroClass in [5]) and (tech == 'Standard'):
              print('            Not applicable for technology {} for euro class {}.'.format(tech, euroClass))
            else:
              for sV in sizeVehs:
                print('        {}...'.format(sV))
                inputSize = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all',
                                                 vehiclesToInclude=[sV],
                                                 tech=tech)
                ssta = details['sizeRowStarts'][sV]
                send = details['sizeRowEnds'][sV]
                for si, sizerow in enumerate(range(ssta, send+1)):
                  print('          {} weight class {} of {}...'.format(sV, si+1, send-ssta+1))
                  sizeConts = {'start': ssta, 'end': send, 'do': sizerow, 'name': sV}
                  excel, newSavedFile, defaultProportions, k, weightname , gotTechs= tools.runAndExtract(
                          fileNameT, location, year, euroClass, tools.ahk_exepath,
                          ahk_ahkpathG, vehsplit, details, versionForOutPut, tech=tech,
                          inputData=inputSize, DoMCycles=False, excel=excel, sizeRow=sizeConts)
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
          toctech = time.clock()
          print('      Processing for tech {} complete in {}.'.format(tech, tools.secondsToString(toctech-tictech, form='long')))
        toceuro = time.clock()
        print('    Processing for euro {} complete in {}.'.format(euroClass, tools.secondsToString(toceuro-ticeuro, form='long')))
      if splitBusCoach and busCoachDone:
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
        if FirstLoc:
          busCoachSplit.to_csv(fileNameBusSplitNotComplete, index=False)
          FirstLoc = False
        else:
          busCoachSplit.to_csv(fileNameBusSplitNotComplete, mode='a', header=False, index=False)
      tocyear = time.clock()
      print('  Processing for year {} complete in {}.'.format(year, tools.secondsToString(tocyear-ticyear, form='long')))
    tocloc = time.clock()
    print('Processing for area {} complete in {}.'.format(location, tools.secondsToString(tocloc-ticloc, form='long')))
  shutil.move(fileNameCSVNotComplete, fileNameCSV)
  shutil.move(fileNameDefaultProportionsNotComplete, fileNameDefaultProportions)
  #if splitBusCoach:
  #  shutil.move(fileNameBusSplitNotComplete, fileNameBusSplit)
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
  for Y in [2017]: #range(2028, 2031):
    for E in [6, 5, 4, 3, 2, 1, 0]: #range(0,7):
      print(Y, E)
      saveFile = 'C:/Users/edward.barratt/Documents/Development/Python/ExtractFromEFT/input/EFT2017_v8.0_{}_E{}.csv'.format(Y, E)
      originalFile = 'C:/Users/edward.barratt/Documents/Development/Python/ExtractFromEFT/input/EFT2017_v8.0_empty.xlsb'
      processEFT(originalFile, 'Scotland', Y, euroClasses=[E], splitSize=True, splitBusCoach=True, splitEuroTech=True, keepTempFiles=True, saveFile=saveFile)