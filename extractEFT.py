
import os
import sys
from os import path
import argparse
import subprocess
import logging
import time
from datetime import datetime
import shutil
import numpy as np
import pandas as pd
import win32com.client as win32


import EFT_Tools as tools


vehsplit = "Alternative Technologies"

ynow = datetime.now().year

def processEFT(fileName, outdir, locations, years, euroClasses=[99,0,1,2,3,4,5,6],
               splitBusCoach=False, splitSize=False, keepTempFiles=False,
               saveFile=None, splitEuroTech=False, completed=None):

  if type(years) is not list:
    years = [years]
  if type(locations) is not list:
    locations = [locations]

  if completed is None:
    completed = pd.DataFrame(columns=['area', 'year', 'euro', 'tech', 'saveloc'])

  # Get the files are ready for processing.
  ahk_ahkpathG, fileNames, versions, versionsForOutput = tools.prepareToExtract(fileName, locations)
  fileName = fileNames[0]
  version = versions[0]
  versionForOutPut = versionsForOutput[0]

  details = tools.versionDetails[version]

  # Make a temporary copy of the filename, so that we do no processing on the
  # original. Just in case we break it. Also define temporary file names and
  # output save locations, etc.
  [oP, FN] = path.split(fileName)
  tempdir = path.join(outdir, 'temp')
  fileNameT = path.join(tempdir, FN)
  shutil.copyfile(fileName, fileNameT)

  # Create the Excel Application object.
  try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
  except TypeError:
    time.sleep(5)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
  excel.DisplayAlerts = False
  NO2FEU = tools.readNO2Factors(mode='ByEuro')
  NO2FRT = tools.readNO2Factors(mode='ByRoadType')


  # And now start the processing!
  first = True
  tempFilesCreated = [fileNameT]

  #vehiclesToSkip=['Taxi (black cab)']

  if splitBusCoach:
    #vehiclesToSkip.append('Bus and Coach')
    BusesOptions = [True, False]
  else:
    BusesOptions = [None]
    #inputDataBusCoach = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToInclude=['Bus and Coach'])

  #if splitSize:
    #sizeVehs = details['sizeRowEnds'].keys()
    #vehiclesToSkip.extend(sizeVehs)
  techOptions = ['All']
  if splitEuroTech:
    techOptions.extend(tools.euroClassTechnologies)
  if splitSize:
    weights = list(range(max(np.array(details['weightRowEnds']) - np.array(details['weightRowStarts'])) + 1))
    weights.insert(0, 99)
  else:
    weights = [99]
  #vehiclesToSkipStandard = vehiclesToSkip.copy()

  vehsToSkipSt5 = ['Rigid HGV', 'Artic HGV', 'Bus and Coach',
                   'B100 Rigid HGV', 'B100 Artic HGV', 'B100 Bus',
                   'Hybrid Bus', 'B100 Coach']


  # Check that all euro class names are understood.
  if path.isfile(tools.ahk_exepath):
    subprocess.Popen([tools.ahk_exepath, ahk_ahkpathG])
  wb = excel.Workbooks.Open(fileNameT)
  excel.Visible = True
  tools.checkEuroClassesValid(wb, details['vehRowStartsMC'], details['vehRowEndsMC'], tools.EuroClassNameColumnsMC, Type=1)
  tools.checkEuroClassesValid(wb, details['vehRowStartsHB'], details['vehRowEndsHB'], tools.EuroClassNameColumnsMC, Type=2)
  tools.checkEuroClassesValid(wb, details['vehRowStarts'], details['vehRowEnds'], tools.EuroClassNameColumns, Type=0)
  wb.Close(True)

  #inputData = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToSkip=vehiclesToSkip)
  for loci, location in enumerate(locations):
    logger.info('{:02d}             Beginning processing for location {} of {}: "{}".'.format(loci+1, loci+1, len(locations), location))
    for yeari, year in enumerate(years):
      logger.info('{:02d} {:02d}          Beginning processing for year {} of {}: "{}".'.format(loci+1, yeari+1, yeari+1, len(years), year))
      #busCoachSplit = pd.DataFrame(columns=['year', 'area', 'version', 'type', 'Buses', 'Coaches'])
      #busCoachDone = False
      for euroi, euroClass in enumerate(euroClasses):
        logger.info('{:02d} {:02d} {:02d}       Beginning processing for euroclass {} of {}: "{}".'.format(loci+1, yeari+1, euroi+1, euroi+1, len(euroClasses), euroClass))
        if euroClass == 99:
          # Euro class of euro 99 means use default mix, and default mix of tech.
          logger.info('{:02d} {:02d} {:02d}       Euro class of 99 specifies using default euro mix, and default tech.'.format(loci+1, yeari+1, euroi+1))
          techs = ['All']
        else:
          techs = techOptions
        for techi, tech in enumerate(techs):
          logger.info('{:02d} {:02d} {:02d} {:02d}    Beginning processing for technology {} of {}: "{}".'.format(loci+1, yeari+1, euroi+1, techi+1, techi+1, len(techs), tech))

          # See if this is already completed.
          matchingRow = completed[(completed['area'] == location) & (completed['year'] == year) & (completed['euro'] == euroClass) & (completed['tech'] == tech)].index.tolist()
          if len(matchingRow) > 0:
            completedfile = completed.loc[matchingRow[0]]['saveloc']
            if completedfile == 'No File':
              logger.info('{:02d} {:02d} {:02d} {:02d}    Processing for these specifications has previously been skipped.'.format(loci+1, yeari+1, euroi+1, techi+1, completedfile))
            else:
              [oP, FNC] = path.split(completedfile)
              logger.info('{:02d} {:02d} {:02d} {:02d}    Processing for these specifications has already been completed.'.format(loci+1, yeari+1, euroi+1, techi+1))
              logger.info('{:02d} {:02d} {:02d} {:02d}    Results saved in {}.'.format(loci+1, yeari+1, euroi+1, techi+1, FNC))
            continue

          # Assign save locations.
          outputFileCSVinPrep = path.join(tempdir, '{}_{:04d}_{:02d}_{}_InPrep.csv'.format(location, year, euroClass, tech))
          outputFileCSV = path.join(outdir, '{}_{:04d}_{:02d}_{}.csv'.format(location, year, euroClass, tech))
          first = True

          # Check to see if this technology is available for this euro class.
          if tech not in tools.euroClassNameVariations[euroClass].keys():
            logger.info('{:02d} {:02d} {:02d} {:02d}    Not available for this euro class.'.format(loci+1, yeari+1, euroi+1, techi+1))
            logger.info('{:02d} {:02d} {:02d} {:02d}    SKIPPED (area, year, euro, tech, saveloc): {}, {}, {}, {}, {}.'.format(loci+1, yeari+1, euroi+1, techi+1, location, year, euroClass, tech, 'No File'))
            continue

          for doBus in BusesOptions:
            if doBus is None:
              pass
            elif doBus:
              logger.info('{:02d} {:02d} {:02d} {:02d}    Buses and coaches.'.format(loci+1, yeari+1, euroi+1, techi+1))
              if tech in ['c', 'd']:
                logger.info('{:02d} {:02d} {:02d} {:02d}    Not available for technology {}.'.format(loci+1, yeari+1, euroi+1, techi+1, tech))
                continue
              elif (euroClass in [5]) and (tech == 'Standard'):
                logger.info('{:02d} {:02d} {:02d} {:02d}    Not applicable for technology {} for euro class {}.'.format(loci+1, yeari+1, euroi+1, techi+1, tech, euroClass))
                continue
            else:
              logger.info('{:02d} {:02d} {:02d} {:02d}    All vehicles except buses and coaches.'.format(loci+1, yeari+1, euroi+1, techi+1))

            for weighti, weight in enumerate(weights):
              logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} Beginning processing for weight row {} of {}.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1, weighti+1, len(weights)))
              if weight == 99:
                logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} Weight row 99 specifies using the default weight mix.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))
              if doBus is None:
                # Extract buses and coaches together, i.e. extract them, along
                # with all other vehicles and don't treat them any differenctly.
                vehs2Skip = ['Taxi (black cab)']
                if (euroClass in [5]) and (tech == 'Standard'):
                  vehs2Skip = vehs2Skip + vehsToSkipSt5
                excel, newSavedFile, b, k, weightclassnames, gotTechs = tools.runAndExtract(
                       fileNameT, vehsplit, details, location, year, euroClass,
                       tools.ahk_exepath, ahk_ahkpathG, versionForOutPut,
                       tech=tech, sizeRow=weight, DoHybridBus=True, DoBusCoach=True,
                       excel=excel, vehiclesToSkip=vehs2Skip)
                if newSavedFile is None:
                  output = None
                else:
                  tempFilesCreated.append(newSavedFile)
                  # Now get the output values as a dataframe.
                  output = tools.extractOutput(newSavedFile, versionForOutPut, year, location, euroClass, details, techDetails=[tech, gotTechs])
              elif not doBus:
                vehs2Skip = ['Taxi (black cab)', 'Bus and Coach', 'B100 Bus',
                               'CNG Bus', 'Biomethane Bus', 'Biogas Bus',
                               'Hybrid Bus', 'FCEV Bus', 'B100 Coach']
                if (euroClass in [5]) and (tech == 'Standard'):
                  vehs2Skip = vehs2Skip + vehsToSkipSt5
                # Extract all vehicles except buses and coaches.
                excel, newSavedFile, b, k, weightclassnames, gotTechs = tools.runAndExtract(
                       fileNameT, vehsplit, details, location, year, euroClass,
                       tools.ahk_exepath, ahk_ahkpathG, versionForOutPut,
                       tech=tech, sizeRow=weight, DoHybridBus=False, DoBusCoach=False,
                       excel=excel, vehiclesToSkip=vehs2Skip)
                if newSavedFile is None:
                  output = None
                else:
                  tempFilesCreated.append(newSavedFile)
                  # Now get the output values as a dataframe.
                  output = tools.extractOutput(newSavedFile, versionForOutPut, year, location, euroClass, details, techDetails=[tech, gotTechs])
                  # Add weight details.
                  output['weight'] = 'None'
                  for vehclass, wcn in weightclassnames.items():
                    for vehclass2 in tools.in2outVeh[vehclass]:
                      vehclass
                      vehclass2
                      wcn
                      output.loc[output.vehicle == vehclass2, 'weight'] = '{} - {}'.format(vehclass, wcn)
              else:
                # Extract only buses and coaches, and split them.
                logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} Buses...'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))
                excel, newSavedFileBus, b, busCoachRatio, weightclassnames, gotTechsB = tools.runAndExtract(
                      fileNameT, vehsplit, details, location, year, euroClass,
                      tools.ahk_exepath, ahk_ahkpathG, versionForOutPut,
                      tech=tech, sizeRow=weight, DoHybridBus=True, DoBusCoach=True,
                      DoMCycles=False, excel=excel, busCoach='bus')
                if newSavedFileBus is None:
                  gotBus = False
                  logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} No buses for this weight class.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))
                else:
                  tempFilesCreated.append(newSavedFileBus)
                  outputBus = tools.extractOutput(newSavedFileBus, versionForOutPut, year, location, euroClass, details, techDetails=[tech, gotTechsB])
                  outputBus = outputBus.loc[[x in ['B100 Bus', 'Bus and Coach', 'Hybrid Bus'] for x in outputBus['vehicle']]]
                  outputBus.loc[outputBus.vehicle == 'Bus and Coach', 'vehicle'] = 'Bus'
                  outputBus['weight'] = 'Bus - {}'.format(weightclassnames['Bus'])
                  gotBus = True
                logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} Coaches...'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))
                excel, newSavedFileCoa, b, busCoachRatio, weightclassnames, gotTechsC = tools.runAndExtract(
                      fileNameT, vehsplit, details, location, year, euroClass,
                      tools.ahk_exepath, ahk_ahkpathG, versionForOutPut,
                      tech=tech, sizeRow=weight, DoHybridBus=False, DoBusCoach=True,
                      DoMCycles=False, excel=excel, busCoach='coach')
                if newSavedFileCoa is None:
                  gotCoach = False
                  logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} No buses for this weight class.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))
                else:
                  tempFilesCreated.append(newSavedFileCoa)
                  outputCoa = tools.extractOutput(newSavedFileCoa, versionForOutPut, year, location, euroClass, details, techDetails=[tech, gotTechsC])
                  outputCoa = outputCoa.loc[[x in ['B100 Coach', 'Bus and Coach'] for x in outputCoa['vehicle']]]
                  outputCoa.loc[outputCoa.vehicle == 'Bus and Coach', 'vehicle'] = 'Coach'
                  outputCoa['weight'] = 'Coach - {}'.format(weightclassnames['Coach'])
                  gotCoach = True
                if gotBus and gotCoach:
                  output = pd.concat([outputBus, outputCoa], axis=0)
                elif gotBus:
                  output = outputBus
                elif gotCoach:
                  output = outputCoa
                else:
                  output = None

              if output is None:
                # No output for this weightclass, there'll be none for any higher either.
                logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} No output for this weight class. Skipping any higher weight classes.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))
                break
              # Now add fuel information, etc.
              try:
                output['fuel'] = output.apply(lambda row: tools.VehDetails[row['vehicle']]['Fuel'], axis=1)
              except E:
                print(output.head(5))
                print(output.head(5))
                print(output['vehicle'].unique())
                print(tools.VehDetails)
                raise E
              output['fuel'] = output.apply(lambda row: tools.VehDetails[row['vehicle']]['Fuel'], axis=1)
              output['vehicle type'] = output.apply(lambda row: tools.VehDetails[row['vehicle']]['Veh'], axis=1)
              output['tech'] = output.apply(lambda row: tools.VehDetails[row['vehicle']]['Tech'], axis=1)

              if euroClass == 99:
                output['NOx2NO2'] = output.apply(lambda row: NO2FRT[row['type']][tools.VehDetails[row['vehicle']]['Veh']][row['year']], axis=1)
              else:
                output['NOx2NO2'] = output.apply(lambda row: NO2FEU[tools.VehDetails[row['vehicle']]['NOxVeh']][row['euro']], axis=1)

              output['NO2 (g/km/s/veh)'] = output['NOx (g/km/s/veh)']*output['NOx2NO2']
              output = output.sort_values(['year', 'area', 'type', 'euro', 'speed', 'vehicle type', 'fuel', 'vehicle', 'weight'])
              logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} Writing {} rows to file.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1, output.shape[0]))
              if first:
                # Save to a new csv file.
                output.to_csv(outputFileCSVinPrep, index=False)
                #defaultProportions.to_csv(fileNameDefaultProportionsNotComplete, index=False)
                first = False
              else:
                # Append to the csv file.
                output.to_csv(outputFileCSVinPrep, mode='a', header=False, index=False)
                #defaultProportions.to_csv(fileNameDefaultProportionsNotComplete, mode='a', header=False, index=False)
              logger.info('{:02d} {:02d} {:02d} {:02d} {:02d} Processing complete for weight row {} of {}.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1, weighti+1, len(weights)))

            if doBus is None:
              pass
            elif doBus:
              logger.info('{:02d} {:02d} {:02d} {:02d}    Processing for Buses and Coaches complete.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))
            else:
              logger.info('{:02d} {:02d} {:02d} {:02d}    Processing for non-Buses complete.'.format(loci+1, yeari+1, euroi+1, techi+1, weighti+1))

          logger.info('{:02d} {:02d} {:02d} {:02d}    Processing complete for technology {} of {}: "{}".'.format(loci+1, yeari+1, euroi+1, techi+1, techi+1, len(techs), tech))
          logger.info('{:02d} {:02d} {:02d} {:02d}    Results saved in {}.'.format(loci+1, yeari+1, euroi+1, techi+1, outputFileCSV))
          logger.info('{:02d} {:02d} {:02d} {:02d}    COMPLETED (area, year, euro, tech, saveloc): {}, {}, {}, {}, {}.'.format(loci+1, yeari+1, euroi+1, techi+1, location, year, euroClass, tech, outputFileCSV))
          shutil.move(outputFileCSVinPrep, outputFileCSV)
        logger.info('{:02d} {:02d} {:02d}       Processing complete for euroclass {} of {}: "{}".'.format(loci+1, yeari+1, euroi+1, euroi+1, len(euroClasses), euroClass))
      logger.info('{:02d} {:02d}          Processing complete for year {} of {}: "{}".'.format(loci+1, yeari+1, yeari+1, len(years), year))
    logger.info('{:02d}             Processing complete for location {} of {}: "{}".'.format(loci+1, loci+1, len(locations), location))
  logger.info('Process Complete.')
  excel.Quit()
  del(excel)

def parseArgs():
  parser = argparse.ArgumentParser(description=("Extract emission values from "
                                                "the EFT, broken down by a "
                                                "choice of year, location, euro "
                                                "class, technology and weight class."))

  parser.add_argument('inputfile', metavar='input file',
                      type=str,   nargs=1,
                      help=("The file to process. This needs to be set up "
                            "correctly (DETAILS!)."))
  parser.add_argument('--outputdir', '-o', metavar='output directory',
                      type=str,   nargs='?', default=None,
                      help=("The directory in which to save output files. If "
                            "the directory does not exist then it will be "
                            "created, assuming required permissions, etc. "
                            "Default current working directory."))
  parser.add_argument('--area', '-a', metavar='areas',
                      type=str, nargs='*', default='Scotland',
                      choices=tools.availableAreas.append('all'),
                      help=("The areas to be processed. One or more of '{}'. "
                            "Default 'Scotland'.").format("', '".join(tools.availableAreas)))
  parser.add_argument('--years', '-y', metavar='year',
                      type=int, nargs='*', default=ynow,
                      choices=range(2008, 2031),
                      help="The year or years to be processed. Default {}.".format(ynow))
  euroChoices = [99]
  euroChoices.extend(tools.availableEuros)
  parser.add_argument('--euros', '-e', metavar='euro classes',
                      type=int, nargs='*', default=99,
                      choices = euroChoices,
                      help=("The euro class or classes to be processed. One of "
                            "more number between 0 and 6, or 99 which will "
                            "instruct the code to use the default euro breakdown "
                            "for the year specified. Default 99."))
  parser.add_argument('--keeptemp',
                      type=bool,  nargs='?', default=False,
                      help=("Whether to keep or delete temporary files. "
                            "Boolean. Default False (delete)."))
  parser.add_argument('--loggingmode',
                      nargs='?', default='INFO',
                      help=("The logging mode."))

  return parser.parse_args()

def prepareDir(outputDirP, mode='makedir'):

  if outputDirP is None:
    # Just use the current working directory. Don't search for any pre-started
    # output files.
    outputDir = os.getcwd()
    logfilename = path.join(outputDir, 'extractEFT_{}.log'.format(datetime.now().strftime('%Y%m%d%H%M%S')))
    if not path.isdir(path.join(outputDir, 'temp')):
      os.makedirs(path.join(outputDir, 'temp'))
  else:
    outputDir = outputDirP
    logfilename = path.join(outputDir, 'extractEFT.log')
    # Does it already exist?
    if path.isdir(outputDir):
      # It does! See if it's empty.
      contents = os.listdir(outputDir)
      if len(contents) == 0:
        # No contents, so we are beginning afresh. Create the temp dir and the log.
        os.makedirs(path.join(outputDir, 'temp'))
      else:
        # There are contents. Is one of them called log?
        if path.isfile(logfilename):
          # There is already a log file. Ask the user if they wish to expand on
          # work already started.
          Append = input(('It looks like processing has already begun in this '
                          'directory. Would you like to continue processing '
                          'where it was left off, based on the contents of the '
                          'log file. [y/n]'))
          if Append.lower() in ['yes', 'y']:
            if not path.isdir(path.join(outputDir, 'temp')):
              os.makedirs(path.join(outputDir, 'temp'))
          else:
            print(('Processing cannot continue because the designated directory '
                   'already contains a log file. Either specify a new directory, '
                   'or delete the log file to start afresh.'))
            exit()
        else:
          # There is not, so we are beginning afresh here too.
          if not path.isdir(path.join(outputDir, 'temp')):
            os.makedirs(path.join(outputDir, 'temp'))
      pass
    else:
      # It doesn't. Create it, and create a log file and temporary file directory.
      os.makedirs(outputDir)
      os.makedirs(path.join(outputDir, 'temp'))
  return outputDir, logfilename

def main():
  global logger

  # Parse the input arguments.
  pargs = parseArgs()

  #tools.combineFiles(pargs.outputdir)

  # Get the assigned directory and prepare temporary directories and log files.
  outputDir, logfilename = prepareDir(pargs.outputdir)

  # Create the log file.
  logger = logging.getLogger('extractEFT')
  logger.setLevel(logging.INFO)

  fileFormatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
  streamFormatter = logging.Formatter('%(asctime)s - %(message)s')
  logfilehandler = logging.FileHandler(logfilename)
  logfilehandler.setFormatter(fileFormatter)
  logstreamhandler = logging.StreamHandler()
  logstreamhandler.setFormatter(streamFormatter)
  logger.addHandler(logfilehandler)
  logger.addHandler(logstreamhandler)
  logger.info('Program started with command: "{}"'.format(' '.join(sys.argv)))
  logger.info('Input arguments parsed as: {}'.format(pargs))

  # Read the log file to see if any combination of location, year, euroclass,
  # and tech have already been completed.
  completed = tools.getCompletedFromLog(logfilename, mode='both')

  # Run the processing routine.
  processEFT(pargs.inputfile, pargs.outputdir, pargs.area, pargs.years,
             euroClasses=pargs.euros, splitSize=True, splitBusCoach=True,
             splitEuroTech=True, keepTempFiles=pargs.keeptemp,
             completed=completed)

if __name__ == '__main__':
  main()