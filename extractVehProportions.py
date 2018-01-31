
import os
import sys
import ast
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

def extractVehProps(fileName, outdir, locations, years,
                    keepTempFiles=False, saveFile=None,
                    completed=None):

  loggerM = logger.getChild('extractVehProps')

  if type(years) is not list:
    years = [years]
  if type(locations) is not list:
    locations = [locations]

  if completed is None:
    completed = pd.DataFrame(columns=['area', 'year', 'saveloc'])

  # Get the files ready for processing.
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
  [FN, FEXT] = path.splitext(FN)
  FN = FN+datetime.strftime(datetime.now(), '%Y%m%d%H%M%S')+FEXT
  fileNameT = path.join(tempdir, FN)
  shutil.copyfile(fileName, fileNameT)

  # Create the Excel Application object.
  try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
  except TypeError:
    time.sleep(5)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
  excel.DisplayAlerts = False

  # And now start the processing!
  first = True
  tempFilesCreated = [fileNameT]

  vehsToSkipSt5 = ['Rigid HGV', 'Artic HGV', 'Bus and Coach',
                   'B100 Rigid HGV', 'B100 Artic HGV', 'B100 Bus',
                   'Hybrid Bus', 'B100 Coach']

  # Check that all euro class names are understood.
  if path.isfile(tools.ahk_exepath):
    subprocess.Popen([tools.ahk_exepath, ahk_ahkpathG])
  wb = excel.Workbooks.Open(fileNameT)
  excel.Visible = True
  tools.checkEuroClassesValid(wb, details['vehRowStartsMC'], details['vehRowEndsMC'], tools.EuroClassNameColumnsMC, Type=1, logger=loggerM)
  tools.checkEuroClassesValid(wb, details['vehRowStartsHB'], details['vehRowEndsHB'], tools.EuroClassNameColumnsMC, Type=2, logger=loggerM)
  tools.checkEuroClassesValid(wb, details['vehRowStarts'], details['vehRowEnds'], tools.EuroClassNameColumns, Type=0, logger=loggerM)
  wb.Close(True)

  #inputData = tools.createEFTInput(vBreakdown=vehsplit, roadTypes='all', vehiclesToSkip=vehiclesToSkip)
  for loci, location in enumerate(locations):
    loggerM.info('{:02d}       Beginning processing for location {} of {}: "{}".'.format(loci+1, loci+1, len(locations), location))
    for yeari, year in enumerate(years):
      loggerM.info('{:02d} {:02d}    Beginning processing for year {} of {}: "{}".'.format(loci+1, yeari+1, yeari+1, len(years), year))

      # See if this is already completed.
      matchingRow = completed[(completed['area'] == location) & (completed['year'] == year)].index.tolist()
      if len(matchingRow) > 0:
        completedfile = completed.loc[matchingRow[0]]['saveloc']
        [oP, FNC] = path.split(completedfile)
        loggerM.info('{:02d} {:02d}    Processing for these specifications has already been completed.'.format(loci+1, yeari+1))
        loggerM.info('{:02d} {:02d}    Results saved in {}.'.format(loci+1, yeari+1, FNC))
        continue

      # Assign save locations.
      EuroAllFN = path.join(outdir, '{}_{:04d}_AllEuroProportions.csv'.format(location, year))
      WeightFN = path.join(outdir, '{}_{:04d}_WeightProportions.csv'.format(location, year))
      EuroConsFN = path.join(outdir, '{}_{:04d}_ConsolidatedEuroProportions.csv'.format(location, year))
      #outputFileCSV = path.join(outdir, '{}_{:04d}_{:02d}_{}.csv'.format(location, year, euroClass, tech))
      #first = True

      df_allEuros, df_weights, df_consEuros = tools.readProportions(fileNameT, details, location, year,
                                          tools.ahk_exepath, ahk_ahkpathG,
                                          versionForOutPut, excel=excel, logger=loggerM)
      loggerM.info('{:02d} {:02d}          Processing complete for year {} of {}: "{}".'.format(loci+1, yeari+1, yeari+1, len(years), year))
      loggerM.info('{:02d} {:02d}          Saving the following files...'.format(loci+1, yeari+1))
      loggerM.info('{:02d} {:02d}          {}'.format(loci+1, yeari+1, EuroAllFN))
      loggerM.info('{:02d} {:02d}          {}'.format(loci+1, yeari+1, EuroConsFN))
      loggerM.info('{:02d} {:02d}          {}'.format(loci+1, yeari+1, WeightFN))
      df_allEuros.to_csv(EuroAllFN, index=False)
      df_weights.to_csv(WeightFN, index=False)
      df_consEuros.to_csv(EuroConsFN, index=False)


    loggerM.info('{:02d}             Processing complete for location {} of {}: "{}".'.format(loci+1, loci+1, len(locations), location))
  loggerM.info('Process Complete.')
  excel.Quit()
  del(excel)

def prepareDir(outputDirP):

  outputDir = outputDirP
  logfilename = path.join(outputDir, 'extractVehProportions.log')
  new = True
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
          new = False
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
  return outputDir, logfilename, new

def parseArgs():
  desc = """
  Extract the default vehicle euroclass proportion, and the default vehicle
  weight class proportion, from the EFT file broken down by location, and year.
  """


  parser = argparse.ArgumentParser(description=desc)

  parser.add_argument('inputfile', metavar='input file',
                      type=str,
                      help=("The file to process. This should be a copy of EFT "
                            "version 7.4 or greater, and it needs a small amount "
                            "of initial setup. Under Select Pollutants select "
                            "NOx, PM10 and PM2.5. Under Traffic Format select "
                            "'Alternative Technologies'. Select 'Emission Rates "
                            "(g/km)' under 'Select Outputs', and "
                            "'Euro Compositions' under 'Advanced Options'. All "
                            "other fields should be either empty or should take "
                            "their default values."))
  parser.add_argument('outputdir', metavar='output directory',
                      type=str,
                      help=("The directory in which to save output files. If "
                            "the directory does not exist then it will be "
                            "created, assuming required permissions, etc."))
  parser.add_argument('-a', metavar='areas',
                      type=str, nargs='*', default='Scotland',
                      choices=tools.availableAreas.append('all'),
                      help=("The areas to be processed. One or more of '{}'. "
                            "Default 'Scotland'.").format("', '".join(tools.availableAreas)))
  parser.add_argument('-y', metavar='year',
                      type=int, nargs='*', default=ynow,
                      choices=range(2008, 2031),
                      help="The year or years to be processed. Default current year.")
  parser.add_argument('--keeptemp',
                      type=bool,  nargs='?', default=False,
                      help=("Whether to keep or delete temporary files. "
                            "Boolean. Default False (delete)."))
  parser.add_argument('--loggingmode',
                      nargs='?', default='INFO',
                      choices=['INFO', 'DEBUG'],
                      help=("The logging mode. Either INFO or DEBUG, default INFO."))
  return parser.parse_args()



def main():
  global logger

  # Parse the input arguments.
  pargs = parseArgs()

  # Get the assigned directory and prepare temporary directories and log files.
  outputDir, logfilename, new = prepareDir(pargs.outputdir)

  if not new:
    # See if the input arguments are identical to the last time this directory
    # was processed.
    tools.compareArgsEqual(pargs, logfilename)

  # Create the log file.
  inString = ' '.join(sys.argv)
  logger = tools.prepareLogger(__name__, logfilename, pargs, inString)

  # Read the log file to see if any combination of location and year has been
  # completed yet.
  completed = tools.getCompletedFromLog(logfilename, mode='proportions')

  # Run the processing routine.
  extractVehProps(pargs.inputfile, pargs.outputdir, pargs.a, pargs.y,
                  keepTempFiles=pargs.keeptemp, completed=completed)



if __name__ == '__main__':
  main()