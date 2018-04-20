# -*- coding: utf-8 -*-
"""
Containing the following functions.
combineFiles
compareArgsEqual
compressLog
getCompletedFromLog
getLogger
logprint
prepareLogger

Created on Fri Apr 20 15:04:03 2018

@author: edward.barratt
"""

import os
import shutil
import pandas as pd
import datetime
import logging
import ast





def combineFiles(directory):
  """
  Combine the created files within a directory. It works by reading the
  directory's log file and reads those files that have been completed.
  """
  # Does the directory exist?
  if not os.path.isdir(directory):
    raise ValueError('Directory cannot be found.')
  # Search for the log file.
  contents = os.listdir(directory)
  if len(contents) == 0:
    raise ValueError('Directory is empty.')
  go = False
  for content in contents:
    if content[-4:] == '.log':
      logfilename = os.path.join(directory, content)
      yn = input(('Combine files listed as completed '
                  'in log file {}? [y/n]'.format(logfilename)))
      if yn.upper() != 'Y':
        continue
      else:
        go = True
        break

  first = True
  if go:
    [fname, ext] = os.path.splitext(logfilename)
    fnew = fname+'_combined.csv'
    completed = getCompletedFromLog(logfilename, mode='completed_file_only')
    filenames = list(completed['saveloc'])

    for fni, fn in enumerate(filenames):
      [pathOld, fName] = os.path.split(fn)
      print('File {:04d} of {:4d}: {}'.format(fni, len(filenames), fName))
      if not os.path.isfile(fn):
        # This could happen if the parent directory has been changed.
        if os.path.abspath(pathOld) == os.path.abspath(directory):
          raise ValueError('The file {} cannot be found in directory {}.'.format(fName, directory))
        else:
          fn = os.path.join(directory, fName)
          if not os.path.isfile(fn):
            #if fName == 'MULTI':
            #  continue
            raise ValueError('The file {} cannot be found in directory {} or directory {}.'.format(fName, directory, pathOld))
      if first:
        shutil.copyfile(fn, fnew)
        first = False
      else:
        df = pd.read_csv(fn)
        df.to_csv(fnew, mode='a', header=False, index=False)
  return fnew

def compareArgsEqual(newargs, logfilename):
  searchStr = 'Input arguments parsed as: '
  with open(logfilename, 'r') as f:
    for line in f:
      # We want the last set of commands.
      if searchStr in line:
        oldargs = line[line.find(searchStr)+len(searchStr):-1]
  # Check that they are equal
  oldargs = ast.literal_eval(oldargs)
  newargs = vars(newargs)

  if oldargs != newargs:
    print('')
    print(('You are attempting to continue evaluation based on a different set '
           'of input arguments:'))
    for key in oldargs.keys():
      print('Old: {}, {}'.format(key, oldargs[key]))
      print('New: {}, {}'.format(key, newargs[key]))

    Cont = input('Do you wish to continue. [y/n]')
    if Cont.lower() in ['yes', 'y']:
      pass
    else:
      exit()

def compressLog(directory):
  # Does the directory exist?
  if not os.path.isdir(directory):
    raise ValueError('Directory cannot be found.')
  # Search for the log file.
  contents = os.listdir(directory)
  if len(contents) == 0:
    raise ValueError('Directory is empty.')
  go = False
  for content in contents:
    if content[-4:] == '.log':
      logfilename = os.path.join(directory, content)
      yn = input(('Compress log file '
                  '{}? [y/n]'.format(logfilename)))
      if yn.upper() != 'Y':
        continue
      else:
        go = True
        break

  if not go:
    return

  # First save the logfile to a new name.
  [nn, ee] = os.path.splitext(logfilename)
  fold = nn + 'Z_AsOf_{}'.format(datetime.datetime.strftime(datetime.datetime.now(), '%y%m%d%H%M%S')) + ee
  shutil.move(logfilename, fold)

  KeepLines = ['COMPLETED (area, year, euro, tech',
               'SKIPPED (area, year, euro,']
  KeepLast = ['Input arguments parsed as:']
  LastLines = ['']
  with open(fold, 'r') as orig:
    with open(logfilename, 'w') as new:
      new.write('Log file compressed from {}.\r\n'.format(fold))
      for line in orig:
        for kl in KeepLines:
          if kl in line:
            new.write(line)
        for ki, kl in enumerate(KeepLast):
          if kl in line:
            LastLines[ki] = line

      for ll in LastLines:
        new.write(ll)

def getCompletedFromLog(logfilename, mode='completed'):
  """
  Read the log file to see if any combination of location, year, euroclass,
  and tech have already been completed. Returns completed parameters in a
  pandas dataframe. Can also return combinations marked as skipped.

  logfilename should be the path to a log file created by extractEFT.py
  mode can be 'completed', 'skipped', 'both' or 'completed_file_only'.
  """

  CompletedSearchStr = 'COMPLETED (area, year, euro, tech, saveloc): '
  CompletedSearchStrV2 = 'COMPLETED (area, year, euro, tech, saveloc, bus, weight): '
  SkippedSearchStr = 'SKIPPED (area, year, euro, tech, saveloc): '
  ProportionsSearchStr = 'COMPLETED (area, year): '
  if mode in ['completed', 'completed_file_only']:
    SearchStrs = [CompletedSearchStr, CompletedSearchStrV2]
  elif mode == 'skipped':
    SearchStrs = [SkippedSearchStr]
  elif mode == 'both':
    SearchStrs = [CompletedSearchStr, CompletedSearchStrV2, SkippedSearchStr]
  elif mode == 'proportions':
    SearchStrs = [ProportionsSearchStr]
  else:
    raise ValueError("mode '{}' is not understood.".format(mode))

  completed = pd.DataFrame(columns=['area', 'year', 'euro', 'tech', 'saveloc', 'busmode', 'weight'])
  ci = 0
  with open(logfilename, 'r') as logf:
    for line in logf:
      for SearchStr in SearchStrs:
        if SearchStr in line:
          ci += 1
          info = line[line.find(SearchStr)+len(SearchStr):-2]
          infosplt = info.split(',')
          if mode == 'completed_file_only':
            if infosplt[4].strip() == 'MULTI':
              continue
          try:
            completed.loc[ci] = [infosplt[0].strip(), int(infosplt[1].strip()),
                                 int(infosplt[2].strip()), infosplt[3].strip(),
                                 infosplt[4].strip(), infosplt[5].strip(),
                                 int(infosplt[6].strip())]
          except IndexError:
            completed.loc[ci] = [infosplt[0].strip(), int(infosplt[1].strip()),
                                 int(infosplt[2].strip()), infosplt[3].strip(),
                                 infosplt[4].strip(), 'NA', -9]
          break
  return completed

def getLogger(logger, modName):
  if logger is None:
    return None
  else:
    if 'EFT_Tools' in logger.name:
      return logger.getChild(modName)
    else:
      return logger.getChild('EFT_Tools.{}'.format(modName))

def logprint(logger, string, level='info'):
  if level.lower() == 'info':
    logfunc = lambda x: logger.info(x)
  elif level.lower() == 'debug':
    logfunc = lambda x: logger.debug(x)
  if logger is not None:
    logfunc(string)
  else:
    print(string)

def prepareLogger(loggerName, logfilename, pargs, inString):
  loggerName = __name__
  logger = logging.getLogger(loggerName)
  if pargs.loggingmode == 'INFO':
    logger.setLevel(logging.INFO)
  elif pargs.loggingmode == 'DEBUG':
    logger.setLevel(logging.DEBUG)
  else:
    raise ValueError("Logging mode '{}' not understood.".format(pargs.loggingmode))

  fileFormatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
  streamFormatter = logging.Formatter('%(asctime)s - %(message)s')
  logfilehandler = logging.FileHandler(logfilename)
  logfilehandler.setFormatter(fileFormatter)
  logstreamhandler = logging.StreamHandler()
  logstreamhandler.setFormatter(streamFormatter)
  logger.addHandler(logfilehandler)
  logger.addHandler(logstreamhandler)

  logger.info('Program started with command: "{}"'.format(inString))
  logger.info('Input arguments parsed as: {}'.format(vars(pargs)))
  return logger