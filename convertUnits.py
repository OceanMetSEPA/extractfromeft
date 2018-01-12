# -*- coding: utf-8 -*-
"""
Created on Tue Jan  9 15:22:48 2018

@author: edward.barratt
"""
import argparse
import os
from os import path
import pandas as pd

import EFT_Tools as tools

def do(directory):
  """
  Combine the created files within a directory. It works by reading the
  directory's log file and reads those files that have been completed.
  """

  # Does the directory exist?
  if not path.isdir(directory):
    raise ValueError('Directory cannot be found.')
  # Search for the log file.
  contents = os.listdir(directory)
  if len(contents) == 0:
    raise ValueError('Directory is empty.')
  go = False
  for content in contents:
    if content[-4:] == '.log':
      logfilename = path.join(directory, content)
      yn = input(('Alter files listed as completed '
                  'in log file {}? You cannot undo this step, '
                  'so make sure that you have made a backup '
                  'copy just in case. [y/n]'.format(logfilename)))
      if yn.upper() != 'Y':
        continue
      else:
        go = True
        break

  Pols = ['NOx', 'NO2', 'PM10', 'PM2.5']

  if go:
    [fname, ext] = os.path.splitext(logfilename)
    fnew = fname+'_combined.csv'
    completed = tools.getCompletedFromLog(logfilename)
    filenames = list(completed['saveloc'])

    for fni, fn in enumerate(filenames):
      print(fn)
      df = pd.read_csv(fn)
      colNames = df
      renames = {}
      gotAny = False
      for P in Pols:
        colNameOld = '{} (g/km/s/veh)'.format(P)
        colNameNew = '{} (g/km/veh)'.format(P)
        if colNameOld in colNames:
          gotAny = True
          print('  converting {} to {}.'.format(colNameOld, colNameNew))
          df[colNameOld] = df[colNameOld] * 3600
          renames[colNameOld] = colNameNew
      if gotAny:
        df = df.rename(columns=renames)
        df.to_csv(fn, mode='w', header=True, index=True)
      else:
        print('  no changes neccesary.')
  return fnew


if __name__ == '__main__':
  desc = """
  Change the units of the emission rate rows from g/km/s/veh to g/km/veh by
  multiplying each value bu 3600.

  Files marked 'COMPLETED' within the log file (which must be present in the
  designated directory) will be processed. Other files within the directory,
  csv or otherwise, will be ignored.
  """

  parser = argparse.ArgumentParser(description=desc)

  parser.add_argument('directory', metavar='directory',
                      type=str,
                      help=("The directory containing the log file for an "
                            "extractEFT.py processing job."))

  pargs = parser.parse_args()
  combined = do(pargs.directory)
  print('Completed')