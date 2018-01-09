# -*- coding: utf-8 -*-
"""
Created on Tue Jan  9 15:22:48 2018

@author: edward.barratt
"""
import argparse

import EFT_Tools as tools

if __name__ == '__main__':
  desc = """
  Combine the csv files produced by extractEFT.py in to one large csv file.

  Files marked 'COMPLETED' within the log file (which must be present in the
  designated directory) will be combined. Other files within the directory,
  csv or otherwise, will be ignored.
  """

  parser = argparse.ArgumentParser(description=desc)

  parser.add_argument('directory', metavar='directory',
                      type=str,
                      help=("The directory containing the log file for an "
                            "extractEFT.py processing job."))

  pargs = parser.parse_args()
  combined = tools.combineFiles(pargs.directory)
  print('All files have been combined, and saved in the new file:')
  print(combined)