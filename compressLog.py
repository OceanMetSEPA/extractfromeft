# -*- coding: utf-8 -*-
"""
Created on Tue Jan  9 15:22:48 2018

@author: edward.barratt
"""
import argparse

import EFT_Tools as tools

if __name__ == '__main__':
  desc = """
  Removes all unneccesary lines from the extractEFT log file within the
  selected directory, leaving only lines about 'COMPLETED' or 'SKIPPED' files,
  since these lines are key lines used by other processes.

  Renames the original log file, so other information is not deiscarded
  completely.

  """

  parser = argparse.ArgumentParser(description=desc)

  parser.add_argument('directory', metavar='directory',
                      type=str,
                      help=("The directory containing the log file for an "
                            "extractEFT.py processing job."))

  pargs = parser.parse_args()
  tools.compressLog(pargs.directory)