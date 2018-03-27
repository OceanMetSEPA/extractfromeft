# -*- coding: utf-8 -*-
"""
Created on Tue Jan  9 15:22:48 2018

@author: edward.barratt
"""
import os
import argparse

#import EFT_Tools as tools

if __name__ == '__main__':
  desc = """
  Combine the csv files produced by extractVehProportions.py in to one large csv file.
  """

  parser = argparse.ArgumentParser(description=desc)

  parser.add_argument('directory', metavar='directory',
                      type=str,
                      help=("The directory containing the log file for an "
                            "extractEFT.py processing job."))

  pargs = parser.parse_args()
  directory = pargs.directory

  Types = ['AllEuroProportions', 'ConsolidatedEuroProportions', 'WeightProportions']
  fs = {}
  for ty in Types:
    fs[ty] = {}
    fs[ty]['path'] = os.path.join(directory, 'AllCombined_{}.csv'.format(ty))
    fs[ty]['f'] = open(fs[ty]['path'], 'w')
    fs[ty]['first'] = True


  files = [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f)) and (f[-4:] == '.csv') and (f[:11] != 'AllCombined')]
  print(files)
  for file in files:
    for ty in Types:
      if ty in file:
        f = open(file, 'r')
        for li, line in enumerate(f):
          if li == 0:
            if fs[ty]['first']:
              fs[ty]['first'] = False
            else:
              continue
          fs[ty]['f'].write(line)




