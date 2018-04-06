# -*- coding: utf-8 -*-
"""
Created on Fri Apr  6 14:49:52 2018

@author: edward.barratt
"""

import os
import argparse
import pandas as pd



if __name__ == '__main__':
  ProgDesc = ("Creates a vehFleetSplit file of the type used by shp2EFT using "
              "the contents of an ANPR data file.")
  ANPRDesc = ("The ANPR file should be a csv file listing all vehicles "
              "passing the ANPR counter (including double counting of vehicles "
              "that have passed more than once). There should be a column each "
              "for vehicle class, euro class, and weight class.")
  parser = argparse.ArgumentParser(description=ProgDesc)
  parser.add_argument('anprfile', type=str,
                      help="The ANPR file to be processed. "+ANPRDesc)
  parser.add_argument('--vehColumnName', metavar='vehicle class column name',
                      type=str, nargs='?', default='Vehicle11Split',
                      help="The column name for the vehicle class.")
  parser.add_argument('--weightColumnName', metavar='weight class column name',
                      type=str, nargs='?', default='WeightClass',
                      help="The column name for the vehicle weight class.")
  parser.add_argument('--euroColumnName', metavar='euro class column name',
                      type=str, nargs='?', default='EuroClass',
                      help="The column name for the vehicle euro class.")

  args = parser.parse_args()
  anprfile = args.anprfile
  colW = args.weightColumnName
  colE = args.euroColumnName

  # Check file exists.
  if not os.path.exists(anprfile):
    raise ValueError('File {} does not exist.'.format(anprfile))

  # Read the file into pandas, but only keep the Euro class and the weight class
  # columns.
  data = pd.read_csv(anprfile, encoding = "ISO-8859-1")

  colnames = list(data)
  for q in [colW, colE]:
    if q not in colnames:
      raise ValueError('Column {} does not exist in file.'.format(q))
  for col in colnames:
    if col not in [colW, colE]:
      data = data.drop(col, 1)

  print(data.head())
  # Split the Weight class column. It should be in three parts, vehicle type,




