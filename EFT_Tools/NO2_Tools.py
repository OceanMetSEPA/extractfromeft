# -*- coding: utf-8 -*-
"""
addNO2
readNO2Factors

Created on Fri Apr 20 14:43:42 2018

@author: edward.barratt
"""

import pandas as pd

def addNO2(dataframe, Factors='input/NAEI_NO2Extracted.xlsx', mode='Average'):
  """
  Function that adds NO2 emission factors to a data frame that already has NOx
  emission factors.

  The original data frame must have one column called 'NOx (g/km/veh)',
  one column called 'year', and one column called 'vehicle'.
  """
  raise Exception("I don't think this function is needed any more. Is it?")
  # Is Factors a string?
  if isinstance(Factors, str):
    Factors = readNO2Factors(Factors, mode=mode)

  if mode == 'Average':
    dataframe['NO2 (g/km/veh)'] = dataframe['NOx (g/km/veh)']*Factors
  elif mode == 'ByYear':
    dataframe['NO2 (g/km/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/veh)']*Factors[row['year']], axis=1)
  elif mode == 'ByUrban':
    UNU = ['NotUrban', 'Urban']
    dataframe['NO2 (g/km/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/veh)']*Factors[UNU[row['Urban']]][row['year']], axis=1)
  elif mode == 'ByFuel':
    dataframe['NO2 (g/km/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/veh)']*Factors[row['Fuel']][row['vehicle']][row['year']], axis=1)
  elif mode == 'ByRoadType':
    dataframe['NO2 (g/km/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/veh)']*Factors[row['type']][row['vehicle']][   max([2013,row['year']])], axis=1)
  elif mode == 'ByEuro':
    dataframe['NO2 (g/km/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/veh)']*Factors[row['vehicle']][row['euro']], axis=1)
  else:
    raise ValueError("Mode '{}' not understood.".format(mode))
  return dataframe


def readNO2Factors(FactorFile='input/NAEI_NO2Extracted.xlsx', mode='Average'):
  """
  Function that reads the contents of the NOxFactorFile.

  The NOxFactorFile contains the values within the the
  PrimaryNO2_factors_NAEIBase_2016_v1.xlsx spreadsheet which is downloaded from
  http://naei.beis.gov.uk/data/ef-transport. It has been neccesary to
  reorganise the structure of that file.

  INPUTS
  FactorFile - string - the path to the NOxFactorFile.
  mode       - string - one of  'ByEuro', 'ByRoadType', 'ByFuel', 'ByUrban',
               'ByArea', 'ByYear', 'Average'.
               The appropriate mode depends on what data you have to use:
               - ByEuro should be the prefered mode. NOx-NO2 factors will be
                 split by vehicle (8 split*) and euro class. All other modes
                 are based on the euro split with assumed auro proportions.
               - ByRoadType splits the factors by vehicle (6 split*), road type
                 (Urban, Rural or Motorway), and year (2013 - 2035).
               - ByFuel splits the factors by vehicle (6 split*), fuel (Petrol
                 or Diesel), and year (2005 - 2035).
               - ByUrban splits the factors by vehicle (6 split*), urban
                 status ('Urban' or 'NotUrban'), and year (2013 - 2035).
               - ByYear splits the factors by year alone (2013 - 2035).
               - Average gives one value alone, the average of all 'ByYear'
                 factors.

  * 8 split is Petrol cars, Diesel cars, Petrol LGVs, Diesel LGV, Rigid HGV,
      Artic HGVs, Buses and coaches, Motorcycles.
    6 split is Cars, LGVs, Rigid HGV, Artic HGV, Buses and coaches, Motorcycles.

  OUTPUTS
  Factors - the NOx to NO2 conversion factors, as a dictionary.
  """
  sheets = {'ByEuro': 'By Euro',
            'ByRoadType': 'By Road Type',
            'ByFuel': 'By Fuel',
            'ByUrban': 'By Urban',
            'ByYear': 'By Year',
            'Average': 'Average'}
  allowedModes = sheets.keys()
  if mode not in allowedModes:
    raise ValueError("Mode '{}' not understood, must be one of '{}'.".format(mode, "', '".join(allowedModes)))

  FactorsDF = pd.read_excel(FactorFile, sheet_name=sheets[mode])
  if mode == 'Average':
    Factors = list(FactorsDF)[0]
  elif mode == 'ByYear':
    Factors = {}
    Years = list(FactorsDF)
    for Y in Years:
      Factors[Y] = FactorsDF[Y][0]
  elif mode == 'ByUrban':
    Factors = {'Urban': {}, 'NotUrban': {}}
    Years = list(FactorsDF)
    Years.remove('Urban')
    for Y in Years:
      Factors['NotUrban'][Y] = FactorsDF[Y][0]
      Factors['Urban'][Y] = FactorsDF[Y][1]
  elif mode == 'ByFuel':
    Factors = {}
    Years = list(FactorsDF)
    Years.remove('Vehicle')
    Years.remove('Fuel')
    Vehicles = FactorsDF['Vehicle'].unique()
    Fuels = FactorsDF['Fuel'].unique()
    for Fuel in Fuels:
      Factors[Fuel] = {}
      FFs = FactorsDF[FactorsDF['Fuel'] == Fuel]
      for Vehicle in Vehicles:
        Factors[Fuel][Vehicle] = {}
        FVs = FFs[FFs['Vehicle'] == Vehicle]
        for Y in Years:
          Factors[Fuel][Vehicle][Y] = list(FVs[Y])[0]
  elif mode == 'ByRoadType':
    Factors = {}
    Years = list(FactorsDF)
    Years.remove('Vehicle')
    Years.remove('Road Type')
    Vehicles = FactorsDF['Vehicle'].unique()
    RoadTypes = FactorsDF['Road Type'].unique()
    for RT in RoadTypes:
      RT2 = '{} (not London)'.format(RT)
      Factors[RT] = {}
      Factors[RT2] = {}
      FFs = FactorsDF[FactorsDF['Road Type'] == RT]
      for Vehicle in Vehicles:
        Factors[RT][Vehicle] = {}
        Factors[RT2][Vehicle] = {}
        FVs = FFs[FFs['Vehicle'] == Vehicle]
        for Y in Years:
          Factors[RT][Vehicle][Y] = list(FVs[Y])[0]
          Factors[RT2][Vehicle][Y] = list(FVs[Y])[0]
  elif mode == 'ByEuro':
    Factors = {}
    Euros = list(FactorsDF)
    Euros.remove('Vehicle')
    Vehicles = FactorsDF['Vehicle'].unique()
    for Vehicle in Vehicles:
      FVs = FactorsDF[FactorsDF['Vehicle'] == Vehicle]
      Factors[Vehicle] = {}
      for E in Euros:
        Factors[Vehicle][E] = list(FVs[E])[0]
  else:
    raise ValueError("Mode '{}' not understood.".format(mode))

  return Factors