#from os import path
import os
from os import path
import datetime, time
import numpy as np
import pandas as pd
import random
import string
import subprocess
import win32com.client as win32
import pywintypes
#from fuzzywuzzy import process as fuzzyprocess

#homeDir = path.expanduser("~")

# Define some global variables. These may need to be augmented if a new EFT
# version is released.
workingDir = os.getcwd()
ahk_exepath = 'C:\Program Files\AutoHotkey\AutoHotkey.exe'
ahk_ahkpath = 'closeWarning.ahk'

versionDetails = {}
versionDetails[7.4] = {}
versionDetails[7.4]['vehRowStarts'] = [69, 79, 91, 101, 114, 130, 146, 161]
versionDetails[7.4]['vehRowEnds']   = [76, 87, 98, 109, 125, 141, 157, 172]
versionDetails[7.4]['vehRowStartsMC'] = [177, 183, 189, 195, 201, 207]
versionDetails[7.4]['vehRowEndsMC']   = [182, 188, 194, 200, 206, 212]
versionDetails[7.4]['busCoachRow']   = [429, 430]
versionDetails[7.4]['SourceNameName'] = 'Source Name'
versionDetails[7.4]['AllLDVName'] = 'All LDVs (g/km/s)'
versionDetails[7.4]['AllHDVName'] = 'All HDVs (g/km/s)'
versionDetails[7.4]['AllVehName'] = 'All Vehicles (g/km/s)'
versionDetails[7.4]['PolName'] = 'Pollutant Name'
versionDetails[7.0] = {}
versionDetails[7.0]['vehRowStarts'] = [69, 79, 100, 110, 123, 139, 155, 170]
versionDetails[7.0]['vehRowEnds']   = [75, 87, 106, 119, 134, 150, 166, 181]
versionDetails[7.0]['vehRowStartsMC'] = [186, 192, 198, 204, 210, 216]
versionDetails[7.0]['vehRowEndsMC']   = [191, 197, 203, 209, 215, 221]
versionDetails[7.0]['busCoachRow']   = [482, 483]
versionDetails[7.0]['SourceNameName'] = 'Source Name'
versionDetails[7.0]['AllLDVName'] = 'All LDVs (g/km/s)'
versionDetails[7.0]['AllHDVName'] = 'All HDVs (g/km/s)'
versionDetails[7.0]['AllVehName'] = 'All Vehicles (g/km/s)'
versionDetails[7.0]['PolName'] = 'Pollutant Name'
versionDetails[6.0] = {}
versionDetails[6.0]['vehRowStarts'] = [69, 79, 100, 110, 123, 139, 155, 170]
versionDetails[6.0]['vehRowEnds']   = [75, 87, 106, 119, 134, 150, 166, 181]
versionDetails[6.0]['vehRowStartsMC'] = [186, 192, 198, 204, 210, 216]
versionDetails[6.0]['vehRowEndsMC']   = [191, 197, 203, 209, 215, 221]
versionDetails[6.0]['busCoachRow']   = [482, 483]
versionDetails[6.0]['SourceNameName'] = 'Source_Name'
versionDetails[6.0]['AllLDVName'] = 'All LDV (g/km/s)'
versionDetails[6.0]['AllHDVName'] = 'All HDV (g/km/s)'
versionDetails[6.0]['AllVehName'] = 'All Vehicle (g/km/s)'
versionDetails[6.0]['PolName'] = 'Pollutant_Name'

availableVersions = versionDetails.keys()
availableAreas = ['England (not London)', 'Northern Ireland', 'Scotland', 'Wales']
availableRoadTypes = ['Urban (not London)', 'Rural (not London)', 'Motorway (not London)']
availableModes = ['ExtractAll', 'ExtractCarRatio', 'ExtractBus']
availableEuros = [0,1,2,3,4,5,6]

euroClassNameVariations = dict()
euroClassNameVariations[0] = ['1Pre-Euro 1', '1Pre-Euro I', '1_Pre-Euro 1', '2Pre-Euro 1',
          '4Pre-Euro 1', '5Pre-Euro 1', '6Pre-Euro 1', '7Pre-Euro 1',
          '1_Pre-Euro 1']
euroClassNameVariations[1] = ['2Euro 1', '2Euro I', '1Euro 1', '2Euro 1', '2Euro 1',
          '4Euro 1', '5Euro 1', '6Euro 1', '7Euro 1', '9 Euro I DPFRF',
          '8Euro 1 DPFRF', '9Euro I DPFRF']
euroClassNameVariations[2] = ['3Euro 2', '3Euro II', '1Euro 2', '2Euro 2', '2Euro 2',
          '4Euro 2', '5Euro 2', '6Euro 2', '7Euro 2', '10 Euro II DPFRF',
          '9Euro II SCRRF', '9Euro 2 DPFRF']
euroClassNameVariations[3] = ['4Euro 3', '4Euro III', '1Euro 3', '2Euro 3', '2Euro 3',
          '4Euro 3', '5Euro 3', '6Euro 3', '7Euro 3', '11 Euro III DPFRF',
          '10Euro III SCRRF', '8Euro 3 DPF', '10Euro 3 DPFRF']
euroClassNameVariations[4] = ['5Euro 4', '5Euro IV', '1Euro 4', '2Euro 4', '2Euro 4',
          '4Euro 4', '5Euro 4', '6Euro 4', '7Euro 4', '12 Euro IV DPFRF',
          '11Euro IV SCRRF', '9Euro 4 DPF']
euroClassNameVariations[5] = ['6Euro 5', '6Euro V', '1Euro 5', '2Euro 5', '2Euro 5',
          '4Euro 5', '5Euro 5', '6Euro 5', '7Euro 5', '7Euro V_SCR',
          '6Euro V_EGR', '12Euro V EGR + SCRRF']
euroClassNameVariations[6] = ['7Euro 6', '6Euro VI', '1Euro 6', '2Euro 6', '2Euro 6',
          '4Euro 6', '5Euro 6', '6Euro 6', '7Euro 6', '8Euro VI',
          '7Euro 6c', '7Euro 6d']

euroClassNameVariationsAll = euroClassNameVariations[0][:]
for ei in range(1,7):
  euroClassNameVariationsAll.extend(euroClassNameVariations[ei])
euroClassNameVariationsAll = list(set(euroClassNameVariationsAll))

EuroClassNameColumns = ["A", "H"]
DefaultEuroColumns = ["B", "I"]
UserDefinedEuroColumns = ["D", "K"]
EuroClassNameColumnsMC = ["B", "H"]
DefaultEuroColumnsMC = ["C", "I"]
UserDefinedBusColumn = ["D"]
UserDefinedBusMWColumn = ["E"]
DefaultBusColumn = ["B"]
DefaultBusMWColumn = ["C"]

def splitSourceNameS(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  row['vehicle'] = v
  return int(s[1:])

def splitSourceNameV(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  return v

def splitSourceNameT(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  return t

def readNO2Factors(FactorFile='input/NO2Extracted.xlsx', mode='Average'):
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

  FactorsDF = pd.read_excel(FactorFile, sheetname=sheets[mode])
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
  #def readNO2Factors

"""
PERHAPS NOT NECESSARY, MORE COMPLICATED NOW THAT READno2fACTORS HAS MULTIPLE MODES.
def getNO2Factor(Factors, Vehicle, Year):
  #
  Function to extract the NOx to NO2 conversion factor from the NOxFactorFile.

  This is neccesary because EFT does not offer NO2 emission rates. NAEI does,
  but all it does is apply a particular factor to the NAEI NOx emission rates,
  which are sourced from EFT anyway. So this method should be identical.

  The NOxFactorFile contains the values within the the
  PrimaryNO2_factors_NAEIBase_2016_v1.xlsx spreadsheet which is downloaded from
  http://naei.beis.gov.uk/data/ef-transport. It has been neccesary to
  reorganise the structure of that file.

  INPUTS
  Factors - string - the path to the NOxFactorFile,
            or a dictionary - the contents of that file (faster).
  Vehicle - string - One of Petrol cars, Diesel cars, Petrol LGVs, Diesel LGVs,
            Rigid HGVs, Artic HGVs, Buses and coaches or Motorcycles. If a
            different string is supplied then fuzzy matching will attempt to
            find the closest match.
  Year    - integer - a value between 2005 and 2035

  OUTPUTS
  Factor  - the NOx to NO2 conversion factor for the specified vehicle and year.
  #
  if Vehicle in ['ElectricCars', 'ElectricLGVs']:
    return 1 # Well it was either that or nothing. NOx emission rates are basically 0 anyway.

  # Is Factors a string?
  if isinstance(Factors, str):
    Factors = readNO2Factors(Factors)

  if Vehicle not in Factors.keys():
    VVV = fuzzyprocess.extractOne(Vehicle, Factors.keys())
    if VVV[1] < 95:
      print("No exact match for Vehicle '{}'. Using '{}'.".format(Vehicle, VVV[0]))
    Vehicle = VVV[0]

  return Factors[Vehicle][Year]
"""


def addNO2(dataframe, Factors='input/NO2Extracted.xlsx', mode='Average'):
  """
  Function that adds NO2 emission factors to a data frame that already has NOx
  emission factors.

  The original data frame must have one column called 'NOx (g/km/s/veh)',
  one column called 'year', and one column called 'vehicle'.
  """

  # Is Factors a string?
  if isinstance(Factors, str):
    Factors = readNO2Factors(Factors, mode=mode)

  if mode == 'Average':
    dataframe['NO2 (g/km/s/veh)'] = dataframe['NOx (g/km/s/veh)']*Factors
  elif mode == 'ByYear':
    dataframe['NO2 (g/km/s/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/s/veh)']*Factors[row['year']], axis=1)
  elif mode == 'ByUrban':
    UNU = ['NotUrban', 'Urban']
    dataframe['NO2 (g/km/s/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/s/veh)']*Factors[UNU[row['Urban']]][row['year']], axis=1)
  elif mode == 'ByFuel':
    dataframe['NO2 (g/km/s/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/s/veh)']*Factors[row['Fuel']][row['vehicle']][row['year']], axis=1)
  elif mode == 'ByRoadType':
    dataframe['NO2 (g/km/s/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/s/veh)']*Factors[row['type']][row['vehicle']][   max([2013,row['year']])], axis=1)
  elif mode == 'ByEuro':
    dataframe['NO2 (g/km/s/veh)'] = dataframe.apply(lambda row: row['NOx (g/km/s/veh)']*Factors[row['vehicle']][row['euro']], axis=1)
  else:
    raise ValueError("Mode '{}' not understood.".format(mode))
  return dataframe


def getInputFile(version, directory='input'):
  """
  Return the absolute path to the appropriate file for the selected mode and
  version. Will return an error if no file is available.
  """

  # First check that the directory exists.
  if not path.isdir(directory):
    raise ValueError('Cannot find directory {}.'.format(directory))

  # Now figure out the file name.
  if version == 6.0:
    vPart = 'EFT2014_v6.0.2'
    ext = '.xls'
  elif version == 7.0:
    vPart = 'EFT2016_v7.0'
    ext = '.xlsb'
  elif version == 7.4:
    vPart = 'EFT2017_v7.4'
    ext = '.xlsb'
  else:
    raise ValueError('Version {} is not recognised.'.format(version))

  fname = '{}/{}_empty{}'.format(directory, vPart, ext)
  # return the absolute paths.
  fname =  path.abspath(fname)

  # Check that file exists.
  if not path.exists(fname):
    raise ValueError('Cannot find file {}.'.format(fname))

  return fname

def secondsToString(seconds, form='short'):
  td = datetime.timedelta(seconds=seconds)
  if form == 'short':
    return str(td)
  elif form == 'long':
    d = td.days
    s = td.seconds
    m, s = divmod(s, 60)
    h, m = divmod(m, 60)
    if m == 0:
      return '{} seconds'.format(s)
    elif h == 0:
      return '{} minutes and {} seconds'.format(m, s)
    elif d == 0:
      return '{} hours, {} minutes and {} seconds'.format(h, m, s)
    else:
      return '{} days, {}hours, {} minutes and {} seconds'.format(d, h, m, s)
  else:
    raise ValueError("Format '{}' is not understood.".format(form))

def romanNumeral(N):
  # Could write a function that deals with any, but I only need up to 10.
  RNs = [0, 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
  return RNs[N]

def extractVersion(fileName, availableVersions=[6.0, 7.0, 7.4]):
  """
  Extract the version number from the filename.
  """
  # See what version we're looking at.
  version = False
  for versiono in availableVersions:
    if fileName.find('v{:.1f}'.format(versiono)) >= 0:
      version = versiono
      version_for_output = versiono
      break
  if version:
    print('{} is EFT of version {}.'.format(fileName, version))
  else:
    # Not one that is predefined, see if we can get the version number.
    fv = fileName.find('v')
    fp = fileName.find('_prefilled')
    if (fv >= 0) and (fp >= 0):
      fl = fileName.find('.', fv, fp)
      if fl >= 0:
        verTry = fileName[fv+1:fl+2]
        try:
          version = float(verTry)
        except:
          pass
    if version:
      # Get closest version number
      versioncloseI = np.argmin(abs(np.array(availableVersions) - version))
      version_for_output = version
      versionp = availableVersions[versioncloseI]
      print('{} looks like EFT of unknown version {}, will process as version {}.'.format(fileName, version, versionp))
      version = versionp
    else:
      maxAvailableVersions = max(availableVersions)
      print('Cannot parse version number from "{}", will attempt to process as version {}.'.format(fileName, maxAvailableVersions))
      version = maxAvailableVersions
      version_for_output = 'Unknown Version as {}'.format(maxAvailableVersions)
    print('You may wish to edit the versionDetails global variables to account for the new version.')
  return version, version_for_output

def randomString(N = 10):
  return ''.join(random.choice(string.ascii_uppercase + string.ascii_lowercase + string.digits) for x in range(N))

def numToLetter(N, ABC=u'ABCDEFGHIJKLMNOPQRSTUVWXYZ' ):
  """
  Converts a number in to a letter, or set of letters, based on the contents of
  ABC. If ABC contains the alphabet (which is the default behaviour) then for a
  number n this will result in the same string as the name of the nth column in
  a normal spreadsheet.

  INPUTS
  N    - integer - any number.
  OPTIONAL INPUTS
  ABC  - string  - any string of characters (or any other symbols I believe).
  OUTPUTS
  L    - string  - the number converted to a string.

  EXAMPLES
  numToLetter(1)
  u'A'
  numToLetter(26)
  u'Z'
  numToLetter(27)
  u'AA'
  numToLetter(345)
  u'MG'
  numToLetter(345, ABC='1234567890')
  u'345'
  """
  L = ''
  sA = len(ABC)
  div = int(N)
  while div > 0:
    R = int((div-1)%sA)
    L = ABC[R] + L
    div = int((div-R)/sA)
  return L


def createEFTInput(vBreakdown='Detailed Option 2',
                   speeds=[5,6,7,8,9,10,12,14,16,18,20,25,30,35,40,
                           45,50,60,70,80,90,100,110,120,130,140],
                   roadTypes=availableRoadTypes):

  VehSplits = {'Basic Split': ['HDV'],
               'Detailed Option 1': ['Car', 'Taxi (black cab)', 'LGV', 'HGV',
                                     'Bus and Coach', 'Motorcycle'],
               'Detailed Option 2': ['Car', 'Taxi (black cab)', 'LGV', 'Rigid HGV',
                                     'Artic HGV', 'Bus and Coach', 'Motorcycle'],
               'Detailed Option 3': ['Petrol Car', 'Diesel Car', 'Taxi (black cab)',
                                     'LGV', 'Rigid HGV', 'Artic HGV',
                                     'Bus and Coach', 'Motorcycle']}

  VehSplit = VehSplits[vBreakdown]
  #RoadTypes = ['Urban (not London)', 'Rural (not London)', 'Motorway (not London)']

  if type(roadTypes) is str:
    if roadTypes in ['all', 'All', 'ALL']:
      roadTypes = availableRoadTypes
    else:
      roadTypes = [roadTypes]

  if vBreakdown == 'Basic Split':
    numRows = 2*len(roadTypes)*len(speeds)
  else:
    numRows = len(roadTypes)*len(speeds)*(len(VehSplit)-1)
  numCols = 6 + len(VehSplit)
  inputDF = pd.DataFrame(index=range(numRows), columns=range(numCols))
  ri = -1
  for rT in roadTypes:
    for sp in speeds:
      for veh in VehSplit:
        if vBreakdown == 'Basic Split':
          ri += 2
          inputDF.set_value(ri-1, 0, 'S{} - LDV - {}'.format(sp, rT))
          inputDF.set_value(ri-1, 1, rT)
          inputDF.set_value(ri-1, 2, 1)
          inputDF.set_value(ri-1, 3, 0)
          inputDF.set_value(ri-1, 4, sp)
          inputDF.set_value(ri-1, 5, 1)
          inputDF.set_value(ri-1, 6, 1)
          inputDF.set_value(ri, 0, 'S{} - HDV - {}'.format(sp, rT))
          inputDF.set_value(ri, 1, rT)
          inputDF.set_value(ri, 2, 1)
          inputDF.set_value(ri, 3, 100)
          inputDF.set_value(ri, 4, sp)
          inputDF.set_value(ri, 5, 1)
          inputDF.set_value(ri, 6, 1)
        else:
          if veh == 'Taxi (black cab)':
            pass
          else:
            ri += 1
            inputDF.set_value(ri, 0, 'S{} - {} - {}'.format(sp, veh, rT))
            inputDF.set_value(ri, 1, rT)
            inputDF.set_value(ri, 2, 1)
            for vehi, vehb in enumerate(VehSplit):
              if vehb == veh:
                inputDF.set_value(ri, 3+vehi, 100)
              else:
                inputDF.set_value(ri, 3+vehi, 0)
            inputDF.set_value(ri, len(VehSplit)+3, sp)
            inputDF.set_value(ri, len(VehSplit)+4, 1)
            inputDF.set_value(ri, len(VehSplit)+5, 1)

  return inputDF

def checkEuroClassesValid(workBook, vehRowStarts, vehRowEnds, EuroClassNameColumns, MC=0):
  """
  Check that all of the available euro classes are specified.
  """
  if MC == 1:
    print("      Checking all motorcycle euro class names are understood.")
  elif MC == -1:
    print("      Checking all non-motorcycle euro class names are understood.")
  else:
    print("      Checking all euro class names are understood.")
  ws_euro = workBook.Worksheets("UserEuro")
  for [vi, vehRowStart] in enumerate(vehRowStarts):
    vehRowEnd = vehRowEnds[vi]
    for [ci, euroNameCol] in enumerate(EuroClassNameColumns):
      euroClassRange = "{col}{rstart}:{col}{rend}".format(col=euroNameCol, rstart=vehRowStart, rend=vehRowEnd)
      euroClassesAvailable = ws_euro.Range(euroClassRange).Value

      for ecn in euroClassesAvailable:
        ecn = ecn[0]
        if ecn is None:
          continue
        if ecn not in euroClassNameVariationsAll:
          raise ValueError('Unrecognized Euro Class Name: "{}".'.format(ecn))
  print("        All understood.")

def euroSearchTerms(N):
  ES = euroClassNameVariations[N]
  return ES

def specifyEuroProportions(euroClass, workBook, vehRowStarts, vehRowEnds,
                 EuroClassNameColumns, DefaultEuroColumns, UserDefinedEuroColumns, MC=False):
  """
  Specify the euro class proportions.
  Will return the defualt proportions.
  """
  defaultProps = {}
  #print("    Setting euro ratios to 100% for euro {}.".format(euroClass))
  ws_euro = workBook.Worksheets("UserEuro")
  for [vi, vehRowStart] in enumerate(vehRowStarts):
    if MC:
      vehNameA = ws_euro.Range("A{row}".format(row=vehRowStart)).Value
      vehNameB = ws_euro.Range("A{row}".format(row=vehRowStart+1)).Value
      if vehNameB is None:
        vehName = 'Motorcycle - {}'.format(vehNameA)
      else:
        vehName = 'Motorcycle - {} - {}'.format(vehNameA, vehNameB)
    else:
      vehName = ws_euro.Range("A{row}".format(row=vehRowStart-1)).Value
    #print("      Setting euro ratios for {}.".format(vehName))
    vehRowEnd = vehRowEnds[vi]
    for [ci, euroNameCol] in enumerate(EuroClassNameColumns):
      userDefinedCol = UserDefinedEuroColumns[ci]
      defaultEuroCol = DefaultEuroColumns[ci]
      euroClassRange = "{col}{rstart}:{col}{rend}".format(col=euroNameCol, rstart=vehRowStart, rend=vehRowEnd)
      euroClassesAvailable = ws_euro.Range(euroClassRange).Value
      # Make sure we don't include trailing 'None' rows.
      euroClassesAvailableR = list(reversed(euroClassesAvailable))
      for eca in euroClassesAvailableR:
        #print(eca)
        if eca[0] is None:
          vehRowEnd = vehRowEnd - 1
        else:
          break
      # See which columns contain a line that specifies the required euro class.
      rowsToDo = []
      euroClass_ = euroClass
      while len(rowsToDo) == 0:
        got = False
        euroSearchTerms_ = euroSearchTerms(euroClass_)
        for [ei, name] in enumerate(euroClassesAvailable):
          name = name[0]
          if name in euroSearchTerms_:
            rowsToDo.append(vehRowStart + ei)
            got = True
        if not got:
          # print('      No values available for euro {}, trying euro {}.'.format(euroClass_, euroClass_-1))
          euroClass_ -= 1
      ignoreForPropRecord = False
      if euroClass_ != euroClass:
        ignoreForPropRecord = True
      # Get the default proportions.
      defaultProportions = []
      for row in rowsToDo:
        propRange = "{col}{row}".format(col=defaultEuroCol, row=row)
        defaultProportion = ws_euro.Range(propRange).Value
        defaultProportions.append(defaultProportion)
        #print(propRange)
        #print(defaultProportions)
      defaultProportions = np.array(defaultProportions)
      if ci == 0:
        if ignoreForPropRecord:
          #print('        Default proportions taken as 0.00%.')
          defaultProps[vehName] = 0
        else:
          #print('        Default proportions are {:.2f}%.'.format(100*sum(defaultProportions)))
          defaultProps[vehName] = 100*sum(defaultProportions)
      # Normalize them.
      if sum(defaultProportions) < 0.00001:
        defaultProportions = defaultProportions + 1
      userProportions = defaultProportions/sum(defaultProportions)
      # And set the values in the sheet.
      # Set all to zero first.
      userRange = "{col}{rstart}:{col}{rend}".format(col=userDefinedCol, rstart=vehRowStart, rend=vehRowEnd)
      ws_euro.Range(userRange).Value = 0
      # Then set the specific values.
      #print(rowsToDo)
      for [ri, row] in enumerate(rowsToDo):
        userRange = "{col}{row}".format(col=userDefinedCol, row=row)
        value = userProportions[ri]
        ws_euro.Range(userRange).Value = value
  #print('    All complete')
  return defaultProps

def specifyBusCoach(wb, busCoach, busCoachRow, UserDefinedBusColumn,
                    UserDefinedBusMWColumn, DefaultBusColumn, DefaultBusMWColumn):
  defaultBusProps = {}
  ws_euro = wb.Worksheets("UserEuro")
  defaultBusProps['bus_non_mw'] = ws_euro.Range("{}{}".format(DefaultBusColumn[0], busCoachRow[0])).Value
  defaultBusProps['coach_non_mw'] = ws_euro.Range("{}{}".format(DefaultBusColumn[0], busCoachRow[1])).Value
  defaultBusProps['bus_mw'] = ws_euro.Range("{}{}".format(DefaultBusMWColumn[0], busCoachRow[0])).Value
  defaultBusProps['coach_mw'] = ws_euro.Range("{}{}".format(DefaultBusMWColumn[0], busCoachRow[1])).Value

  if busCoach != 'default':
    if busCoach == 'bus':
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[0])).Value = 1
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[1])).Value = 0
      try:
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[0])).Value = 1
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[1])).Value = 0
      except pywintypes.com_error:
        # Doesn't work in version 6, it doesn't let you specify the motorway proportion.
        pass
    elif busCoach == 'coach':
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[0])).Value = 0
      ws_euro.Range("{}{}".format(UserDefinedBusColumn[0], busCoachRow[1])).Value = 1
      try:
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[0])).Value = 0
        ws_euro.Range("{}{}".format(UserDefinedBusMWColumn[0], busCoachRow[1])).Value = 1
      except pywintypes.com_error:
        pass
    else:
      raise ValueError("busCoach should be either 'bus' or 'coach'. '{}' is not understood.".format(busCoach))
  return defaultBusProps

def prepareToExtract(fileNames, locations):
  """
  Extract the pre-processing information from the filenames and locations.
  """
  # Make sure location is a list that can be iterated through.
  #if type(locations) is str:
  #  locations = [locations]
  # Make sure fileNames is a list that can be iterated through.
  if type(fileNames) is str:
    fileNames = [fileNames]

  # Check that the auto hot key executable, and control file, are available.
  if not path.isfile(ahk_exepath):
    raise ValueError('The Autohotkey executable file {} could not be found.'.format(ahk_exepath))
  if not path.isfile(ahk_ahkpath):
    ahk_ahkpath_ = workingDir + '\\' + ahk_ahkpath
    if not path.isfile(ahk_ahkpath_):
      raise ValueError('The Autohotkey file {} could not be found.'.format(ahk_ahkpath))
    else:
      ahk_ahkpathGot = ahk_ahkpath_
  else:
    ahk_ahkpathGot = ahk_ahkpath

  versionNos = []
  versionForOutputs = []
  for fNi, fN in enumerate(fileNames):
    # Extract the version number.
    version, versionForOutput = extractVersion(fN)
    versionNos.append(version)
    versionForOutputs.append(versionForOutput)

    # Get the absolute path to the file. The excel win32 stuff doesn't seem to
    # work with relative paths.
    fN_ = path.abspath(fN)
    if not path.isfile(fN):
      raise ValueError('Could not find {}.'.format(fN))
    fileNames[fNi] = fN_

  return ahk_ahkpathGot, fileNames, versionNos, versionForOutputs

def extractOutput(fileName, versionForOutPut, year, location, euroClass, details):
  ex = pd.ExcelFile(fileName)
  output = ex.parse("Output")
  # Add some other columns to the dataframe.
  output['version'] = versionForOutPut
  output['year'] = year
  output['area'] = location
  output['type'] = output.apply(splitSourceNameT, SourceName=details['SourceNameName'], axis=1)
  output['vehicle'] = output.apply(splitSourceNameV, SourceName=details['SourceNameName'], axis=1)
  output['euro'] = euroClass
  output['speed'] = output.apply(splitSourceNameS, SourceName=details['SourceNameName'], axis=1)
  # Drop columns that are not required anymore.
  output = output.drop(details['SourceNameName'], 1)
  output = output.drop(details['AllLDVName'], 1)
  output = output.drop(details['AllHDVName'], 1)
  # Pivot the table so each pollutant has a column.
  pollutants = list(output[details['PolName']].unique())
  # Rename, because after the pivot the 'column' name will become the
  # index name.
  output = output.rename(columns={details['PolName']: 'RowIndex'})
  output = output.pivot_table(index=['year', 'area', 'euro', 'version',
                                     'speed', 'vehicle', 'type'],
                                    columns='RowIndex',
                                    values=details['AllVehName'])
  output = output.reset_index()

  renames = {}
  # Rename the pollutant columns to include the units.
  for Pol in pollutants:
    if Pol == 'PM25':
      Pol_ = 'PM2.5'
    else:
      Pol_ = Pol
    renames[Pol] = '{} (g/km/s/veh)'.format(Pol_)
  output = output.rename(columns=renames)
  return output


def runAndExtract(fileName, location, year, euroClass, ahk_exepath,
                  ahk_ahkpathG, vehSplit, details, versionForOutPut, excel=None,
                  checkEuroClasses=False, DoMCycles=True, DoBusCoach=False,
                  inputData='prepare', busCoach='default'):
  """
  Prepare the file for running the macro.
  euroClass of -9 will retain default euro breakdown.
  """
  closeExcel = False
  if excel is None:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    closeExcel = True

  # Start off the autohotkey script as a (parallel) subprocess. This will
  # continually check until the compatibility warning appears, and then
  # close the warning.
  subprocess.Popen([ahk_exepath, ahk_ahkpathG])

  # Open the document.
  wb = excel.Workbooks.Open(fileName)
  excel.Visible = True

  if checkEuroClasses:
    # Check that all of the euro class names within the document are as
    # we would expect. An error will be raised if there are any surprises
    # and this will mean that the global variables at the start of the
    # code will need to be edited.
    if DoMCycles:
      checkEuroClassesValid(wb, details['vehRowStartsMC'], details['vehRowEndsMC'], EuroClassNameColumnsMC, MC=1)
    checkEuroClassesValid(wb, details['vehRowStarts'], details['vehRowEnds'], EuroClassNameColumns, MC=-1)

  # Set the default values in the Input Data sheet.
  ws_input = wb.Worksheets("Input Data")
  ws_input.Range("B4").Value = location
  ws_input.Range("B5").Value = year
  ws_input.Range("B6").Value = vehSplit

  if type(inputData) is str:
    if inputData == 'prepare':
      # Prepare the input data.
      inputData = createEFTInput(vBreakdown=vehSplit)
      inputData = inputData.as_matrix()
    else:
      raise ValueError("inputData '{}' is not understood.".format(inputData))

  numRows, numCols = np.shape(inputData)
  ws_input.Range("A10:{}{}".format(numToLetter(numCols), numRows+9)).Value = inputData

  # Now we need to populate the UserEuro table with the defaults. Probably
  # only need to do this once per year, per area, but will do it every time
  # just in case.
  excel.Application.Run("PasteDefaultEuroProportions")

  # Now specify that we only want the specified euro class, by turning the
  # proportions for that class to 1, (or a weighted value if there are more
  # than one row for the particular euro class). This function also reads
  # the default proportions.
  if euroClass == -9:
    # Just stick with default euroclass.
    defaultProportions = 'NotMined'
    busCoachProportions = 'NotMined'
    pass
  else:
    defaultProportions = pd.DataFrame(columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
    # Motorcycles first
    if DoMCycles:
      print('      Assigning fleet euro proportions for motorcycles.')
      defaultProportionsMC_ = specifyEuroProportions(euroClass, wb,
                                  details['vehRowStartsMC'], details['vehRowEndsMC'],
                                  EuroClassNameColumnsMC, DefaultEuroColumnsMC,
                                  UserDefinedEuroColumns, MC=True)
      for key, value in defaultProportionsMC_.items():
        defaultProportionsRow = pd.DataFrame([[year, location, key, euroClass, value]],
                                             columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
        defaultProportions = defaultProportions.append(defaultProportionsRow)
      print('      Assigning fleet euro proportions for all other vehicle types.')
    else:
      print('      Assigning fleet euro proportions for all vehicle types except motorcycles.')
    # And all other vehicles
    defaultProportions_ = specifyEuroProportions(euroClass, wb,
                             details['vehRowStarts'], details['vehRowEnds'],
                             EuroClassNameColumns, DefaultEuroColumns,
                             UserDefinedEuroColumns)
    # Organise the default proportions.
    for key, value in defaultProportions_.items():
      defaultProportionsRow = pd.DataFrame([[year, location, key, euroClass, value]],
                                           columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
      defaultProportions = defaultProportions.append(defaultProportionsRow)
    defaultProportions['version'] = versionForOutPut

    busCoachProportions = 'NotMined'
    if DoBusCoach:
      # Set the bus - coach proportions.
      busCoachProportions = specifyBusCoach(wb, busCoach, details['busCoachRow'],
                                            UserDefinedBusColumn, UserDefinedBusMWColumn,
                                            DefaultBusColumn, DefaultBusMWColumn)

  # Now run the EFT tool.
  ws_input.Select() # Select the appropriate sheet, we can't run the macro
                    # from another sheet.
  print('      Running EFT routine. Ctrl+C will pause processing at the end of the routine...')
  alreadySaved = False
  try:
    excel.Application.Run("RunEfTRoutine")
    print('        Complete. Ctrl+C will now halt entire programme as usual.')
    time.sleep(0.5)
  except KeyboardInterrupt:
    print('Process paused at {}.'.format(datetime.strftime(datetime.now(), '%H:%M:%S on %d-%m-%Y')))
    # Save and Close. Saving as an xlsm, rather than a xlsb, file, so that it
    # can be opened by pandas.
    (FN, FE) =  path.splitext(fileName)
    if DoBusCoach:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}_{}).xlsm'.format(location, year, euroClass, busCoach))
    else:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}).xlsm'.format(location, year, euroClass))
    wb.SaveAs(tempSaveName, win32.constants.xlOpenXMLWorkbookMacroEnabled)
    wb.Close()
    excel.Quit()
    alreadySaved = True
    time.sleep(1)
    raw_input('Press enter to resume.')
    print('Resumed at {}.'.format(datetime.strftime(datetime.now(), '%H:%M:%S on %d-%m-%Y')))
    excel = win32.gencache.EnsureDispatch('Excel.Application')

  if not alreadySaved:
    # Save and Close. Saving as an xlsm, rather than a xlsb, file, so that it
    # can be opened by pandas.
    (FN, FE) =  path.splitext(fileName)
    if DoBusCoach:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}_{}).xlsm'.format(location, year, euroClass, busCoach))
    else:
      tempSaveName = fileName.replace(FE, '({}_{}_E{}).xlsm'.format(location, year, euroClass))
    wb.SaveAs(tempSaveName, win32.constants.xlOpenXMLWorkbookMacroEnabled)
    wb.Close()

  time.sleep(1) # To allow all systems to catch up.
  if closeExcel:
    excel.Quit()
    del(excelObj) # Make sure it's gone. Apparently some people have found this neccesary.
  return excel, tempSaveName, defaultProportions, busCoachProportions



if __name__ == '__main__':
  # For testing.

  aa = createEFTInput(vBreakdown='Detailed Option 2')
  print(aa.head(30))
