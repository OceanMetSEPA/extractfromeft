# -*- coding: utf-8 -*-
"""
Functions that allow the traffic counts within a shape file to be fed into the
EFT. In particular it is designed to work with the Noise traffic count shape
files.
"""
from __future__ import print_function

import sys
import os
import shutil
import subprocess
import datetime
#import numpy as np
import pandas as pd
import geopandas as gpd
import time
#import shapely
#from shapely.geometry import LineString

import win32com.client as win32

# Set some defaults.
ahk_exepath = 'C:\Program Files\AutoHotkey\AutoHotkey.exe'
ahk_ahkpath = 'closeWarning.ahk'
ahk_exist = True
for ahk_ in [ahk_exepath, ahk_ahkpath]:
  if not os.path.isfile(ahk_):
    ahk_exist = False
if not ahk_exist:
  print(["Either Autohotkey is not installed, or it is not installed in the ",
         "usual place, or the required AutoHotKey script 'closeWarning.ahk' ",
         "cannot be found. the function will work but you will have to watch ",
         "and close the EFT warning dialogue manually."])

MotorwayNames = ['Motorway', 'M8', 'M74', 'M73', 'M77', 'M80', 'M876', 'M898', 'M9', 'M90']

# Vehicle subtypes for extracting NO2 factors.
VehSubTypesNO2 = {u'Petrol cars': [u'Petrol Cars (g/km/s)', u'Full Hybrid Petrol Cars (g/km/s)', u'Plug-In Hybrid Petrol Cars (g/km/s)', u'LPG Cars (g/km/s)'],
               u'Diesel cars': [u'Diesel Cars (g/km/s)', u'Taxis (g/km/s)', u'Full Hybrid Diesel Cars (g/km/s)', u'E85 Bioethanol Cars (g/km/s)'],
               u'Petrol LGVs': [u'Petrol LGVs (g/km/s)', u'Full Hybrid Petrol LGVs (g/km/s)', u'Plug-In Hybrid Petrol LGVs (g/km/s)', u'LPG LGVs (g/km/s)'],
               u'Diesel LGVs': [u'Diesel LGVs (g/km/s)', u'E85 Bioethanol LGVs (g/km/s)'],
               u'Rigid HGVs': [u'Rigid HGVs (g/km/s)', u'B100 Rigid HGVs (g/km/s)'],
               u'Artic HGVs': [u'Artic HGVs (g/km/s)', u'B100 Artic HGVs (g/km/s)'],
               u'Buses and coaches': [u'Buses/Coaches (g/km/s)', u'B100 Buses (g/km/s)', u'CNG Buses (g/km/s)', u'Biomethane Buses (g/km/s)', u'Biogas Buses (g/km/s)', u'Hybrid Buses (g/km/s)', u'FCEV Buses (g/km/s)', u'B100 Coaches (g/km/s)'],
               u'Motorcycles': [u'Motorcycles (g/km/s)'],
               u'ElectricCars': [u'Battery EV Cars (g/km/s)', u'FCEV Cars (g/km/s)'],
               u'ElectricLGVs': [u'Battery EV LGVs (g/km/s)', u'FCEV LGVs (g/km/s)']}
# Vehicle subtyps for summing all emissions to 4 classes.
VehSubTypes = {'Cars': [u'Petrol cars', u'Diesel cars', u'ElectricCars'],
               'LGV': [u'Petrol LGVs', u'Diesel LGVs', u'ElectricLGVs' ],
               'HGV': [u'Rigid HGVs', u'Artic HGVs'],
               'Bus': [u'Buses and coaches'],
               'Ignore': [u'Motorcycles']}


def readNO2Factors(FactorFile):
  """
  Function that reads the contents of the NOxFactorFile.

  The NOxFactorFile is the contents of the "Fleet-avg by_vehicle_fuel_type'
  sheet from the PrimaryNO2_factors_NAEIBase_2016_v1.xlsx spreadsheet which is
  downloaded from http://naei.beis.gov.uk/data/ef-transport. This sheet is
  appropriate because we begin with only 4 vehicle classes.

  INPUTS
  FactorFile - string - the path to the NOxFactorFile.

  OUTPUTS
  Factors - the NOx to NO2 conversion factors, as a dictionary.
  """
  FactorsDF = pd.read_csv(FactorFile)
  years = list(FactorsDF)
  years.remove('f-NO2')
  Factors = {}
  #years = [int(y) for y in years]
  vehs = list(FactorsDF['f-NO2'])
  for veh in vehs:
    Factors[veh] = {}
    row = FactorsDF[FactorsDF['f-NO2'] == veh]
    for y in years:
      Factors[veh][int(y)] = float(row[y])
  return Factors

def getNO2Factor(Factors, Vehicle, Year):
  """
  Function to extract the NOx to NO2 conversion factor from the NOxFactorFile.

  This is neccesary because EFT does not offer NO2 emission rates. NAEI does,
  but all it does is apply a particular factor to the NAEI NOx emission rates,
  which are sourced from EFT anyway. So this method should be identical.

  The NOxFactorFile is the contents of the "Fleet-avg by_vehicle_fuel_type'
  sheet from the PrimaryNO2_factors_NAEIBase_2016_v1.xlsx spreadsheet which is
  downloaded from http://naei.beis.gov.uk/data/ef-transport. This sheet is
  appropriate because we begin with only 4 vehicle classes.

  INPUTS
  Factors - string - the path to the NOxFactorFile,
            or a dictionary - the contents of that file.
  Vehicle - string - One of Petrol cars, Diesel cars, Petrol LGVs, Diesel LGVs,
            Rigid HGVs, Artic HGVs, Buses and coaches or Motorcycles
  Year    - integer - a value between 2005 and 2035

  OUTPUTS
  Factor  - the NOx to NO2 conversion factor for the specified vehicle and year.
  """
  if Vehicle in ['ElectricCars', 'ElectricLGVs']:
    return 1 # Well it was either that or nothing. NOx emission rates are basically 0 anyway.

  # Is Factors a string?
  if isinstance(Factors, str):
    Factors = readNO2Factors(Factors)

  return Factors[Vehicle][Year]

def doEFT(data, fName, uniqueID='UID', excel='Create', AA='AADT'):
  """
  Function that adds data to the EFT, runs the EFT, and then extracts the data.


  INPUTS
  data  - pandas dataframe - A pandas dataframe containing the
  fName - string - the path to an empty EFT file that has been set up with the
                   correct year and area. Traffic format must be set to
                   'Detailed Option 1', NOx, PM10, and PM2.5 should be selected
                   under 'Select Pollutants'. 'Air Quality Modelling (g/km/s)'
                   should be selected under 'Select Outputs' and 'Breakdown by
                   Vehicle' should be selected under 'Additional Outputs'.
  """

  excelCreated = False
  if excel == 'Create':
    excelCreated = True
    excel = win32.gencache.EnsureDispatch('Excel.Application')

  # Make a copy of the empty EFT file.
  [FN, FE] =  os.path.splitext(fName)
  TempEFT = '{}_TEMP{:%Y%m%d%H%M%S}{}'.format(FN, datetime.datetime.now(), FE)
  TempEFTm = '{}_TEMP{:%Y%m%d%H%M%S}{}'.format(FN, datetime.datetime.now(), '.xlsm')
  shutil.copyfile(fName, TempEFT)
  TempEFT = os.path.abspath(TempEFT)        # Neccesary because win32 seems to
  TempEFTm = os.path.abspath(TempEFTm)      # struggle with relative paths.

  if AA == 'AADT':
    AAA = 'D'
  elif AA == 'AAWT':
    AAA = 'W'

  # We need to order the appropriate columns so that they can be copied into the
  # EFT input page.
  outData = data[[uniqueID, 'RoadType', AA, '{}_Cars'.format(AA),
                  '{}_Taxi'.format(AA), '{}_LGV'.format(AA), '{}_HGV'.format(AA),
                  '{}_Bus'.format(AA), '{}_MC'.format(AA), '{}SPD'.format(AA),
                  'Duration', 'Length']]

  numRows_ = len(outData.index)
  # Start off the autohotkey script as a (parallel) subprocess. This will
  # continually check until the compatibility warning appears, and then
  # close the warning.
  if ahk_exist:
    subprocess.Popen([ahk_exepath, ahk_ahkpath])


  # Open the document.
  wb = excel.Workbooks.Open(TempEFT)
  excel.Visible = True
  ws_input = wb.Worksheets("Input Data")
  year = ws_input.Range("B5").Value

  # Copy the values to the spreadsheet.
  ws_input.Range("A10:L{}".format(numRows_+9)).Value = outData.values.tolist()

  # Run the macro.
  excel.Application.Run("RunEfTRoutine")

  # Save the file as an .xlsm, so that it can be opened more easily.
  wb.SaveAs(TempEFTm, win32.constants.xlOpenXMLWorkbookMacroEnabled)
  wb.Close()

  # Extract the results.
  ex = pd.ExcelFile(TempEFTm)
  output = ex.parse("Output")
  pollutants = list(set(output['Pollutant Name']))
  #colnames = list(output)
  numRows = len(output.index)

  # Consolidate the different vehicle classes in preparation for creating NO2
  # emission rates.
  for VehSubType, FieldNames in VehSubTypesNO2.items():
    output[VehSubType] = [0]*numRows
    for fn in FieldNames:
      output[VehSubType] += output[fn]
      output.drop(fn, 1)

  # Drop the pre-consolidated columns.
  for fn in ['All Vehicles (g/km/s)', 'All LDVs (g/km/s)', 'All HDVs (g/km/s)']:
    output = output.drop(fn, 1)

  # Open the NO2 factor file.
  NO2Factors = readNO2Factors(NO2FactorFile)

  # Pivot the table so that each vehicle class has a column for each pollutant.
  for m in VehSubTypesNO2.keys():
    output_m = output.pivot_table(index='Source Name',
                                  columns='Pollutant Name',
                                  values=m)
    output_m = output_m.reset_index()
    output_m = output_m.sort_values('Source Name')
    if list(data[uniqueID]) != list(output_m['Source Name']):
      raise ValueError('SourceName and {} are not identical.'.format(uniqueID))
    for pol in pollutants:
      colName = '{}_{}{}'.format(m, pol, AAA)
      colName = colName.replace('.', '')
      data = data.assign(colName = list(output_m[pol]))
      data = data.rename(columns={'colName': colName}) # What!!!! Shouldn't be neccesary, but it is.
    # Create an NO2 column.
    colNameNO2 = '{}_NO2{}'.format(m, AAA)
    colNameNOx = '{}_NOx{}'.format(m, AAA)
    data[colNameNO2] = data[colNameNOx] * getNO2Factor(NO2Factors, m, year)
  pollutants.append('NO2')
  # Consolidate the columns into the 4 vehicle split.
  for veh4, vehp in VehSubTypes.items():
    for pol in pollutants:
      colName = '{}_{}{}'.format(veh4, pol, AAA)
      colName = colName.replace('.', '')
      data[colName] = [0]*len(data.index)
      for veh in vehp:
        colNamep = '{}_{}{}'.format(veh, pol, AAA)
        colNamep = colNamep.replace('.', '')
        data[colName] += data[colNamep]
        data = data.drop(colNamep, 1)
      if veh4 == 'Ignore':
        data = data.drop(colName, 1)

  # remove the temporary files.
  os.remove(TempEFT)
  os.remove(TempEFTm)
  if excelCreated:
    excel.Quit()
    del(excelObj) # Make sure it's gone. Apparently some people have found this neccesary.
  return data

def processNetwork(InputShapefile, EmptyEFT, NO2FactorFile, OutputShapefile='default', uniqueID='UID', Head=False, MaxRows=10000, area='NotSet'):
  """
  A function that will run the road counts within an input shape file through
  the EFT, and then extract the emission rates from the EFT and add them to a
  new shapefile.

  INPUTS
  InputShapefile - string - The path to the input shape file. This is designed
                            to work with the shapefiles with the noise traffic
                            counts. Any polyline shape file with the following
                            fields should work.
                   SPEED  - Numeric
                            The assumed average traffic speed on the road, in
                            kmph. Any integer value between 5 and 140.
                   Urban  - Numeric
                            Either 0 or 1. 0 for non urban areas, 1 for urban
                            areas.
                   Class  - String
                            This string is used to signify whether a road is a
                            motorway or not. If it takes any value within
                            MotorwayNames then it is assumed to be  motorway.
                            Otherwise it is not.
                   AADT and/or AAWT - numeric
                            The annual average daily traffic total along the
                            feature. AADT signifies the average for any day of
                            the week, AAWT signifies that only weekdays were
                            considered.
                   AAXT_Cars - numeric
                   AAXT_Bus    These are the proportions of each vehicle class
                   AAXT_LGV    that comprise the full vehicle count. e.g.
                   AAXT_HGV    AADT_Cars is the proportion of vehicles in AADT
                               that are Cars.
  EmptyEFT       - String - The path to an empty EFT file that has been set up
                            with the correct year and area (note that these
                            functions will only work for the 'Non-London'
                            areas). Traffic format must be set to 'Detailed
                            Option 1', NOx, PM10, and PM2.5 should be selected
                            under 'Select Pollutants'. 'Air Quality Modelling
                            (g/km/s)' should be selected under 'Select Outputs'
                            and 'Breakdown by Vehicle' should be selected under
                            'Additional Outputs'.
  NO2FactorFile  - String - The path to the NOx to NO2 conversion factor file.

  """
  # See if we can parse the area.
  if area == 'NotSet':
    posAreas = ['Dundee', 'Aberdeen', 'Glasgow', 'Edinburgh']
    for ar in posAreas:
      if InputShapefile.find(ar) >= 0:
        area = ar
        break

  # Import the data from the shapefile.
  print('Importing features from {}.'.format(InputShapefile))
  Data = gpd.read_file(InputShapefile)
  if Head:
    Data = Data.head(20)
  numRows = len(Data.index)
  columnNames = list(Data)
  print('  Done. Imported {} features.'.format(numRows))
  # Get the crs well known text, so that it can be assigned to the file to save.
  prj_file = InputShapefile.replace('.shp', '.prj')
  crs_wkt = [l.strip() for l in open(prj_file,'r')][0]
  print('Organising data.')

  # Make sure that the specified unique identifier is actually a unique identifier.
  if uniqueID not in columnNames:
    raise ValueError('The specified unique identifier "{}" is not an available field.'.format(uniqueID))
  Data = Data.sort_values(uniqueID)
  if len(set(Data[uniqueID])) != numRows:
    # It is not.
    seen = []
    notUnique = []
    IDs = list(Data[uniqueID])
    for ID in IDs:
      if ID in seen:
        notUnique.append(ID)
      else:
        seen.append(ID)
    print('{} is not a unique identifier. The following IDs are repeated at least once:'.format(uniqueID))
    print(notUnique)
    raise ValueError('{} is not a unique identifier'.format(uniqueID))

  # Check that other required fields exist.
  Required = ['Urban']                    # Need each of these.
  CountDayTypes = ['AADT', 'AAWT']        # Need at least one of these, and for
  VehTypes = ['Cars', 'Bus', 'LGV', 'HGV']# each of these we also require
                                          # 'AAXT_Cars', 'AAXT_Bus', 'AAXT_HGV'
                                          # and 'AAXT_LGV'. Plus we also need
                                          # 'AAXTSPD', the speed for that period.
  for Req in Required:
    if Req not in columnNames:
      raise ValueError('The required field "{}" is missing.'.format(Req))
  gotAtleastOne = False
  for Req in CountDayTypes:
    if Req in columnNames:
      gotAtleastOne = True
      AAveh = '{}SPD'.format(Req)
      if AAveh not in columnNames:
        raise ValueError('The required field "{}" is missing.'.format(AAveh))
      for veh in VehTypes:
        AAveh = '{}_{}'.format(Req, veh)
        if AAveh not in columnNames:
          raise ValueError('The required field "{}" is missing.'.format(AAveh))
  if not gotAtleastOne:
    raise ValueError('Neither "AADT" nor "AAWT" fields are available.')
  # And Class, if it does not exist then mark everything as 'other'.
  if 'Class' not in columnNames:
    Data['Class'] = ['other']*numRows

  # Find and remove streets that are either spatially identical or the exact
  # reverse of other streets. This is neccesary
  # because the original noise data had two roads for every two way street;
  # one in each direction.
  # To find the duplicated roads based on their geometry is incredibly slow, so
  # instead we use the UID, which if it was generated using the ArcGIS model
  # should be a set format. Unfortunately this makes the code less adaptable.

  print('Combining and removing duplicate streets.')
  DataUIDs = Data.ix[:, [uniqueID, 'SegmentID']]
  DataUIDs['AB_U'] = DataUIDs['SegmentID'].astype(str).str[:-2] + DataUIDs[uniqueID].astype(str).str[-2:]
  # Find duplicates
  DataUIDs = DataUIDs.ix[:, ['AB_U']]
  DataUIDs = DataUIDs.sort_values('AB_U')
  Duplicates1 = pd.DataFrame(DataUIDs.duplicated(keep='first'), columns=['Dup'])
  Duplicates1 = Duplicates1[Duplicates1['Dup'] == True]
  DupIndex1 = list(Duplicates1.index)
  Duplicates2 = pd.DataFrame(DataUIDs.duplicated(keep='last'), columns=['Dup'])
  Duplicates2 = Duplicates2[Duplicates2['Dup'] == True]
  DupIndex2 = list(Duplicates2.index)

  for DupI in xrange(len(DupIndex1)):
    Dup1 = DupIndex1[DupI]
    Dup2 = DupIndex2[DupI]
    row1 = Data.loc[Dup1]
    row2 = Data.loc[Dup2]
    for AA in CountDayTypes:
      AA1 = row1[AA]
      AA2 = row2[AA]
      if AA1+AA2 < 1e-8:   # If both are basically 0.
        Scaling1, Scaling2 = 0.5, 0.5
      else:
        Scaling1 = AA1/(AA1+AA2)
        Scaling2 = AA2/(AA1+AA2)
      ColsToDo = ['_'+q for q in VehTypes]
      ColsToDo.append('SPD')
      for Col in ColsToDo:
        ColName = AA+Col
        Col1 = row1[ColName]
        Col2 = row2[ColName]
        Value = Scaling1*Col1 + Scaling2*Col2
        Data = Data.set_value(Dup1, ColName, Value)
      Col1 = row1[AA]
      Col2 = row2[AA]
      Value = Col1 + Col2
      Data = Data.set_value(Dup1, AA, Value)
    Data = Data.drop([Dup2])

  numRowsOld = numRows
  numRows = len(Data.index)
  print('Done, removed {} features, {} remaining.'.format(numRowsOld - numRows, numRows))

  # Class is currently 'Motorway' or 'Other'.
  # LEGEND is currently 'Small Urban Area polygon' or 'Large Urban Area polygon',
  # or None.
  # We want to define all motorways as 'Motorway (not London)', all 'Large Urban
  # Area polygon' as 'Urban (not London)', and all other roads as 'Rural (not London)'.
  Data['RoadType'] = ['Rural (not London)'] * numRows
  Data.loc[Data.Urban == 1, 'RoadType'] = 'Urban (not London)'
  print('The following "other" road classes are defined. If any look like they should be considered motorways then you may need to edit the MotorwayNames variable in the script.')
  for cl in set(Data.Class):
    if cl not in MotorwayNames:
      print('  {}'.format(cl))
  for mway in MotorwayNames:
    Data.loc[Data.Class == mway, 'RoadType'] = 'Motorway (not London)'

  # And add a row for Duration (24 hours) and for link length.
  # Note that the link length is approximate because the roads are not all
  # georectified, but that doesn't matter because we're getting emission rate
  # per km anyway.
  Data['Duration'] = 24
  Data['Length'] = Data.length/1000

  # Now open the copied version of the EFT, and fill in the data.
  # Create the Excel Application object.
  excelObj = win32.gencache.EnsureDispatch('Excel.Application')

  for AA in CountDayTypes:
    # Deal with rows that do not have a sensible vehicle class breakdown. First
    # create a column that just has a flag to say what class it is.
    # If the vehicle class breakdown sums to more than 0.95 (it should be 1) then the flag will be 1.
    # If the vehicle class breakdown sums to less than 0.05 then the flag will be 0.
    # If the vehicle class breakdown sums to between 0.05 and 0.05 then the flag will be -1.
    Flag = '{}_FLAG'.format(AA)
    Car = '{}_Cars'.format(AA)
    Bus = '{}_Bus'.format(AA)
    HGV = '{}_HGV'.format(AA)
    LGV = '{}_LGV'.format(AA)
    Data[Flag] = 1
    Data.loc[Data[Car]+Data[Bus]+Data[HGV]+Data[LGV] < 0.95, Flag] = -1
    Data.loc[Data[Car]+Data[Bus]+Data[HGV]+Data[LGV] < 0.05, Flag] = 0
    # For rows flagged 0 above, say that all vehicles are cars.
    Data.loc[Data[Flag] == 0, Car] = 1.0
    Data.loc[Data[Flag] == 0, Bus] = 0.0
    Data.loc[Data[Flag] == 0, HGV] = 0.0
    Data.loc[Data[Flag] == 0, LGV] = 0.0
    # Now make all of the other rows sum to 1, by normalising the breakdown.
    # We are also multiplying the numbers by 100 to convert to percentages here too.
    Tot = Data[Car]+Data[Bus]+Data[HGV]+Data[LGV]
    for colName in [Car, Bus, HGV, LGV]:
      Data[colName] = 100.0 * Data[colName]/Tot

    # Get the road speed, etc.
    SPD = '{}SPD'.format(AA)
    Data['{}_Taxi'.format(AA)] = [0] * numRows
    Data['{}_MC'.format(AA)] = [0] * numRows
    Data.loc[Data[SPD] < 5, SPD] = 5
    Data.loc[Data[SPD] > 140, SPD] = 140

    # And start adding the data to the EFT, block by block.
    Start = 0
    End = MaxRows
    count = 0
    First = True
    while End < numRows:
      DataSlice = Data.iloc[Start:End]
      count += len(DataSlice.index)
      outData = doEFT(DataSlice, EmptyEFT, excel=excelObj, uniqueID=uniqueID, AA=AA)
      if First:
        outDataAll = outData
        First = False
      else:
        outDataAll = outDataAll.append(outData)
      Start = End
      End = Start + MaxRows
    # get the last few lines.
    End = numRows
    print('Processing row {} to {} for {}.'.format(Start, End, AA))
    DataSlice = Data.iloc[Start:End]
    count += len(DataSlice.index)
    outData = doEFT(DataSlice, EmptyEFT,  excel=excelObj, uniqueID=uniqueID, AA=AA)
    if First:
      outDataAll = outData
    else:
      outDataAll = outDataAll.append(outData)

    Data = outDataAll

    Data.drop('{}_Taxi'.format(AA), 1)
    Data.drop('{}_MC'.format(AA), 1)

  excelObj.Quit()
  del(excelObj) # Make sure it's gone. Apparently some people have found this neccesary.
  print('Processing complete.')

  # Add the area to the data.
  Data['Area'] = area

  # Prepare the save location.
  if OutputShapefile == 'default':
    # Create a save location.
    [FN, FE] =  os.path.splitext(InputShapefile)
    OutputShapefile = '{}_wEmissions{}'.format(FN, FE)
    t = 1
    while os.path.isfile(OutputShapefile):
      t += 1
      OutputShapefile = '{}_wEmissions({}){}'.format(FN, t, FE)
  print('Saving output shape file to {}.'.format(OutputShapefile))
  # Save the updated data file as a shapefile again.
  Data.to_file(OutputShapefile, driver='ESRI Shapefile', crs_wkt=crs_wkt)
  print('Processing complete.')

if __name__ == '__main__':
  EmptyEFT = 'input\EFT2017_v7.4_NoiseEmpty.xlsb'
  NO2FactorFile = 'input/NO2Extracted.csv'
  args = sys.argv
  args = args[1:]

  if '--EFT' in args:
    ei = args.index('--EFT') + 1
    EmptyEFT = args[ei]
    del args[ei]
    args.remove('--EFT')
  if '--NO2' in args:
    ei = args.index('--NO2') + 1
    NO2FactorFile = args[ei]
    del args[ei]
    args.remove('--NO2')

  for inputfile in args:
    processNetwork(inputfile, EmptyEFT, NO2FactorFile)


