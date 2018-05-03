# -*- coding: utf-8 -*-
"""
Functions that reads traffic counts within a shape file to be fed into the
EFT.
"""
from __future__ import print_function

import os
import argparse
import shutil
import subprocess
from datetime import datetime
import pandas as pd
import geopandas as gpd

import win32com.client as win32

import EFT_Tools as tools

# Set some defaults.
ahk_exist = True
for ahk_ in [tools.ahk_exepath, tools.ahk_ahkpath]:
  if not os.path.isfile(ahk_):
    ahk_exist = False
if not ahk_exist:
  print(["Either Autohotkey is not installed, or it is not installed in the ",
         "usual place, or the required AutoHotKey script 'closeWarning.ahk' ",
         "cannot be found. the function will work but you will have to watch ",
         "and close the EFT warning dialogue manually."])


defaultNO2File = 'input/NAEI_NO2Extracted.xlsx'
defaultVehClasses = ['MCYCLE', 'CAR', 'TAXI', 'LGV', 'RHGV_2X', 'RHGV_3X', 'RHGV_4X', 'AHGV_34X', 'AHGV_5X', 'AHGV_6X', 'BUS']
defaultVehBreakdown = 'Detailed Option 2'
defaultVehReClass = {'Detailed Option 1': {'Car': ['CAR', 'TAXI'],
                                           'Taxi': [],
                                           'LGV': ['LGV'],
                                           'HGV': ['RHGV_2X', 'RHGV_3X', 'RHGV_4X', 'AHGV_34X', 'AHGV_5X', 'AHGV_6X'],
                                           'Bus and Coach': ['BUS'],
                                           'Motorcycle': ['MCYCLE'],
                                           'Ignore': []},
                     'Detailed Option 2': {'Car': ['CAR', 'TAXI'],
                                           'Taxi': [],
                                           'LGV': ['LGV'],
                                           'Rigid HGV': ['RHGV_2X', 'RHGV_3X', 'RHGV_4X'],
                                           'Artic HGV': ['AHGV_34X', 'AHGV_5X', 'AHGV_6X'],
                                           'Bus and Coach': ['BUS'],
                                           'Motorcycle': ['MCYCLE'],
                                           'Ignore': []}}

MotorwayNames = ['MOTORWAY', 'motorway', 'Motorway', 'Motorway (not London)', 'M8', 'M74', 'M73', 'M77', 'M80', 'M876', 'M898', 'M9', 'M90']
RuralNames = ['Rural', 'rural', 'RURAL', 'Rural (not London)']

# Vehicle subtypes for extracting NO2 factors.
VehSubTypesNO2 = {u'Petrol cars': [u'Petrol Cars (g/km)', u'Full Hybrid Petrol Cars (g/km)', u'Plug-In Hybrid Petrol Cars (g/km)', u'LPG Cars (g/km)'],
               u'Diesel cars': [u'Diesel Cars (g/km)', u'Taxis (g/km)', u'Full Hybrid Diesel Cars (g/km)', u'E85 Bioethanol Cars (g/km)'],
               u'Petrol LGVs': [u'Petrol LGVs (g/km)', u'Full Hybrid Petrol LGVs (g/km)', u'Plug-In Hybrid Petrol LGVs (g/km)', u'LPG LGVs (g/km)'],
               u'Diesel LGVs': [u'Diesel LGVs (g/km)', u'E85 Bioethanol LGVs (g/km)'],
               u'Rigid HGVs': [u'Rigid HGVs (g/km)', u'B100 Rigid HGVs (g/km)'],
               u'Artic HGVs': [u'Artic HGVs (g/km)', u'B100 Artic HGVs (g/km)'],
               u'Buses and coaches': [u'Buses/Coaches (g/km)', u'B100 Buses (g/km)', u'CNG Buses (g/km)', u'Biomethane Buses (g/km)', u'Biogas Buses (g/km)', u'Hybrid Buses (g/km)', u'FCEV Buses (g/km)', u'B100 Coaches (g/km)'],
               u'Motorcycles': [u'Motorcycles (g/km)'],
               u'ElectricCars': [u'Battery EV Cars (g/km)', u'FCEV Cars (g/km)'],
               u'ElectricLGVs': [u'Battery EV LGVs (g/km)', u'FCEV LGVs (g/km)']}
# Vehicle subtyps for summing all emissions to 4 classes.
VehSubTypes = {'Cars': [u'Petrol cars', u'Diesel cars', u'ElectricCars'],
               'LGV': [u'Petrol LGVs', u'Diesel LGVs', u'ElectricLGVs' ],
               'HGV': [u'Rigid HGVs', u'Artic HGVs'],
               'Bus': [u'Buses and coaches'],
               'Ignore': [u'Motorcycles']}



def getEFTFile(version, directory='input'):
  """
  Return the absolute path to the appropriate file for the selected version.
  Will return an error if no file is available.
  """
  # First check that the directory exists.
  if not os.path.isdir(directory):
    raise ValueError('Cannot find directory {}.'.format(directory))

  # Now figure out the file name.
  if version == 6.0:
    fName = 'EFT2014_v6.0.2_NoiseEmpty.xls'
  elif version == 7.0:
    fName = 'EFT2016_v7.0_NoiseEmpty.xlsb'
  elif version == 7.4:
    fName = 'EFT2017_v7.4_NoiseEmpty.xlsb'
  else:
    raise ValueError('Version {} is not recognised.'.format(version))

  fname = '{}/{}'.format(directory, fName)

  # Check that file exists.
  if not os.path.exists(fname):
    raise ValueError('Cannot find file {}.'.format(fname))

  # return the absolute paths.
  return os.path.abspath(fname)

def doEFT(data, fName, area, year, vehBreakdown, no2file, fleetProportions={},
          excel='Create', keeptemp=False, saveloc=os.getcwd(), version=7.0):

  """
  Function that adds data to the EFT, runs the EFT, and then extracts the data.


  INPUTS
  data  - pandas dataframe - A pandas dataframe containing the
  fName - string - the path to an empty EFT file that has been set up with the
                   correct year and area. Traffic format must be set to
                   'Detailed Option 1', NOx, PM10, and PM2.5 should be selected
                   under 'Select Pollutants'. 'Air Quality Modelling (g/km)'
                   should be selected under 'Select Outputs' and 'Breakdown by
                   Vehicle' should be selected under 'Additional Outputs'.
  """

  excelCreated = False
  if excel == 'Create':
    excelCreated = True
    excel = win32.gencache.EnsureDispatch('Excel.Application')

  # create a copy of the EFT file, and prepare some file paths.
  savedir, bb = os.path.split(saveloc)
  bb, eftext = os.path.splitext(eftfile)
  tempeftfile =  os.path.join(savedir, 'TEMP_EFT_{}{}'.format(datetime.strftime(datetime.now(), '%Y%m%d_%H%M%S'), eftext))
  tempeftfileM =  os.path.join(savedir, 'TEMP_EFT_{}{}'.format(datetime.strftime(datetime.now(), '%Y%m%d_%H%M%S'), '.xlsm'))
  tempeftfile = os.path.abspath(tempeftfile)
  tempeftfileM = os.path.abspath(tempeftfileM)
  shutil.copyfile(eftfile, tempeftfile)

  # We need to order the appropriate columns so that they can be copied into the
  # EFT input page.
  colstodo = ['EFT_Index', 'EFT_RoadType', 'EFT_TrafficFlow']
  for veh in defaultVehReClass[vehBreakdown].keys():
    if veh != 'Ignore':
      colstodo.append('EFT_{}'.format(veh))
  colstodo.extend(['EFT_Speed', 'EFT_NoHours'])
  outData = data[colstodo]

  numRows_ = len(outData.index)
  # Start off the autohotkey script as a (parallel) subprocess. This will
  # continually check until the compatibility warning appears, and then
  # close the warning.
  if ahk_exist:
    subprocess.Popen([tools.ahk_exepath, tools.ahk_ahkpath])


  # Open the document.
  wb = excel.Workbooks.Open(tempeftfile)
  excel.Visible = True
  ws_input = wb.Worksheets("Input Data")
  #year = ws_input.Range("B5").Value

  # Copy the values to the spreadsheet.
  ws_input.Range("B4").Value = area
  ws_input.Range("B5").Value = year
  ws_input.Range("B6").Value = vehBreakdown
  ws_input.Range("A10:L{}".format(numRows_+9)).Value = outData.values.tolist()

  # Now we need to populate the UserEuro table with the defaults.
  ws_euro = wb.Worksheets("UserEuro")
  ws_euro.Select()
  # There is a macro to do this, but for some reason it fails on versions 7.4
  # and 8.0 when run on my computer. So we must do it ourselves..
  tools.pasteDefaultEuroProportions(ws_euro, tools.versionDetails[version])
  #excel.Application.Run("PasteDefaultEuroProportions")

  # Change the fleet proportions, if required.
  if bool(fleetProportions):
    ws_euro = wb.Worksheets("UserEuro")
    for key, value in fleetProportions.items():
      ws_euro.Range(key).Value = value
  ws_input.Select()

  # Run the macro.
  excel.Application.Run("RunEfTRoutine")

  # Save the file as an .xlsm, so that it can be opened more easily.
  wb.SaveAs(tempeftfileM, win32.constants.xlOpenXMLWorkbookMacroEnabled)
  wb.Close()

  # Extract the results.
  ex = pd.ExcelFile(tempeftfileM)
  output = ex.parse("Output")
  pollutants = list(set(output['Pollutant Name']))
  numRows = len(output.index)

  if version < 8.0:
    # Consolidate the different vehicle classes in preparation for creating NO2
    # emission rates.
    for VehSubType, FieldNames in VehSubTypesNO2.items():
      output[VehSubType] = [0]*numRows
      for fn in FieldNames:
        output[VehSubType] += output[fn]
        output.drop(fn, 1)

    # Drop the pre-consolidated columns.
    for fn in ['All Vehicles (g/km)', 'All LDVs (g/km)', 'All HDVs (g/km)']:
      output = output.drop(fn, 1)

    # Open the NO2 factor file.
    NO2Factors = tools.readNO2Factors(no2file, mode='ByFuel')

    # Pivot the table so that each vehicle class has a column for each pollutant.
    for m in VehSubTypesNO2.keys():
      output_m = output.pivot_table(index='Source Name',
                                    columns='Pollutant Name',
                                    values=m)
      output_m = output_m.reset_index()
      output_m = output_m.sort_values('Source Name')
      for pol in pollutants:
        colName = '{}_{}'.format(m, pol)
        colName = colName.replace('.', '')
        data = data.assign(colName = list(output_m[pol]))
        data = data.rename(columns={'colName': colName}) # What!!!! Shouldn't be neccesary, but it is.
      # Create an NO2 column.
      colNameNO2 = '{}_NO2'.format(m)
      colNameNOx = '{}_NOx'.format(m)
      if m.find('Petrol') >= 0:
        fuel = 'Petrol'
      if m.find('Motorcycle') >= 0:
        fuel = 'Petrol'
      elif m.find('Electric') >= 0:
        fuel = 'Electric'
      else:
        fuel = 'Diesel'
      if m.find('Bus') >= 0:
        mv = 'Bus and Coach'
      elif m.find('Rigid') >= 0:
        mv = 'Rigid HGV'
      elif m.find('Artic') >= 0:
        mv = 'Artic HGV'
      elif m.find('LGV') >= 0:
        mv = 'LGV'
      elif m.find('Motorcycle') >= 0:
        mv = 'Motorcycle'
      elif m.upper().find('CAR') >= 0:
        mv = 'Car'
      else:
        print(m)
        mv = None # Will raise an error, unless fuel is electric.
      if fuel == 'Electric':
        data[colNameNO2] = data[colNameNOx]
      else:
        data[colNameNO2] = data[colNameNOx] * NO2Factors[fuel][mv][year] #tools.getNO2Factor(NO2Factors, m, year)

    pollutants.append('NO2')

    # Now consolidate all of the vehicle columns.
    for pol in pollutants:
      data['E_{}'.format(pol)] = 0

    for colname in list(data):
      for pol in pollutants:
        pname = pol.replace('.', '')
        if colname[-1*len(pname):] == pname:
          data['E_{}'.format(pol)] += data[colname]
          data = data.drop(colname, 1)


  else: # Version is 8.0 or higher.
    # Keep only the columns we need.
    colsToKeep = ['Source Name', 'Pollutant Name', 'All Vehicles (g/km)']
    colNames = list(output)
    for col in colNames:
      if col not in colsToKeep:
        output = output.drop(col, 1)

    # Pivot the table so that each vehicle class has a column for each pollutant.
    output = output.pivot_table(index='Source Name',
                                columns='Pollutant Name',
                                values='All Vehicles (g/km)')
    output = output.reset_index()

    renames = {}
    colstodo = []
    for pol in pollutants:
      pol_ = pol.replace('.', '')
      rename = 'E_{}'.format(pol_)
      renames[pol] = rename
      colstodo.append(rename)
    output = output.rename(columns=renames)

    if 'NOx' in pollutants:
      outputNO2 = ex.parse("Output_f-NO2")
      # Check that source name is in the same order as in output.
      sN_1 = list(output['Source Name'])
      sN_2 = list(outputNO2['Source Name'])
      if sN_1 != sN_2:
        # Shouldn't happen, but just in case.
        raise Exception('Source Names are not equal!!!')
      output['E_f-NO2'] = list(outputNO2['f-NO2 Value'])
      output['E_NO2'] = output['E_NOx'] * output['E_f-NO2']
      colstodo.extend(['E_f-NO2', 'E_NO2'])
      pollutants.append('NO2')

    data = data.copy()
    for col in colstodo: #.loc
      data[col] = list(output[col])

  # remove the temporary files.
  if keeptemp:
    os.remove(tempeftfile)
    os.remove(tempeftfileM)
  else:
    print('The following temporary files have not been removed.')
    print(tempeftfile)
    print(tempeftfileM)
  if excelCreated:
    excel.Quit()
    del(excelObj) # Make sure it's gone. Apparently some people have found this neccesary.

  # Calculate total emission masses for each road.
  #data['LengthKM'] = data.length/1000.
  #data = data.copy()
  for pol in pollutants:
    pol_ = pol.replace('.', '')
    data['T_{}'.format(pol_)] = data['E_{}'.format(pol_)] * data.length/1000.#['LengthKM']


  #print(data.head())



  return data, pollutants



def processNetwork(shapefile, eftfile, no2file=None, saveloc=None,
                   year=datetime.now().year, area='Scotland', keepTemp=False,
                   fleetProportions={},
                   vehFieldNames=defaultVehClasses,
                   vehBreakdown=defaultVehBreakdown,
                   speedFieldName='SPEED',
                   classFieldName='class',
                   MaxRows=10000, Head=False, keeptemp=False):

  """


  """

  # Import the data from the shapefile.
  print('Importing features from {}.'.format(shapefile))
  Data = gpd.read_file(shapefile)
  if Head:
    Data = Data.head(20)
  numRows = len(Data.index)
  columnNames = list(Data)
  print('  Done. Imported {} features.'.format(numRows))
  # Get the crs well known text, so that it can be assigned to the file to save.
  prj_file = shapefile.replace('.shp', '.prj')
  crs_wkt = [l.strip() for l in open(prj_file,'r')][0]

  print('Organising data.')
  # Check that required fields exist.
  required = [speedFieldName]                    # Need each of these.
  if classFieldName is not None:
    required.append(classFieldName)
  for req in required:
    if req not in columnNames:
      raise ValueError('The required field "{}" is missing.'.format(req))

  # check that atleast one of the vehicles exists.
  got1Veh = False
  for veh in vehFieldNames:
    if veh not in columnNames:
      print('There is no field for vehicle {}.'.format(veh))
    else:
      got1Veh = True
  if not got1Veh:
    raise ValueError('No fields exist for any vehicle.')

  # Add the appropriate fields that will be needed by the EFT.
  Data['EFT_RoadType'] = 'Urban (not London)'
  if classFieldName is not None:
    for rur in RuralNames:
      Data.loc[Data[classFieldName] == rur, 'EFT_RoadType'] = 'Rural (not London)'
    for mot in MotorwayNames:
      Data.loc[Data[classFieldName] == mot, 'EFT_RoadType'] = 'Motorway (not London)'
  Data['EFT_TrafficFlow'] = 0
  for veh in vehFieldNames:
    Data['EFT_TrafficFlow'] += Data[veh]
  # Veh counts
  for key, value in defaultVehReClass[vehBreakdown].items():
    if key == 'Ignore':
      continue
    colName = 'EFT_{}'.format(key)
    Data[colName] = 0
    for v in value:
      if (v in vehFieldNames) and (v in columnNames):
        Data[colName] += Data[v]
    Data[colName] = 100*Data[colName]/Data['EFT_TrafficFlow']
  # Speed
  Data['EFT_Speed'] = Data[speedFieldName]
  Data['EFT_NoHours'] = 24  # Assume traffic was counted over a day, or has been normalized to 24 hours.
  Data['EFT_Index'] = range(len(Data.index))

  # Now open the copied version of the EFT, and fill in the data.
  # Create the Excel Application object.
  excelObj = win32.gencache.EnsureDispatch('Excel.Application')

  # And start adding the data to the EFT, block by block.
  Start = 0
  End = MaxRows
  count = 0
  First = True
  while End < numRows:
    print('Processing row {} to {}.'.format(Start, End))
    DataSlice = Data.iloc[Start:End]
    count += len(DataSlice.index)
    outData, _ = doEFT(DataSlice, eftfile, area, year, vehBreakdown, no2file, excel=excelObj, keeptemp=keeptemp, saveloc=saveloc, version=version, fleetProportions=fleetProportions)
    if First:
      outDataAll = outData
      First = False
    else:
      outDataAll = outDataAll.append(outData)
    Start = End
    End = Start + MaxRows
  # get the last few lines.
  End = numRows
  print('Processing row {} to {}.'.format(Start, End))
  DataSlice = Data.iloc[Start:End]
  count += len(DataSlice.index)
  outData, pols = doEFT(DataSlice, eftfile, area, year, vehBreakdown, no2file, keeptemp=keeptemp, saveloc=saveloc, excel=excelObj, version=version, fleetProportions=fleetProportions)
  if First:
    outDataAll = outData
  else:
    outDataAll = outDataAll.append(outData)

  Data = outDataAll

  excelObj.Quit()
  del(excelObj) # Make sure it's gone. Apparently some people have found this neccesary.

  print('Processing complete.')
  print('Total emission masses are as follows:')
  for pol in pols:
    pol_ = pol.replace('.', '')
    print('{:6}: {:9.0f} kg'.format(pol, Data['T_{}'.format(pol_)].sum()))

  print('Saving output shape file to {}.'.format(OutputShapefile))
  # Save the updated data file as a shapefile again.
  Data.to_file(OutputShapefile, driver='ESRI Shapefile', crs_wkt=crs_wkt)
  print('Done.')

if __name__ == '__main__':
  ShapefileDescription = ("This programme will take is designed to work with shape files "
                          "produced for the traffic noise modelling project. "
                          "See details below.")

  parser = argparse.ArgumentParser(description="Processes the contents of a "
                                   "shape file through the Emission Factor "
                                   "Toolkit (EFT).")
  parser.add_argument('shapefile', type=str,
                      help="The shapefile to be processed.")# "+ShapefileDescription)
  parser.add_argument('eftfile', type=str,
                      help="The EFT file to use.")
  parser.add_argument('--vehFleetSplit', metavar='Vehicle euro class and weight split file.',
                      type=str, nargs='?', default=None,
                      help=("A euro split and weight split proportions file. A "
                            "template is available in the 'input' directory of "
                            "the repository."))
  parser.add_argument('--vehCountNames', metavar='Vehicle count field names',
                      type=str, nargs='?', default=defaultVehClasses,
                      help=("The shapefile field names for the vehicles "
                            "counts. Default \"{}\".").format(" ".join(defaultVehClasses)))
  parser.add_argument('--trafficFormat', metavar='Traffic Format',
                      type=str, nargs='?', default=defaultVehBreakdown,
                      help=("The traffic format to be used by the EFT. "
                            "Default '{}'.").format(defaultVehBreakdown))
  parser.add_argument('-a', metavar='area',
                      type=str, nargs='?', default='Scotland',
                      help="The areas to be processed. One of '{}'. Default 'Scotland'.".format("', '".join(tools.availableAreas)))
  parser.add_argument('-y', metavar='year',
                      type=int, nargs='?', default=datetime.now().year,
                      choices=range(2008, 2031),
                      help="The year to be processed. Default present year.")
  parser.add_argument('--saveloc', metavar='output shape file location',
                      type=str,   nargs='?', default=None,
                      help=("Location to save the output shape file. If not "
                            "assigned one will be created based on the input "
                            "shapefile."))
  parser.add_argument('--speedFieldName', metavar='speed field name',
                      type=str, nargs='?', default='SPEED',
                      help=("The shapefile field name for the road speed, "
                            "which itself should be in kmh."))
  parser.add_argument('--classFieldName', metavar='class field name',
                      type=str, nargs='?', default=None,
                      help=("The shapefile field name for the road class. "
                            "Roads will be processed as Urban roads, "
                            "unless they are marked, in this field, as 'Motorway' "
                            "(or any Scottish motorway name, e.g. 'M8') or 'Rural'. "
                            "Default None, which will set all road to Urban."))
  parser.add_argument('--no2file', metavar='no2 factor file',
                      type=str,   nargs='?', default=None,
                      help=("The NOx to NO2 conversion factor file to use. Has "
                            "no effect for EFT v8.0. Default {}.".format(defaultNO2File)))


  parser.add_argument('--keeptemp', metavar='keeptemp',
                      type=bool,  nargs='?', default=False,
                      help="Whether to keep or delete temporary files. Boolean. Default False (delete).")
  args = parser.parse_args()

  shapefile = args.shapefile
  eftfile = args.eftfile
  propfile = args.vehFleetSplit
  no2file = args.no2file
  saveloc = args.saveloc
  #combine = args.combine_coalligned
  year = args.y
  keeptemp = args.keeptemp
  version, versionblah = tools.extractVersion(eftfile)
  files2check = ((shapefile, 'Shape file'),
                 (eftfile, 'EFT file'),
                 (propfile, 'Fleet proportions file'),
                 (no2file, 'NO2 conversion file'))
  for f2c in files2check:
    fpath = f2c[0]
    print(fpath)
    if fpath is not None:
      # It won't matter if it is None, the parser will have caught required files.
      if not os.path.exists(f2c[0]):
        raise ValueError('{} cannot be found at {}.'.format(f2c[1], f2c[0]))

  ## Get the NOx to NO2 file, if relevent, and other version dependent stuff.
  if version == 8.0:
    availableYears = range(2015, 2031)
    if no2file is not None:
      raise ValueError('An NOx to NO2 file has been specified for EFT v8.0. This is not neccesary.')
    no2file = defaultNO2File # Set it anyway, it's useful for debugging and doesn't do any harm.
  else:
    if no2file == 'default':
      no2file = defaultNO2File
    if version == 6.0:
      availableYears = range(2008, 2031)
    elif version in [7.0, 7.4]:
      availableYears = range(2013, 2031)
    else:
      raise ValueError('Version {} is not known.'.format(version))
    if not os.path.exists(no2file):
      raise ValueError('NO2 conversion file cannot be found at {}.'.format(no2file))
  if year not in availableYears:
    raise ValueError('Year {} is not allowed for EFT version {}.'.format(year, version))

  ## Read the Fleet proportions file, if set.
  fleetprops = tools.readFleetProps(propfile)

  ## see if all the vehicle classes make sense.
  vehs = args.vehCountNames
  gots = [0]*len(vehs)
  vehBreakdown = args.trafficFormat
  vall = []
  for value in defaultVehReClass[vehBreakdown].values():
    vall.extend(value)
  for veh in vehs:
    if veh not in vall:
      raise ValueError('Vehicle {} is not included for breakdown {}.'.format(veh, vehBreakdown))

  ## Specify a save location.
  if saveloc is None:
    # Create a save location.
    [FN, FE] =  os.path.splitext(shapefile)
    OutputShapefile = '{}_wEmissions{}{}'.format(FN, year, FE)
    t = 1
    while os.path.isfile(OutputShapefile):
      t += 1
      OutputShapefile = '{}_wEmissions{}({}){}'.format(FN, year, t, FE)
    saveloc = OutputShapefile



  processNetwork(shapefile, eftfile,
                 no2file=no2file,
                 saveloc=saveloc,
                 year=year,
                 area=args.a,
                 fleetProportions=fleetprops,
                 vehFieldNames=vehs,
                 vehBreakdown=vehBreakdown,
                 speedFieldName=args.speedFieldName,
                 classFieldName=args.classFieldName,
                 keeptemp=args.keeptemp)