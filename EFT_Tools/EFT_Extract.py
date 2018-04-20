# -*- coding: utf-8 -*-
"""
extractOutput
readProportions

Created on Fri Apr 20 15:30:03 2018

@author: edward.barratt
"""

import pandas as pd

from EFT_Tools import (in2outVeh,
                       splitSourceNameS,
                       splitSourceNameT,
                       splitSourceNameV)


def extractOutput(fileName, versionForOutPut, year, location, euroClass, details, techDetails=[None, None]):
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
  #output = output.drop(details['SourceNameName'], 1)
  output = output.drop(details['AllLDVName'], 1)
  output = output.drop(details['AllHDVName'], 1)
  tech = 'Default Mix'
  if not any([v is None for v in techDetails]):
    # Drop rows with vehicles that are not relavent to the technology specified.
    vehiclesPresent = list(output['vehicle'].unique())
    vehGot = {}
    for vP in vehiclesPresent:
      vehGot[vP] = False

    tech = techDetails[0]
    gotTechs = techDetails[1]
    for key, values in in2outVeh.items():
      got = False
      for veh in values:
        if veh in gotTechs.keys():
          if gotTechs[veh]:
            got = True
            break
      if got:
        if key in vehiclesPresent:
          vehGot[key] = True

    for veh, vehGot_ in vehGot.items():
      if not vehGot_:
        output = output[output['vehicle'] != veh]
  # Change the name of the vehicles to recognise the specified technology.
  output['mitigation tech'] = tech

  # Pivot the table so each pollutant has a column.
  pollutants = list(output[details['PolName']].unique())
  # Rename, because after the pivot the 'column' name will become the
  # index name.
  output = output.rename(columns={details['PolName']: 'RowIndex'})
  output = output.pivot_table(index=[details['SourceNameName'], 'year', 'area',
                                     'euro', 'version', 'speed', 'vehicle',
                                     'type', 'mitigation tech'],
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
    renames[Pol] = '{} (g/km/veh)'.format(Pol_)
  output = output.rename(columns=renames)

  # See if the f-NO2 output sheet is available (only available in version 8,
  # and if requested).
  if 'Output_f-NO2' in ex.sheet_names:
    output_f_NO2 = ex.parse('Output_f-NO2')
    if not output_f_NO2.empty:
      # The data's there. join it to our output, but remove a couple of columns first.
      colnames = list(output_f_NO2)
      # drop unneccesary columns.
      for colname in colnames:
        if colname not in [details['SourceNameName'], 'f-NO2 Value']:
          output_f_NO2 = output_f_NO2.drop(colname, axis=1)
      output = pd.merge(output, output_f_NO2, how='inner',
                        left_on=details['SourceNameName'],
                        right_on=details['SourceNameName'])
      # Rename the f-NO2 Column
      output = output.rename(columns={'f-NO2 Value': 'f-NO2'})
  output = output.drop(details['SourceNameName'], 1)
  return output

def readProportions(fileName, details, location, year,
                    ahk_exepath, ahk_ahkpathG,
                    versionForOutPut, excel=None, logger=None):
  closeExcel = False
  if excel is None:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    closeExcel = True

  # Get the logging details.
  loggerM = getLogger(logger, 'readProportions')

  # Start off the autohotkey script as a (parallel) subprocess. This will
  # continually check until the compatibility warning appears, and then
  # close the warning.
  if path.isfile(ahk_exepath):
    subprocess.Popen([ahk_exepath, ahk_ahkpathG])

  # Open the document.
  wb = excel.Workbooks.Open(fileName)
  excel.Visible = True

  # Set the default values in the Input Data sheet.
  ws_input = wb.Worksheets("Input Data")
  ws_input.Range("B4").Value = location
  ws_input.Range("B5").Value = year

  ws_euro = wb.Worksheets("UserEuro")

  logprint(loggerM, 'Extracting default euro split proportions.', level='info')

  startrows = ['vehRowStarts', 'vehRowStartsMC', 'vehRowStartsHB']
  endrows = ['vehRowEnds', 'vehRowEndsMC', 'vehRowEndsHB']
  vehtypes = ['Most Vehicles', 'Motorcycles', 'Hybrid Buses']
  EuroClassNameColumnsDict = {'Most Vehicles': EuroClassNameColumns,
                              'Motorcycles': EuroClassNameColumnsMC,
                              'Hybrid Buses': EuroClassNameColumnsHB}
  DefaultEuroColumnsDict = {'Most Vehicles': DefaultEuroColumns,
                            'Motorcycles': DefaultEuroColumnsMC,
                            'Hybrid Buses': DefaultEuroColumnsHB}
  #UserEuroColumnsDict = {'Most Vehicles': UserDefinedEuroColumns,
  #                       'Motorcycles': UserDefinedEuroColumnsMC,
  #                       'Hybrid Buses': UserDefinedBusColumn}

  first = True
  for ci, poltype in enumerate(['NOx', 'PM']):
    for ri in range(len(startrows)):
      logprint(loggerM, '  Dealing with euro proportions for {} - {}.'.format(vehtypes[ri], poltype), level='info')
      #print(ri, ci)
      #print(vehtypes[ri])
      #print(UserEuroColumnsDict[vehtypes[ri]])
      ColName = EuroClassNameColumnsDict[vehtypes[ri]][ci]
      ColProp = DefaultEuroColumnsDict[vehtypes[ri]][ci]
      ColUser = UserDefinedEuroColumns[ci]
      vehRowStarts = details[startrows[ri]]
      vehRowEnds = details[endrows[ri]]
      propdf = getProportions(ws_euro, ColName, ColProp, ColUser, vehRowStarts,
                              vehRowEnds, mode=vehtypes[ri], logger=loggerM)
      propdf['poltype'] = poltype
      if first:
        df = propdf
        first = False
      else:
        df = df.append(propdf)
  df['year'] = year
  df['area'] = location
  df_allEuros = df

  # Weight split!
  logprint(loggerM, 'Extracting default weight split proportions.', level='info')

  startrows = ['weightRowStarts', 'weightRowStartsBus']
  endrows = ['weightRowEnds', 'weightRowEndsBus']
  vehtypes = ['Most Vehicles', 'Buses']

  first = True
  for ri in range(len(startrows)):
    logprint(loggerM, '  Dealing with weight proportions for {}.'.format(vehtypes[ri]), level='info')
    ColName = NameWeightColumn
    ColProp = DefaultWeightColumn
    ColUser = UserDefinedWeightColumn
    vehRowStarts = details[startrows[ri]]
    vehRowEnds = details[endrows[ri]]
    propdf = getProportions(ws_euro, ColName, ColProp, ColUser, vehRowStarts,
                            vehRowEnds, mode='Weights', logger=loggerM)
    if first:
      df = propdf
      first = False
    else:
      df = df.append(propdf)
  df['year'] = year
  df['area'] = location
  df_weights = df

  wb.Close(True)
  if closeExcel:
    excel.Quit()
    del(excelObj)

  # Now extract euro classes on their own (no techs)
  logprint(loggerM, 'Concolidating euro classes.', level='info')

  vehs = set(df_allEuros['vehicle'])
  eurs = set(df_allEuros['euroclass'])
  pols = set(df_allEuros['poltype'])
  first = True
  for veh in vehs:
    df_1 = df_allEuros[df_allEuros['vehicle'] == veh]
    for eur in eurs:
      df_2 = df_1[df_1['euroclass'] == eur]
      for pol in pols:
        df_3 = df_2[df_2['poltype'] == pol]
        ss = df_3['proportion'].sum()
        df_ = pd.DataFrame([[veh, eur, pol, ss]],
                           columns=['vehicle', 'euroclass',
                                    'poltype', 'proportion'])
        if first:
          df = df_
          first = False
        else:
          df = df.append(df_)
  df['year'] = year
  df['area'] = location
  df_consEuros = df
  return df_allEuros, df_weights, df_consEuros