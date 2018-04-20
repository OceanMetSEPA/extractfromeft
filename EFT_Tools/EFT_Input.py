# -*- coding: utf-8 -*-
"""
checkEuroClassesValid
createEFTInput
getProportions
readFleetProps
specifyBusCoach
specifyEuroProportions
SpecifyWeight

Created on Fri Apr 20 14:51:30 2018

@author: edward.barratt
"""
import re
import inspect
import numpy as np
import pandas as pd
import pywintypes
import xlrd

from EFT_Tools import (availableRoadTypes,
                       euroClassNameVariations,
                       euroClassNameVariationsAll,
                       euroClassNameVariationsIgnore,
                       euroSearchTerms,
                       euroTechs,
                       getLogger,
                       logprint,
                       techVehs,
                       VehSplits,
                       weightClassNameVariations)

def checkEuroClassesValid(workBook, vehRowStarts, vehRowEnds, EuroClassNameColumns, Type=99, logger=None):
  """
  Check that all of the available euro classes are specified.
  """
  parentFrame = inspect.currentframe().f_back
  (filename, xa, xb, xc, xd) = inspect.getframeinfo(parentFrame)

  # Get the logging details.
  loggerM = getLogger(logger, 'checkEuroClassesValid')

  if Type == 1:
    logprint(loggerM, "Checking all motorcycle euro class names are understood.")
  elif Type == 2:
    logprint(loggerM, "Checking all hybrid bus euro class names are understood.")
  elif Type == 0:
    logprint(loggerM, "Checking all other euro class names are understood.")
  else:
    logprint(loggerM, "Checking all euro class names are understood.")

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
          if ecn not in euroClassNameVariationsIgnore:
            raise ValueError('Unrecognized Euro Class Name: "{}".'.format(ecn))

def createEFTInput(vBreakdown='Detailed Option 2',
                   speeds=[5,6,7,8,9,10,12,14,16,18,20,25,30,35,40,
                           45,50,60,70,80,90,100,110,120,130,140],
                   roadTypes=availableRoadTypes,
                   vehiclesToSkip=['Taxi (black cab)'],
                   vehiclesToInclude=None,
                   tech='All',
                   logger=None):
  """
  vehiclesToInclude trumps (and overwrites) vehiclesToSkip
  """
  # Get the logging details.
    # Get the logging details.
  loggerM = getLogger(logger, 'createEFTInput')

  logprint(loggerM, 'Creating EFT input.', level='debug')
  logprint(loggerM, 'Initial vehiclesToSkip: {}'.format(', '.join(vehiclesToSkip)), level='debug')
  logprint(loggerM, 'Initial vehiclesToInclude: {}'.format(', '.join(vehiclesToInclude)), level='debug')
  VehSplit = VehSplits[vBreakdown]
  logprint(loggerM, 'VehSplit: {}'.format(', '.join(VehSplit)), level='debug')

  if vehiclesToInclude is not None:
    # Populate vehiclesToSkip with those vehicles that are not included.
    vehiclesToSkip = []
    for veh in VehSplit:
      if veh not in vehiclesToInclude:
        vehiclesToSkip.append(veh)
  logprint(loggerM, 'Intermediate vehiclesToSkip: {}'.format(', '.join(vehiclesToSkip)), level='debug')
  #RoadTypes = ['Urban (not London)', 'Rural (not London)', 'Motorway (not London)']
  if tech != 'All':
    # Add vehicles to vehiclesToSkip that are irrelevant for the chosen technology.
    for veh in VehSplit:
      if veh not in techVehs[tech]:
        vehiclesToSkip.append(veh)
  vehiclesToSkip = list(set(vehiclesToSkip))
  logprint(loggerM, 'Final vehiclesToSkip: {}'.format(', '.join(vehiclesToSkip)), level='debug')

  if type(roadTypes) is str:
    if roadTypes in ['all', 'All', 'ALL']:
      roadTypes = availableRoadTypes
    else:
      roadTypes = [roadTypes]

  if vBreakdown == 'Basic Split':
    numRows = 2*len(roadTypes)*len(speeds)
  else:
    numRows = len(roadTypes)*len(speeds)*(len(VehSplit)-len(vehiclesToSkip))
  numCols = 6 + len(VehSplit)

  inputDF = pd.DataFrame(index=range(numRows), columns=range(numCols))
  ri = -1

  for rT in roadTypes:
    logprint(loggerM, 'roadType - {}'.format(rT), level='debug')
    for sp in speeds:
      logprint(loggerM, '  speed - {}'.format(sp), level='debug')
      for veh in VehSplit:
        logprint(loggerM, '    vehicle - {}'.format(veh), level='debug')
        #print('    veh - {}'.format(veh))
        if vBreakdown == 'Basic Split':
          ri += 2
          #inputDF.set_value(ri-1, 0, 'S{} - LDV - {}'.format(sp, rT))
          inputDF.iat[ri-1, 0] = 'S{} - LDV - {}'.format(sp, rT)
          inputDF.iat[ri-1, 1] = rT
          inputDF.iat[ri-1, 2] = 1
          inputDF.iat[ri-1, 3] = 0
          inputDF.iat[ri-1, 4] = sp
          inputDF.iat[ri-1, 5] = 1
          inputDF.iat[ri-1, 6] = 1
          inputDF.iat[ri, 0] = 'S{} - HDV - {}'.format(sp, rT)
          inputDF.iat[ri, 1] = rT
          inputDF.iat[ri, 2] = 1
          inputDF.iat[ri, 3] = 100
          inputDF.iat[ri, 4] = sp
          inputDF.iat[ri, 5] = 1
          inputDF.iat[ri, 6] = 1
        else:
          if veh in vehiclesToSkip:
            logprint(loggerM, '      skipped', level='debug')
            pass
          else:
            logprint(loggerM, '      including', level='debug')
            ri += 1
            inputDF.iat[ri, 0] = 'S{} - {} - {}'.format(sp, veh, rT)
            inputDF.iat[ri, 1] = rT
            inputDF.iat[ri, 2] = 1
            for vehi, vehb in enumerate(VehSplit):
              if vehb == veh:
                inputDF.iat[ri, 3+vehi] = 100
              else:
                inputDF.iat[ri, 3+vehi] = 0
            inputDF.iat[ri, len(VehSplit)+3] = sp
            inputDF.iat[ri, len(VehSplit)+4] = 1 # 1 hour. Not neccesary for g/km output.
            inputDF.iat[ri, len(VehSplit)+5] = 1 # 1 km. Not neccesary either.
            logprint(loggerM, '        done', level='debug')

  inputData = inputDF.as_matrix()
  inputShape = np.shape(inputData)
  logprint(loggerM, 'input created with dimensions {} by {}.'.format(inputShape[0], inputShape[1]), level='debug')
  return inputData

def getProportions(ws, ColName, ColProp, ColUser, vehRowStarts,
                   vehRowEnds, mode='Most Vehicles', logger=None):

  # Get the logging details.
  loggerM = getLogger(logger, 'getProportions')

  # Start a pandas dateframe.
  df = pd.DataFrame(columns=['vehicle', 'euroname', 'euroclass', 'technology',
                             'proportion', 'sourceCell', 'userCell'])
  for vehi in range(len(vehRowStarts)):
    starow = vehRowStarts[vehi]
    endrow = vehRowEnds[vehi]
    if mode == 'Most Vehicles':
      vehName = ws.Range("{}{}".format(ColName, starow-1)).Value
      while vehName is None:
        vehName = ws.Range("{}{}".format(ColName, starow)).Value
        starow += 1
    elif mode == 'Motorcycles':
      stroke_ = ws.Range("A{}".format(starow)).Value
      weight_ = ws.Range("A{}".format(starow+1)).Value
      if stroke_ == '0-50cc':
        vehName = 'Motorcycle - 0-50cc'
      else:
        vehName = 'Motorcycle - {} - {}'.format(stroke_, weight_)
    elif mode == 'Hybrid Buses':
      decker_ = ws.Range("A{}".format(starow)).Value
      vehName = 'Hybrid Buses - {}'.format(decker_)
      starow += 1  # Grrrrr. Poor formatting in the EFT
      endrow += 1
    elif mode == 'Weights':
      vehName = ws.Range("{}{}".format(ColName, starow-1)).Value
    else:
      raise ValueError("mode '{}' is not recognised.".format(mode))
    for row in range(starow, endrow+1):
      euroName = ws.Range("{}{}".format(ColName, row)).Value
      if euroName is not None:
        sourceCell = "{}{}".format(ColProp, row)
        userCell = "{}{}".format(ColUser, row)
        proportion = ws.Range(sourceCell).Value
        if not isinstance(proportion, float):
          logprint(loggerM, 'Bad proportion value "{}" for veh {}, euro {}.'.format(proportion, vehName, euroName), level='info')
          sourceCell = "{}{}".format(ColUser, row)
          proportion = ws.Range(sourceCell).Value
          if not isinstance(proportion, float):
            #print(proportion)
            raise ValueError('Proportion must be a float.')
          else:
            logprint(loggerM, 'Fixed. Proportion value {}.'.format(proportion), level='info')
        logprint(loggerM, 'vehName: {}, euroName: {}, proportion: {}'.format(vehName, euroName, proportion), level='debug')
        got = False
        if mode == 'Weights':
          euroName = weightClassNameVariations[euroName]
          df1 = pd.DataFrame([[vehName, euroName, -99, '--', proportion, sourceCell, userCell]],
                               columns=['vehicle', 'euroname', 'euroclass',
                                        'technology', 'proportion', 'sourceCell', 'userCell'])
          df = df.append(df1, 1)
          continue
        for euroI, euronames in euroClassNameVariations.items():
          if euroI == 99:
            continue
          if euroName in euronames['All']:
            got = True
            tech = 'Standard'
            for techname, euronamestechs in euronames.items():
              if techname == 'All':
                continue
              if euroName in euronamestechs:
                tech = techname
                break
            df1 = pd.DataFrame([[vehName, euroName, euroI, tech, proportion, sourceCell, userCell]],
                               columns=['vehicle', 'euroname', 'euroclass',
                                        'technology', 'proportion', 'sourceCell', 'userCell'])
            df = df.append(df1, 1)

        if not got:
          raise ValueError("Can't identify euro class from {}.".format(euroName))
  if mode == 'Weights':
    df = df.rename(columns={'euroname': 'weightclass'})
    df = df.drop('euroclass', 1)
    df = df.drop('technology', 1)
  #print(df.head())
  return df


def readFleetProps(fname):
  """
  Read the fleet proportion file and return a dictionary with cell references
  and proportions to set.
  """
  props = {}
  if fname is None:
    return props
  try:
    # Assume excel document like the template.
    workbook = xlrd.open_workbook(fname)
    sheet = workbook.sheet_by_index(0)
    mode = 1
  except xlrd.biffh.XLRDError:
    # Is it a csv?
    sheet = pd.read_csv(fname) # Will raise an error if that fails.
    mode = 0

  Collect = False

  if mode == 0:
    csv_df = pd.read_csv(fname)
    for ii, row in csv_df.iterrows():
      if row['Cell'] != '---':
        props[row['Cell']] = row['Proportion']
  elif mode == 1:
    for rowID in range(sheet.nrows):
      # Find the "Default?" cells.
      if (sheet.row(rowID)[1].value == 'Default?') and (sheet.row(rowID)[2].value == 'No'):
        Collect = True
      elif sheet.row(rowID)[1].value == '':
        Collect = False
      if Collect:
        if sheet.row(rowID)[6].value != '':
          props[sheet.row(rowID)[6].value] = sheet.row(rowID)[1].value
        if sheet.row(rowID)[7].value != '':
          props[sheet.row(rowID)[7].value] = sheet.row(rowID)[4].value
  return props

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

def specifyEuroProportions(euroClass, workBook, vehRowStarts, vehRowEnds,
                 EuroClassNameColumns, DefaultEuroColumns, UserDefinedEuroColumns,
                 SubType=None, tech='All'):
  """
  Specify the euro class proportions.
  Will return the defualt proportions.
  """
  firstGot = {}
  defaultProps = {}
  #print("    Setting euro ratios to 100% for euro {}.".format(euroClass))
  ws_euro = workBook.Worksheets("UserEuro")
  for [vi, vehRowStart] in enumerate(vehRowStarts):
    if SubType == 'MC':
      vehNameA = ws_euro.Range("A{row}".format(row=vehRowStart)).Value
      vehNameB = ws_euro.Range("A{row}".format(row=vehRowStart+1)).Value
      if vehNameB is None:
        vehName = 'Motorcycle - {}'.format(vehNameA)
      else:
        vehName = 'Motorcycle - {} - {}'.format(vehNameA, vehNameB)
    elif SubType == 'HB':
      vehNameA = ws_euro.Range("A{row}".format(row=vehRowStart)).Value
      vehNameB = ws_euro.Range("A{row}".format(row=vehRowStart+1)).Value
      if vehNameB is None:
        vehName = 'Hybrid Bus - {}'.format(vehNameA)
      else:
        vehName = 'Hybrid Bus - {} - {}'.format(vehNameA, vehNameB)
    else:
      vehName = ws_euro.Range("A{row}".format(row=vehRowStart-1)).Value
    if vehName is None:
      vehName = ws_euro.Range("A{row}".format(row=vehRowStart)).Value
    firstGot[vehName] = [False, False]
    #print("      Setting euro ratios for {}.".format(vehName))
    vehRowEnd = vehRowEnds[vi]
    for [ci, euroNameCol] in enumerate(EuroClassNameColumns):
      first = True
      # There are only two columns, one for NOx, one for particulates.
      userDefinedCol = UserDefinedEuroColumns[ci]
      defaultEuroCol = DefaultEuroColumns[ci]

      # A quick fix to the hybrid bus problem! Grrrr!
      if ci == 1:
        if vehRowStart in [324, 328, 332]:
          vehRowStart = vehRowStart+1
          vehRowEnd = vehRowEnd+1

      euroClassRange = "{col}{rstart}:{col}{rend}".format(col=euroNameCol, rstart=vehRowStart, rend=vehRowEnd)

      # Get all of the euro class names for this vehicle.
      euroClassesAvailable = ws_euro.Range(euroClassRange).Value
      # Make sure we don't include trailing 'None' rows, by going backwards.
      euroClassesAvailableR = list(reversed(euroClassesAvailable))
      for eca in euroClassesAvailableR:
        #print(eca)
        if eca[0] is None:
          vehRowEnd = vehRowEnd - 1
        else:
          break
      # See which columns contain a line that specifies the required euro class.
      rowsToDo = []
      rowsToDoOther = []
      euroClass_ = euroClass
      euroClassp = euroClass
      while len(rowsToDo) == 0:
        if tech == 'All':
          euroSearchTerms_ = euroSearchTerms(euroClass_)
          euroSearchTerms_Other = []
        else:
          if tech in euroTechs(euroClass_):
            euroSearchTerms_ = euroSearchTerms(euroClass_, tech=tech)
            euroSearchTerms_Other = euroSearchTerms(euroClass_)
          else:
            euroSearchTerms_ = euroSearchTerms(euroClass_)
            euroSearchTerms_Other = []
        got, gotOther = False, False
        for [ei, name] in enumerate(euroClassesAvailable):
          name = name[0]
          if name in euroSearchTerms_:
            rowsToDo.append(vehRowStart + ei)
            got = True
          if name in euroSearchTerms_Other:
            rowsToDoOther.append(vehRowStart + ei)
            gotOther = True
        if got:
          if first:
            firstGot[vehName][ci] = True
        elif gotOther:
          rowsToDo = rowsToDoOther
        else:
          first = False
          euroClass_ -= 1
          if euroClass_ < 0:
            euroClassp += 1
            euroClass_ = euroClassp

          #print('      No values available for euro {}, trying euro {}.'.format(euroClass_o, euroClass_))
      #print('    found the following {}.'.format(', '.join([str(x) for x in rowsToDo])))
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
      try:
        ws_euro.Range(userRange).Value = 0
      except:
        for rn in range(vehRowStart, vehRowEnd+1):
          userRange = "{}{}".format(userDefinedCol, rn)
          try:
            ws_euro.Range(userRange).Value = 0
          except:
            pass

      # Then set the specific values.
      #print(rowsToDo)
      for [ri, row] in enumerate(rowsToDo):
        userRange = "{col}{row}".format(col=userDefinedCol, row=row)
        value = userProportions[ri]
        ws_euro.Range(userRange).Value = value
  #print('    All complete')
  return defaultProps, firstGot

def SpecifyWeight(workBook, start, end, do):
  ws_euro = workBook.Worksheets("UserEuro")
  ws_euro.Range("D{}:D{}".format(start, end)).Value = 0
  ws_euro.Range("D{}".format(do)).Value = 1

  wn = ws_euro.Range("A{}".format(do)).Value
  if wn in ['Single Decker', 'Double Decker', 'Articulated']:
    # Hybrid buses have the weight names on the wrong line.
    wn = ws_euro.Range("B{}".format(do)).Value
  p = re.compile('\d_*')
  m = p.match(wn)
  if m is not None:
    wn = wn[m.end():]
  return wn