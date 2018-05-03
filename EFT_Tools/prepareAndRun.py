# -*- coding: utf-8 -*-
"""
This file contains the function
prepareAndRun

Created on Fri Apr 20 15:10:47 2018

@author: edward.barratt
"""
import os
import subprocess
import time
import numpy as np
import pandas as pd
import win32com.client as win32
import pywintypes

from EFT_Tools import (checkEuroClassesValid,
                       createEFTInput,
                       DefaultBusColumn,
                       DefaultBusMWColumn,
                       DefaultEuroColumns,
                       DefaultEuroColumnsHB,
                       DefaultEuroColumnsMC,
                       DefaultLGVFuelColumns,
                       DefaultWeightColumn,
                       EuroClassNameColumns,
                       EuroClassNameColumnsHB,
                       EuroClassNameColumnsMC,
                       getLogger,
                       in2outVeh,
                       logprint,
                       numToLetter,
                       specifyBusCoach,
                       specifyEuroProportions,
                       SpecifyWeight,
                       techVehs,
                       UserDefinedBusColumn,
                       UserDefinedBusMWColumn,
                       UserDefinedEuroColumns,
                       UserDefinedEuroColumnsHB,
                       UserDefinedEuroColumnsMC,
                       UserDefinedLGVFuelColumns,
                       UserDefinedWeightColumn,
                       )

def pasteDefaultEuroProportions(ws, details, logger=None):
  # Get the logging details.
  loggerM = getLogger(logger, 'pasteDefaultEuroProportions')
  logprint(loggerM, '               Pasting default euro proportions.', level='info')

  # Bus vs Coach
  DefaultRange = '{colstart}{rowstart}:{colend}{rowend}'.format(
                                           colstart=DefaultBusColumn[0],
                                           colend=DefaultBusMWColumn[0],
                                           rowstart=details['busCoachRow'][0],
                                           rowend=details['busCoachRow'][1])
  UserRange = '{colstart}{rowstart}:{colend}{rowend}'.format(
                                           colstart=UserDefinedBusColumn[0],
                                           colend=UserDefinedBusMWColumn[0],
                                           rowstart=details['busCoachRow'][0],
                                           rowend=details['busCoachRow'][1])
  logprint(loggerM, '{} - {}'.format(DefaultRange, UserRange), level='debug')
  ws.Range(UserRange).Value = ws.Range(DefaultRange).Value

  # LGV Fule Types
  if 'lgvFuelRows' in details.keys():
    DefaultRange = '{colstart}{rowstart}:{colend}{rowend}'.format(
                                            colstart=DefaultLGVFuelColumns[0],
                                            colend=DefaultLGVFuelColumns[-1],
                                            rowstart=details['lgvFuelRows'][0],
                                            rowend=details['lgvFuelRows'][-1])
    UserRange = '{colstart}{rowstart}:{colend}{rowend}'.format(
                                            colstart=UserDefinedLGVFuelColumns[0],
                                            colend=UserDefinedLGVFuelColumns[-1],
                                            rowstart=details['lgvFuelRows'][0],
                                            rowend=details['lgvFuelRows'][-1])
    logprint(loggerM, '{} - {}'.format(DefaultRange, UserRange), level='debug')
    ws.Range(UserRange).Value = ws.Range(DefaultRange).Value

  # Sizes
  for rs, re in zip(details['weightRowStarts'], details['weightRowEnds']):
    DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=DefaultWeightColumn, rowstart=rs, rowend=re)
    UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=UserDefinedWeightColumn, rowstart=rs, rowend=re)
    #print('{} - {}'.format(DefaultRange, UserRange))
    ws.Range(UserRange).Value = ws.Range(DefaultRange).Value

  # Sizes Buses
  for rs, re in zip(details['weightRowStartsBus'], details['weightRowEndsBus']):
    if rs == 512: # Grumble grumbe grumble
      cd = 'C'
    else:
      cd = DefaultWeightColumn
    DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=cd, rowstart=rs, rowend=re)
    UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=UserDefinedWeightColumn, rowstart=rs, rowend=re)
    logprint(loggerM, '{} - {}'.format(DefaultRange, UserRange), level='debug')
    ws.Range(UserRange).Value = ws.Range(DefaultRange).Value

  # Hybrids
  for rs, re in zip(details['vehRowStartsHB'], details['vehRowEndsHB']):
    for qq, (d, u) in enumerate(zip(DefaultEuroColumnsHB, UserDefinedEuroColumnsHB)):
      DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=d, rowstart=rs+qq, rowend=re+qq)
      UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=u, rowstart=rs+qq, rowend=re+qq)
      logprint(loggerM, '{} - {}'.format(DefaultRange, UserRange), level='debug')
      ws.Range(UserRange).Value = ws.Range(DefaultRange).Value

  # Motorcycles
  for rs, re in zip(details['vehRowStartsMC'], details['vehRowEndsMC']):
    for d, u in zip(DefaultEuroColumnsMC, UserDefinedEuroColumnsMC):
      DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=d, rowstart=rs, rowend=re)
      UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=u, rowstart=rs, rowend=re)
      logprint(loggerM, '{} - {}'.format(DefaultRange, UserRange), level='debug')
      ws.Range(UserRange).Value = ws.Range(DefaultRange).Value

  # Standard euro class
  for rs, re in zip(details['vehRowStarts'], details['vehRowEnds']):

    for d, u in zip(DefaultEuroColumns, UserDefinedEuroColumns):
      DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=d, rowstart=rs, rowend=re)
      UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=u, rowstart=rs, rowend=re)
      logprint(loggerM, '{} - {}'.format(DefaultRange, UserRange), level='debug')
      try: # Nested tries, messy but works.
        ws.Range(UserRange).Value = ws.Range(DefaultRange).Value
      except pywintypes.com_error:
        # Try skipping the last row. This is neccesary with a few vehicle types.
        logprint(loggerM, 'skip Last', level='debug')
        DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=d, rowstart=rs, rowend=re-1)
        UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=u, rowstart=rs, rowend=re-1)
        try:
          ws.Range(UserRange).Value = ws.Range(DefaultRange).Value
        except pywintypes.com_error:
          # Try skipping the first row.
          # This is neccesary for the NOx column for B100 Rigid HGVs
          logprint(loggerM, 'skip First', level='debug')
          DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=d, rowstart=rs+1, rowend=re)
          UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=u, rowstart=rs+1, rowend=re)
          try:
            ws.Range(UserRange).Value = ws.Range(DefaultRange).Value
          except pywintypes.com_error:
            # Try skipping the first two row.
            # This is neccesary for the PM column for biodeiesel buses.
            logprint(loggerM, 'skip First Two', level='debug')
            DefaultRange = '{col}{rowstart}:{col}{rowend}'.format(col=d, rowstart=rs+2, rowend=re)
            UserRange = '{col}{rowstart}:{col}{rowend}'.format(col=u, rowstart=rs+2, rowend=re)
            ws.Range(UserRange).Value = ws.Range(DefaultRange).Value

def prepareAndRun(fileName, vehSplit, details, location, year, euroClass,
                  ahk_exepath, ahk_ahkpathG, versionForOutPut, excel=None,
                  checkEuroClasses=False, DoMCycles=True, DoHybridBus=True, DoBusCoach=False,
                  inputData='prepare', busCoach='default', sizeRow=99, tech='All', vehiclesToSkip=[], logger=None):
  """
  Prepare the file for running the macro.
  euroClass of 99 will retain default euro breakdown.
  """
  closeExcel = False
  if excel is None:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    closeExcel = True

  # Get the logging details.
  loggerM = getLogger(logger, 'prepareAndRun')

  # Start off the autohotkey script as a (parallel) subprocess. This will
  # continually check until the compatibility warning appears, and then
  # close the warning.
  if os.path.isfile(ahk_exepath):
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
      checkEuroClassesValid(wb, details['vehRowStartsMC'], details['vehRowEndsMC'], EuroClassNameColumnsMC, Type=1)
    if DoHybridBus:
      checkEuroClassesValid(wb, details['vehRowStartsHB'], details['vehRowEndsHB'], EuroClassNameColumnsMC, Type=2)
    checkEuroClassesValid(wb, details['vehRowStarts'], details['vehRowEnds'], EuroClassNameColumns, Type=0)

  # Set the default values in the Input Data sheet.
  ws_input = wb.Worksheets("Input Data")
  ws_input.Range("B4").Value = location
  ws_input.Range("B5").Value = year
  if ws_input.Range("B6").Value != vehSplit:
    # Don't change unless you have to.
    ws_input.Range("B6").Value = vehSplit


  if DoBusCoach:
    weightRowStarts = details['weightRowStartsBus']
    weightRowEnds = details['weightRowEndsBus']
    weightRowNames = details['weightRowNamesBus']
  else:
    weightRowStarts = details['weightRowStarts']
    weightRowEnds = details['weightRowEnds']
    weightRowNames = details['weightRowNames']

  if sizeRow != 99:
    # Set the designated size row.
    weightclassnames = {}
    for wri, wrs in enumerate(weightRowStarts):
      wre = weightRowEnds[wri]
      wrn = weightRowNames[wri]
      wrd = wrs + sizeRow
      if wrd > wre:
        # No more weight classes for this particular vehicle class.
        continue
      #Wally
      weightname = SpecifyWeight(wb, wrs, wre, wrd)
      weightclassnames[wrn] = weightname
  else:
    weightclassnames = {'Car': 'DefaultMix',
                        'LGV': 'DefaultMix',
                        'Rigid HGV': 'DefaultMix',
                        'Artic HGV': 'DefaultMix',
                        'Bus and Coach': 'DefaultMix',
                        'Bus': 'DefaultMix',
                        'Coach': 'DefaultMix',
                        'Motorcycle': 'DefaultMix'}

  logprint(loggerM, 'weightclassnames: {}'.format(weightclassnames), level='debug')
  vehsToInclude = []
  for veh, wn in weightclassnames.items():
    if (veh in techVehs[tech]) and (veh not in vehiclesToSkip):
      logprint(loggerM, '               Including weight "{}" for vehicles of class "{}."'.format(wn, veh))
      vehsToInclude.extend(in2outVeh[veh])
  for veh in vehiclesToSkip:
    if veh in vehsToInclude:
      vehsToInclude.remove(veh)

  if DoBusCoach:
    if busCoach == 'bus':
      if (sizeRow >= 3) and (sizeRow != 99):
        vehsToInclude = []
      else:
        vehsToInclude = ['Bus and Coach', 'B100 Bus', 'CNG Bus', 'Biomethane Bus',
                         'Biogas Bus', 'Hybrid Bus', 'FCEV Bus', 'B100 Coach']
    elif busCoach == 'coach':
      if (sizeRow >= 2) and (sizeRow != 99):
        vehsToInclude = []
      else:
        vehsToInclude = ['Bus and Coach', 'B100 Coach']

  if len(vehsToInclude) == 0:
    time.sleep(1) # To allow all systems to catch up.
    wb.Close(False)
    if closeExcel:
      excel.Quit()
      del(excelObj) # Make sure it's gone. Apparently some people have found this neccesary.
    return excel, None, None, None, None, None


  if type(inputData) is str:
    if inputData == 'prepare':
      # Prepare the input data.
      inputData = createEFTInput(vBreakdown=vehSplit, vehiclesToInclude=vehsToInclude, tech=tech, logger=logger)
      #inputData = inputData.as_matrix()
    else:
      raise ValueError("inputData '{}' is not understood.".format(inputData))
  if 'Motorcycle' not in vehsToInclude:
    DoMCycles = False
  if 'Hybrid Bus' not in vehsToInclude:
    DoHybridBus = False

  numRows, numCols = np.shape(inputData)
  inputData = tuple(map(tuple, inputData))
  ws_input.Range("A10:{}{}".format(numToLetter(numCols), numRows+9)).Value = inputData
  # Now we need to populate the UserEuro table with the defaults. Probably
  # only need to do this once per year, per area, but will do it every time
  # just in case.
  ws_euro = wb.Worksheets("UserEuro")
  ws_euro.Select()
  # There is a macro to do this, but for some reason it fails on versions 7.4
  # and 8.0 when run on my computer. So we must do it ourselves..
  pasteDefaultEuroProportions(ws_euro, details, logger=loggerM)
  #excel.Application.Run("PasteDefaultEuroProportions")

  # Now specify that we only want the specified euro class, by turning the
  # proportions for that class to 1, (or a weighted value if there are more
  # than one row for the particular euro class). This function also reads
  # the default proportions.
  if euroClass == 99:
    # Just stick with default euroclass.
    defaultProportions = 'NotMined'
    busCoachProportions = 'NotMined'
    gotTechs = None
    pass
  else:
    defaultProportions = pd.DataFrame(columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
    # Motorcycles first
    if DoMCycles:
      logprint(loggerM, '               Assigning fleet euro proportions for motorcycles.')
      defaultProportionsMC_, gotTechsMC = specifyEuroProportions(euroClass, wb,
                                  details['vehRowStartsMC'], details['vehRowEndsMC'],
                                  EuroClassNameColumnsMC, DefaultEuroColumnsMC,
                                  UserDefinedEuroColumns, SubType='MC', tech=tech)
      for key, value in defaultProportionsMC_.items():
        defaultProportionsRow= pd.DataFrame([[year, location, key, euroClass, value]],
                                             columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
        defaultProportions= defaultProportions.append(defaultProportionsRow)
    else:
      gotTechsMC = {}
    if DoHybridBus:
      logprint(loggerM, '               Assigning fleet euro proportions for hybrid buses.')
      defaultProportionsHB_, gotTechsHB = specifyEuroProportions(euroClass, wb,
                                  details['vehRowStartsHB'], details['vehRowEndsHB'],
                                  EuroClassNameColumnsHB, DefaultEuroColumnsHB,
                                  UserDefinedEuroColumns, SubType='HB', tech=tech)
      for key, value in defaultProportionsHB_.items():
        defaultProportionsRow= pd.DataFrame([[year, location, key, euroClass, value]],
                                             columns=['year', 'area', 'vehicle', 'euro', 'proportion'])
        defaultProportions= defaultProportions.append(defaultProportionsRow)
    else:
      gotTechsHB = {}
    logprint(loggerM, "               Assigning fleet euro proportions for all 'other' vehicle types.")
    # And all other vehicles
    defaultProportions_, gotTechs = specifyEuroProportions(euroClass, wb,
                             details['vehRowStarts'], details['vehRowEnds'],
                             EuroClassNameColumns, DefaultEuroColumns,
                             UserDefinedEuroColumns, tech=tech)

    gotTechs = {**gotTechs, **gotTechsMC, **gotTechsHB}
    for key, value in gotTechs.items():
      if any(value):
        gotTechs[key] = True
      else:
        gotTechs[key] = False
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
  logprint(loggerM, '               Running EFT routine.')

  excel.Application.Run("RunEfTRoutine")
  logprint(loggerM, '                 Complete.')
  time.sleep(0.5)

  # Save and Close. Saving as an xlsm, rather than a xlsb, file, so that it
  # can be opened by pandas.
  (FN, FE) =  os.path.splitext(fileName)
  if DoBusCoach:
    tempSaveName = fileName.replace(FE, '({}_{}_E{}_{}_{})'.format(location, year, euroClass, busCoach, sizeRow))
  else:
    tempSaveName = fileName.replace(FE, '({}_{}_E{}_{})'.format(location, year, euroClass, sizeRow))
  p = 1
  tempSaveName_ = tempSaveName
  while os.path.exists('{}.xlsm'.format(tempSaveName)):
    p += 1
    tempSaveName = '{}{}'.format(tempSaveName_, p)
  tempSaveName = '{}.xlsm'.format(tempSaveName)
  wb.SaveAs(tempSaveName, win32.constants.xlOpenXMLWorkbookMacroEnabled)
  wb.Close()

  time.sleep(1) # To allow all systems to catch up.
  if closeExcel:
    excel.Quit()
    del(excelObj) # Make sure it's gone. Apparently some people have found this neccesary.
  return excel, tempSaveName, defaultProportions, busCoachProportions, weightclassnames, gotTechs