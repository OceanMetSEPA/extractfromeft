# -*- coding: utf-8 -*-
"""
Created on Fri Apr  6 14:49:52 2018

@author: edward.barratt
"""

import os
import argparse
import numpy as np
import pandas as pd
from datetime import datetime


#import EFT_Tools as tools
WeightClasses = {}


def writeChanges(changes, saveloc):
  print("Saving changes to '{}'.".format(saveloc))
  with open(saveloc, 'w') as f:
    f.write('Created Using, fleetSplitFromANPR, -\n')
    f.write('Created On, {}, -\n'.format(datetime.now()))
    f.write('Vehicle Class, Cell, Proportion\n')
    for vehClass, PropsAll in changes.items():
      for cell, Propsin in PropsAll.items():
        WStr = '{}, {}, {}\n'.format(vehClass, cell, Propsin)
        #print(WStr)
        f.write(WStr)


def mergeVDicts(D):
  D_ = {}
  TotVehs = 0
  for Di in D:
    for key, value in Di.items():
      if key == 'Unknown':
        continue
      TotVehs += value['num']
      try :
        D_[key]['num'] += value['num']
      except KeyError:
        D_[key] = {'num': value['num']}
  for key, value in D_.items():
    vvv = D_[key]['num']/TotVehs
    D_[key]['normFract'] = vvv
    #print(key, vvv)
  return D_

def getchanges(ED, WD, eftE_veh, eftW_veh, verbose=False, vehName=''):
  changes = {}

  if verbose:
    print('')
    if vehName:
      print(vehName)
      print('-'*len(vehName))

  if isinstance(ED, list):
    # Add all the numbers together.
    ED = mergeVDicts(ED)
    if verbose:
      for euro, EDD in ED.items():
        print('Euro   {:20.0f}: {:6d} vehs, {:9.6f}%'.format(
            euro, EDD['num'], 100.*EDD['normFract']))

  if isinstance(WD, list):
    # Add all the numbers together.
    WD = mergeVDicts(WD)
    if verbose:
      for weight, WDD in WD.items():
        if weight != 'Unknown':
          print('Weight {:>20s}: {:6d} vehs, {:9.6f}%.'.format(
              weight, WDD['num'], 100.*WDD['normFract']))

  propTot = np.zeros_like(EFTPolTypes)
  lastCellName = ['-']*len(EFTPolTypes)
  for Euro in range(7):
    if Euro in ED.keys():
      V = ED[Euro]
    else:
      V = {'normFract': 0.0}
    eft_veh_euro = eftE_veh[eftE_veh['euroclass']==Euro]
    for polI, polType in enumerate(EFTPolTypes):
      # Get the correct values and cells for the given pollutant type'
      eft_veh_euro_p = eft_veh_euro[eft_veh_euro['poltype']==polType]
      eft_props = np.array(eft_veh_euro_p['proportion'])
      eft_cells = list(eft_veh_euro_p['userCell'])

      if len(eft_cells) == 1:
        # Just one value to fill.
        vvvv = round(V['normFract'], 8)
        changes[eft_cells[0]] = vvvv
        propTot[polI] += vvvv
        lastCellName[polI] = eft_cells[0]
      else:
        # A few cells to fill, so split the proportion we have between these
        # cells, weighted by the default split.
        if sum(eft_props) == 0:
          # All zeros, split evenly.
          eft_props_norm = np.ones_like(eft_props) / len(eft_props)
        else:
          eft_props_norm = eft_props / sum(eft_props)
        for ei, ec in enumerate(eft_cells):
          vvvv = round(eft_props_norm[ei] * V['normFract'], 8)
          propTot[polI] += vvvv
          changes[ec] = vvvv
          lastCellName[polI] = ec

  # Adjust the values to ensure they sum to 1.
  for pi, pt in enumerate(propTot):
    diff = 1.0 - pt
    if abs(diff) > 1e-15:
      if abs(diff) > 1e-7:
        print(pt)
        raise ValueError("Doesn't sum to 1!")
      print('Adjusting Euro Nums')
      changes[lastCellName[pi]] = round(changes[lastCellName[pi]] + diff, 8)


  # Weight
  EFTWeights = eftW_veh['weightclass'].unique()
  Weights = list(WD.keys())
  WeightsGot = dict.fromkeys(Weights, False)
  propTot = 0.0
  lastCellName = ''
  for EFTWeight in EFTWeights:
    # Get the default values from the EFT.
    eft_veh_weight = eftW_veh[eftW_veh['weightclass'] == EFTWeight]
    eft_props = np.array(eft_veh_weight['proportion'])
    eft_cells = list(eft_veh_weight['userCell'])

    # Assume 0 initially.
    #for eft_cell in eft_cells:
    changes[eft_cells[0]] = 0.0
    # Now add the values from the ANPR.
    if EFTWeight in Weights:
      V = WD[EFTWeight]
      # Just one value to fill for all weight classes.
      vvvv = round(V['normFract'], 8)
      #print(vvvv)
      changes[eft_cells[0]] = vvvv
      propTot += vvvv
      lastCellName = eft_cells[0]
      WeightsGot[EFTWeight] = True

  for W, B in WeightsGot.items():
    if (not B) and (W != 'Unknown'):
      raise ValueError("No value assigned for weight '{}'.".format(W))

  diff = 1.0 - propTot
  if abs(diff) > 1e-15:
    if abs(diff) > 1e-7:
      print(propTot)
      raise ValueError("Doesn't sum to 1!")
    print('Adjusting Weight Nums')
    print(propTot)
    print(diff)
    print(changes[lastCellName])
    changes[lastCellName] = round(changes[lastCellName] + diff, 8)
    print(changes[lastCellName])

  return changes


def getFromEFT(year, area, euroProportionsFile='default', weightProportionsFile='default'):
  """
  Returns the default proportions from the EFT. The

  """

  defaultDir = 'input'
  defaultDir = os.path.abspath(defaultDir)
  defaultEPF = os.path.join(defaultDir, 'AllCombined_AllEuroProportions.csv')
  defaultWPF = os.path.join(defaultDir, 'AllCombined_WeightProportions.csv')
  if euroProportionsFile == 'default':
    euroProportionsFile = defaultEPF
  if weightProportionsFile == 'default':
    weightProportionsFile = defaultWPF

  EData = pd.read_csv(euroProportionsFile)
  EData = EData[EData.area == area]
  EData = EData[EData.year == year]

  WData = pd.read_csv(weightProportionsFile)
  WData = WData[WData.area == area]
  WData = WData[WData.year == year]

  return EData, WData

def getBreakdown(data, colE, colW, verbose=False, vehName=''):
  """
  Returns the vehicle fleet breakdown for euro class and weight class from the
  data frame.
  """
  if verbose:
    print('')
    if vehName:
      print(vehName)
      print('-'*len(vehName))


  numTot = len(data.index)
  # Groupby Euro class.
  eurogroups = data.groupby([colE])
  euroDict = {}
  fractions = np.array([])
  for euro, group in eurogroups:
    numvehs = len(group.index)
    fraction = numvehs/numTot
    euroDict[euro] = {'num': numvehs, 'fraction': fraction}
    fractions = np.append(fractions, [fraction])
  fractionsS = sum(fractions)
  for ei, euro in enumerate(euroDict.keys()):
    ED = euroDict[euro]
    normFrac = np.round(ED['fraction']/fractionsS, 8)
    ED['normFract'] = normFrac
    if verbose:
      print('Euro   {:20.0f}: {:6d} vehs, {:9.6f}%, normalized to {:9.6f}%'.format(
            euro, ED['num'], 100.*ED['fraction'], 100.*ED['normFract']))

  # Groupby weight class.
  weightgroups = data.groupby([colW])
  weightDict = {}
  fractions = np.array([])
  for weight, group in weightgroups:
    numvehs = len(group.index)
    fraction = numvehs/numTot
    weightDict[weight] = {'num': numvehs, 'fraction': fraction}
    if weight != 'Unknown':
      fractions = np.append(fractions, [fraction])
  fractionsS = sum(fractions)

  for weight in weightDict.keys():
    WD = weightDict[weight]
    WD['normFract'] = WD['fraction']/fractionsS
    if weight != 'Unknown':
      if verbose:
        print('Weight {:>20s}: {:6d} vehs, {:9.6f}%, normalized to {:9.6f}%.'.format(
              weight, WD['num'], 100.*WD['fraction'], 100.*WD['normFract']))

  return euroDict, weightDict

if __name__ == '__main__':
  ProgDesc = ("Creates a vehFleetSplit file of the type used by shp2EFT using "
              "the contents of an ANPR data file.")
  ANPRDesc = ("The ANPR file should be a csv file listing all vehicles "
              "passing the ANPR counter (including double counting of vehicles "
              "that have passed more than once). There should be a column each "
              "for vehicle class, euro class, weight class and fuel.")
  parser = argparse.ArgumentParser(description=ProgDesc)
  parser.add_argument('anprfile', type=str,
                      help="The ANPR file to be processed. "+ANPRDesc)
  parser.add_argument('--saveloc', metavar='save location',
                      type=str, nargs='?',
                      help="Path where the outpt csv file should be saved.")
  parser.add_argument('--vehColumnName', metavar='vehicle class column name',
                      type=str, nargs='?', default='Vehicle11Split',
                      help="The column name for the vehicle class.")
  parser.add_argument('--weightColumnName', metavar='weight class column name',
                      type=str, nargs='?', default='WeightClassEFT',
                      help="The column name for the vehicle weight class.")
  parser.add_argument('--euroColumnName', metavar='euro class column name',
                      type=str, nargs='?', default='EuroClass',
                      help="The column name for the vehicle euro class.")
  parser.add_argument('--fuelColumnName', metavar='fuel column name',
                      type=str, nargs='?', default='Fuel',
                      help="The column name for the vehicle fuel.")

  args = parser.parse_args()
  anprfile = args.anprfile
  colV = args.vehColumnName
  colW = args.weightColumnName
  colE = args.euroColumnName
  colF = args.fuelColumnName
  saveloc = args.saveloc
  reqCols = [colV, colW, colE, colF]

  # Check that the anpr file exists.
  if not os.path.exists(anprfile):
    raise ValueError('File {} does not exist.'.format(anprfile))

  # Get the default proportions.
  EFTEuroDefault, EFTWeightDefault = getFromEFT(2018, 'Scotland')
  EFTPolTypes = EFTEuroDefault['poltype'].unique()

  # Read the file into pandas, but only keep the Euro class and the weight class
  # columns.
  data = pd.read_csv(anprfile, encoding = "ISO-8859-1")

  colnames = list(data)
  for q in reqCols:
    if q not in colnames:
      raise ValueError('Column {} does not exist in file.'.format(q))
  for col in colnames:
    if col not in reqCols:
      data = data.drop(col, 1)

  print(data[colV].unique())
  print(EFTEuroDefault['vehicle'].unique())
  print(EFTWeightDefault['vehicle'].unique())

  changes = {}

  # Cars
  data_cars = data[data[colV] == '2. CAR']
  # Diesel Cars
  vehName = 'Diesel Car'
  data_veh = data_cars[data_cars[colF] == 'HEAVY OIL']
  eftE_veh = EFTEuroDefault[EFTEuroDefault['vehicle'] == vehName]
  eftW_veh = EFTWeightDefault[EFTWeightDefault['vehicle'] == vehName]
  ED, WD = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)
  changes[vehName] = getchanges(ED, WD, eftE_veh, eftW_veh)

  # Petrol Cars
  vehName = 'Petrol Car'
  vehNameW = 'Petrol car'
  data_veh = data_cars[data_cars[colF] == 'PETROL']
  eftE_veh = EFTEuroDefault[EFTEuroDefault['vehicle'] == vehName]
  eftW_veh = EFTWeightDefault[EFTWeightDefault['vehicle'] == vehNameW]
  ED, WD = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)
  changes[vehName] = getchanges(ED, WD, eftE_veh, eftW_veh)


  # LGVs
  data_lgvs = data[data[colV] == '4. LGV']
  # Diesel LGVs
  vehName='Diesel LGV'
  data_veh = data_lgvs[data_lgvs[colF] == 'HEAVY OIL']
  eftE_veh = EFTEuroDefault[EFTEuroDefault['vehicle'] == vehName]
  eftW_veh = EFTWeightDefault[EFTWeightDefault['vehicle'] == vehName]
  ED, WD = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)
  changes[vehName] = getchanges(ED, WD, eftE_veh, eftW_veh)

  # Petrol LGVs
  vehName='Petrol LGV'
  data_veh = data_lgvs[data_lgvs[colF] == 'HEAVY OIL']
  eftE_veh = EFTEuroDefault[EFTEuroDefault['vehicle'] == vehName]
  eftW_veh = EFTWeightDefault[EFTWeightDefault['vehicle'] == vehName]
  ED, WD = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)
  changes[vehName] = getchanges(ED, WD, eftE_veh, eftW_veh)
  print(changes[vehName])
  # Buses
  vehName='Bus'
  vehName2 = 'Buses'
  data_veh = data[data[colV] == '5. BUS']
  #print(data_veh)
  eftE_veh = EFTEuroDefault[EFTEuroDefault['vehicle'] == vehName2]
  #print(eftE_veh)
  eftW_veh = EFTWeightDefault[EFTWeightDefault['vehicle'] == vehName2]
  ED, WD = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)
  changes[vehName] = getchanges(ED, WD, eftE_veh, eftW_veh)

  # RHGV 2X
  vehName='Rigid HGV 2 Axle'
  data_veh = data[data[colV] == '6a. RHGV_2X']
  ED2, WD2 = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)

  # RHGV 3X
  vehName='Rigid HGV 3 Axle'
  data_veh = data[data[colV] == '6b. RHGV_3X']
  ED3, WD3 = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)

  # RHGV 4X
  vehName='Rigid HGV 4 Axle'
  data_veh = data[data[colV] == '6c. RHGV_4X']
  ED4, WD4 = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)

  # Get changes for Rigid HGVs
  vehName2 = 'Rigid HGV'
  eftE_veh = EFTEuroDefault[EFTEuroDefault['vehicle'] == vehName2]
  eftW_veh = EFTWeightDefault[EFTWeightDefault['vehicle'] == vehName2]
  changes[vehName] = getchanges([ED2, ED3, ED4], [WD2, WD3, WD4], eftE_veh, eftW_veh,
                                verbose=True, vehName=vehName2)

  # AHGV 34X
  vehName='Artic HGV 3&4 Axle'
  data_veh = data[data[colV] == '7a. AHGV_34X']
  ED3, WD3 = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)

  # AHGV 5X
  vehName='Artic HGV 5 Axle'
  data_veh = data[data[colV] == '7b. AHGV_5X']
  ED5, WD5 = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)

  # AHGV 6X
  vehName='Artic HGV 6 Axle'
  data_veh = data[data[colV] == '7c. AHGV_6X']
  ED6, WD6 = getBreakdown(data_veh, colE, colW, verbose=True, vehName=vehName)

  # Get changes for Artic HGVs
  vehName2 = 'Artic HGV'
  eftE_veh = EFTEuroDefault[EFTEuroDefault['vehicle'] == vehName2]
  eftW_veh = EFTWeightDefault[EFTWeightDefault['vehicle'] == vehName2]
  changes[vehName] = getchanges([ED3, ED5, ED6], [WD3, WD5, WD6], eftE_veh, eftW_veh,
                                verbose=True, vehName=vehName2)


  print()

  if saveloc is None:
    saveloc = anprfile.replace('.csv', '_EFTProportionChanges.csv')
  writeChanges(changes, saveloc)