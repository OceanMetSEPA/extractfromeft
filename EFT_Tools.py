from os import path
import os
import shutil
import re
import ast
import datetime, time
import numpy as np
import pandas as pd
import logging
import inspect
import random
import string
import subprocess
import win32com.client as win32
import pywintypes
#from fuzzywuzzy import process as fuzzyprocess

#homeDir = path.expanduser("~")
#logger = logging.getLogger(__name__)

# Define some global variables. These may need to be augmented if a new EFT
# version is released.
workingDir = os.getcwd()
ahk_exepath = 'C:\Program Files\AutoHotkey\AutoHotkey.exe'
ahk_ahkpath = 'closeWarning.ahk'

versionDetails = {}
versionDetails[7.4] = {}
versionDetails[7.4]['vehRowStarts'] = [ 69, 79, 91,101,114,130,146,161,218,226,
                                       231,242,252,261,267,278,288,296,310,353,
                                       367]
versionDetails[7.4]['vehRowEnds'] =   [ 76, 87, 98,110,125,141,157,172,223,229,
                                       234,249,257,264,270,285,293,305,319,364,
                                       378]
versionDetails[7.4]['vehRowStartsMC'] = [177,183,189,195,201,207]
versionDetails[7.4]['vehRowEndsMC']   = [182,188,194,200,206,212]
versionDetails[7.4]['vehRowStartsHB'] = [324,328,332] # Hybrid buses
versionDetails[7.4]['vehRowEndsHB']   = [327,331,335]
versionDetails[7.4]['busCoachRow']   = [429, 430]
versionDetails[7.4]['sizeRowStarts'] = {'LGV': [392, 397],
                                        'Rigid HGV': [402],
                                        'Artic HGV': [412]}
versionDetails[7.4]['sizeRowEnds'] = {'LGV': [394, 399],
                                      'Rigid HGV': [409],
                                      'Artic HGV': [416]}
versionDetails[7.4]['SourceNameName'] = 'Source Name'
versionDetails[7.4]['AllLDVName'] = 'All LDVs (g/km)'
versionDetails[7.4]['AllHDVName'] = 'All HDVs (g/km)'
versionDetails[7.4]['AllVehName'] = 'All Vehicles (g/km)'
versionDetails[7.4]['PolName'] = 'Pollutant Name'
versionDetails[8.0] = versionDetails[7.4]
versionDetails[8.0]['weightRowStarts'] = [382, 387, 392, 397, 402, 412, 436,
                                          441, 446, 457, 462, 467, 472, 483,
                                          488, 493, 503]
versionDetails[8.0]['weightRowEnds'] =   [384, 389, 394, 399, 409, 416, 438,
                                          443, 448, 459, 464, 469, 474, 485,
                                          490, 500, 507]
versionDetails[8.0]['weightRowNames'] = ['Car', 'Car', 'LGV', 'LGV',
                                         'Rigid HGV', 'Artic HGV',  'Car',
                                         'Car', 'Car', 'Car', 'Car', 'LGV',
                                         'LGV', 'LGV', 'LGV', 'Rigid HGV',
                                         'Artic HGV']
versionDetails[8.0]['weightRowStartsBus'] = [419, 425, 512, 532, 537]
versionDetails[8.0]['weightRowEndsBus'] =   [421, 426, 514, 534, 538]
versionDetails[8.0]['weightRowNamesBus'] = ['Bus', 'Coach', 'Bus', 'Bus',
                                            'Coach']
versionDetails[7.0] = {}
versionDetails[7.0]['vehRowStarts'] = [69, 79, 100, 110, 123, 139, 155, 170]
versionDetails[7.0]['vehRowEnds']   = [75, 87, 106, 119, 134, 150, 166, 181]
versionDetails[7.0]['vehRowStartsMC'] = [186, 192, 198, 204, 210, 216]
versionDetails[7.0]['vehRowEndsMC']   = [191, 197, 203, 209, 215, 221]
versionDetails[7.0]['busCoachRow']   = [482, 483]
versionDetails[7.0]['SourceNameName'] = 'Source Name'
versionDetails[7.0]['AllLDVName'] = 'All LDVs (g/km)'
versionDetails[7.0]['AllHDVName'] = 'All HDVs (g/km)'
versionDetails[7.0]['AllVehName'] = 'All Vehicles (g/km)'
versionDetails[7.0]['PolName'] = 'Pollutant Name'
versionDetails[6.0] = {}
versionDetails[6.0]['vehRowStarts'] = [69, 79, 100, 110, 123, 139, 155, 170]
versionDetails[6.0]['vehRowEnds'] = [75, 87, 106, 119, 134, 150, 166, 181]
versionDetails[6.0]['vehRowStartsMC'] = [186, 192, 198, 204, 210, 216]
versionDetails[6.0]['vehRowEndsMC']   = [191, 197, 203, 209, 215, 221]
versionDetails[6.0]['busCoachRow']   = [482, 483]
versionDetails[6.0]['SourceNameName'] = 'Source_Name'
versionDetails[6.0]['AllLDVName'] = 'All LDV (g/km)'
versionDetails[6.0]['AllHDVName'] = 'All HDV (g/km)'
versionDetails[6.0]['AllVehName'] = 'All Vehicle (g/km)'
versionDetails[6.0]['PolName'] = 'Pollutant_Name'
availableVersions = versionDetails.keys()
availableAreas = ['England (not London)', 'Northern Ireland',
                  'Scotland', 'Wales']
availableRoadTypes = ['Urban (not London)', 'Rural (not London)',
                      'Motorway (not London)']
availableModes = ['ExtractAll', 'ExtractCarRatio', 'ExtractBus']
availableEuros = [0,1,2,3,4,5,6]

weightClassNameVariations = {'1_<1400': '<1400 cc',
                             '2_1400-2000': '1400-2000 cc',
                             '3_>2000': '>2000 cc',
                             '1N1 (I)': 'N1 (I)',
                             '2N1 (II)': 'N1 (II)',
                             '3N1 (III)': 'N1 (III)',
                             '1_3.5-7.5 t': '3.5-7.5 t',
                             '2_7.5-12 t': '7.5-12 t',
                             '3_12-14 t': '12-14 t',
                             '4_14-20 t': '14-20 t',
                             '5_20-26 t': '20-26 t',
                             '6_26-28 t': '26-28 t',
                             '7_28-32 t': '28-32 t',
                             '8_>32 t': '>32 t',
                             '1_14-20 t': '14-20 t',
                             '2_20-28 t': '20-28 t',
                             '3_28-34 t': '28-34 t',
                             '4_34-40 t': '34-40 t',
                             '5_40-50 t': '40-50 t',
                             '1Urban Buses Midi <=15 t': '<=15 t',
                             '2Urban Buses Standard 15 - 18 t': '15-18 t',
                             '3Urban Buses Articulated >18 t': '>18 t',
                             '1Coaches Standard <=18 t': '<=18 t',
                             '2Coaches Articulated >18 t': '>18 t',
                             'Single Decker': '<=15 t',
                             'Double Decker': '15-18 t',
                             'Articulated': '>18 t'}

euroClassNameVariations = {}
euroClassNameVariations[0] = {'All': ['1Pre-Euro 1', '1Pre-Euro I',
                                      '1_Pre-Euro 1', '2Pre-Euro 1',
                                      '4Pre-Euro 1', '5Pre-Euro 1',
                                      '6Pre-Euro 1', '7Pre-Euro 1',
                                      '1_Pre-Euro 1'],
                              'Standard': ['1Pre-Euro 1', '1Pre-Euro I',
                                           '1_Pre-Euro 1', '2Pre-Euro 1',
                                           '4Pre-Euro 1', '5Pre-Euro 1',
                                           '6Pre-Euro 1', '7Pre-Euro 1',
                                           '1_Pre-Euro 1']}
euroClassNameVariations[1] = {'All': ['2Euro 1', '2Euro I', '1Euro 1',
                                      '2Euro 1', '2Euro 1', '4Euro 1',
                                      '5Euro 1', '6Euro 1', '7Euro 1',
                                      '9 Euro I DPFRF', '8Euro 1 DPFRF',
                                      '9Euro I DPFRF'],
                              'DPF': ['9 Euro I DPFRF', '8Euro 1 DPFRF',
                                      '9Euro I DPFRF'],
                              'Standard': ['2Euro 1', '2Euro I', '1Euro 1',
                                           '2Euro 1', '2Euro 1', '4Euro 1',
                                           '5Euro 1', '6Euro 1', '7Euro 1']}
euroClassNameVariations[2] = {'All': ['3Euro 2', '3Euro II', '1Euro 2',
                                      '2Euro 2', '2Euro 2', '4Euro 2',
                                      '5Euro 2', '6Euro 2', '7Euro 2',
                                      '10 Euro II DPFRF', '9Euro II SCRRF',
                                      '9Euro 2 DPFRF'],
                              'DPF': ['10 Euro II DPFRF', '9Euro 2 DPFRF'],
                              'SCR': ['9Euro II SCRRF'],
                              'Standard': ['3Euro 2', '3Euro II', '1Euro 2',
                                           '2Euro 2', '2Euro 2', '4Euro 2',
                                           '5Euro 2', '6Euro 2', '7Euro 2']}
euroClassNameVariations[3] = {'All': ['4Euro 3', '4Euro III', '1Euro 3',
                                      '2Euro 3', '2Euro 3', '4Euro 3',
                                      '5Euro 3', '6Euro 3', '7Euro 3',
                                      '11 Euro III DPFRF', '10Euro III SCRRF',
                                      '8Euro 3 DPF', '10Euro 3 DPFRF'],
                              'DPF': ['11 Euro III DPFRF', '8Euro 3 DPF',
                                      '10Euro 3 DPFRF'],
                              'SCR': ['10Euro III SCRRF'],
                              'Standard': ['4Euro 3', '4Euro III', '1Euro 3',
                                           '2Euro 3', '2Euro 3', '4Euro 3',
                                           '5Euro 3', '6Euro 3', '7Euro 3']}
euroClassNameVariations[4] = {'All': ['5Euro 4', '5Euro IV', '1Euro 4',
                                      '2Euro 4', '2Euro 4', '4Euro 4',
                                      '5Euro 4', '6Euro 4', '7Euro 4',
                                      '12 Euro IV DPFRF', '11Euro IV SCRRF',
                                      '9Euro 4 DPF'],
                              'DPF': ['12 Euro IV DPFRF', '9Euro 4 DPF'],
                              'SCR': ['11Euro IV SCRRF'],
                              'Standard': ['5Euro 4', '5Euro IV', '1Euro 4',
                                           '2Euro 4', '2Euro 4', '4Euro 4',
                                           '5Euro 4', '6Euro 4', '7Euro 4']}
euroClassNameVariations[5] = {'All': ['6Euro 5', '6Euro V', '1Euro 5',
                                      '2Euro 5',
                                      '2Euro 5', '4Euro 5', '5Euro 5',
                                      '6Euro 5', '7Euro 5', '7Euro V_SCR',
                                      '6Euro V_EGR', '12Euro V EGR + SCRRF'],
                              'EGR': ['6Euro V_EGR'],
                              'SCR': ['7Euro V_SCR'],
                              'EGR + SCRRF': ['12Euro V EGR + SCRRF'],
                              'Standard': ['6Euro 5', '6Euro V', '1Euro 5',
                                           '2Euro 5', '2Euro 5', '4Euro 5',
                                           '5Euro 5', '6Euro 5', '7Euro 5']}
euroClassNameVariations[6] = {'All': ['7Euro 6', '6Euro VI', '1Euro 6',
                                      '2Euro 6', '2Euro 6', '4Euro 6', '5Euro 6',
                                      '6Euro 6', '7Euro 6', '8Euro VI',
                                      '7Euro 6c', '7Euro 6d'],
                              'Standard': ['7Euro 6', '6Euro VI', '1Euro 6',
                                           '2Euro 6', '2Euro 6', '4Euro 6',
                                           '5Euro 6', '6Euro 6', '7Euro 6',
                                           '8Euro VI'],
                              'c': ['7Euro 6c'],
                              'd': ['7Euro 6d']}
euroClassNameVariations[99] = {'All': ['7Euro 6', '6Euro VI', '1Euro 6',
                                       '2Euro 6', '2Euro 6', '4Euro 6',
                                       '5Euro 6', '6Euro 6', '7Euro 6',
                                       '8Euro VI', '7Euro 6c', '7Euro 6d']}

euroClassNameVariationsIgnore = ['B100 Rigid HGV', 'Biodiesel Buses',
                                 'Biodiesel Coaches', 'Hybrid Buses',
                                 'Biodiesel Buses', 'Biodiesel Coaches']

AllowedTechs = {'LGV': ['c', 'd', 'Standard', 'DPF'],
                'Rigid HGV': ['Standard', 'DPF', 'EGR', 'SCR', 'EGR + SCRRF'],
                'Artic HGV': ['Standard', 'DPF', 'EGR', 'SCR', 'EGR + SCRRF']}

VehSplits = {'Basic Split': ['HDV'],
             'Detailed Option 1': ['Car', 'Taxi (black cab)', 'LGV', 'HGV',
                                   'Bus and Coach', 'Motorcycle'],
             'Detailed Option 2': ['Car', 'Taxi (black cab)', 'LGV', 'Rigid HGV',
                                   'Artic HGV', 'Bus and Coach', 'Motorcycle'],
             'Detailed Option 3': ['Petrol Car', 'Diesel Car',
                                   'Taxi (black cab)', 'LGV', 'Rigid HGV',
                                   'Artic HGV', 'Bus and Coach', 'Motorcycle'],
             'Alternative Technologies': ['Petrol Car', 'Diesel Car',
                                   'Taxi (black cab)', 'LGV', 'Rigid HGV',
                                   'Artic HGV', 'Bus and Coach', 'Motorcycle',
                                   'Full Hybrid Petrol Cars',
                                   'Plug-In Hybrid Petrol Cars',
                                   'Full Hybrid Diesel Cars', 'Battery EV Cars',
                                   'FCEV Cars', 'E85 Bioethanol Cars',
                                   'LPG Cars', 'Full Hybrid Petrol LGV',
                                   'Plug-In Hybrid Petrol LGV',
                                   'Battery EV LGV', 'FCEV LGV',
                                   'E85 Bioethanol LGV', 'LPG LGV',
                                   'B100 Rigid HGV', 'B100 Artic HGV', 'B100 Bus',
                                   'CNG Bus', 'Biomethane Bus', 'Biogas Bus',
                                   'Hybrid Bus', 'FCEV Bus', 'B100 Coach']}
AllVehs = []
for val in VehSplits.values():
  AllVehs.extend(val)
AllVehs = list(set(AllVehs))
allowedVehSplits = list(VehSplits.keys())

techVehs = {'DPF': ['Car', 'Diesel Car', 'Taxi (black cab)', 'LGV', 'HGV',
                    'Rigid HGV', 'Artic HGV', 'Bus and Coach', 'B100 Rigid HGV',
                    'B100 Artic HGV', 'B100 Bus' 'B100 Coach', 'Bus', 'Coach'],
            'SCR': ['HGV', 'Rigid HGV', 'Artic HGV', 'Bus and Coach', 'B100 Rigid HGV',
                    'B100 Artic HGV', 'B100 Bus', 'Hybrid Bus', 'B100 Coach', 'Bus', 'Coach'],
            'EGR': ['HGV', 'Rigid HGV', 'Artic HGV', 'Bus and Coach', 'B100 Rigid HGV',
                    'B100 Artic HGV', 'B100 Bus', 'Hybrid Bus', 'B100 Coach', 'Bus', 'Coach'],
            'EGR + SCRRF': ['HGV', 'Rigid HGV', 'Artic HGV', 'Bus and Coach', 'B100 Rigid HGV',
                    'B100 Artic HGV', 'B100 Bus', 'B100 Coach', 'Bus', 'Coach'],
            'c': ['Car', 'Petrol Car', 'Diesel Car', 'Taxi (black cab)', 'LGV',
                  'Full Hybrid Petrol Cars', 'Plug-In Hybrid Petrol Cars',
                  'Full Hybrid Diesel Cars', 'E85 Bioethanol Cars', 'Full Hybrid Petrol LGV',
                  'Plug-In Hybrid Petrol LGV', 'E85 Bioethanol LGV'],
            'd': ['Car', 'Diesel Car', 'Taxi (black cab)', 'LGV',
                  'Full Hybrid Diesel Cars'],
            'Standard': AllVehs,
            'All': AllVehs}

in2outVeh = {'Car': ['Petrol Car', 'Diesel Car', 'Full Hybrid Petrol Cars',
                     'Plug-In Hybrid Petrol Cars', 'Full Hybrid Diesel Cars',
                     'E85 Bioethanol Car', 'LPG Car'],
             'Petrol Car': ['Petrol Car'],
             'Diesel Car': ['Diesel Car'],
             'LGV': ['LGV', 'Petrol LGV', 'Diesel LGV', 'Full Hybrid Petrol LGV',
                     'Plug-In Hybrid Petrol LGV', 'E85 Bioethanol LGV', 'LPG LGV'],
             'HGV': ['Rigid HGV', 'B100 Rigid HGV', 'Artic HGV', 'B100 Artic HGV'],
             'Rigid HGV': ['Rigid HGV', 'B100 Rigid HGV'],
             'Artic HGV': ['Artic HGV', 'B100 Artic HGV'],
             'Bus and Coach': ['Buses (Not London Buses)', 'Coaches', 'Biodiesel Buses',
                               'Biodiesel Coaches'],
             'Motorcycle': ['Motorcycle - 0-50cc', 'Motorcycle - 2-stroke - 50-100cc',
                            'Motorcycle - 4-stroke - 50-150cc', 'Motorcycle - 4-stroke - 150-250cc',
                            'Motorcycle - 4-stroke - 250-750cc', 'Motorcycle - 4-stroke - >750-cc'],
             'Full Hybrid Petrol Cars': ['Full Hybrid Petrol Car'],
             'Plug-In Hybrid Petrol Cars': ['Plugin Hybrid Petrol Car'],
             'Full Hybrid Diesel Cars': ['Full Diesel Hybrid Car'],
             'E85 Bioethanol Cars': ['E85 Bioethanol Car'],
             'LPG Cars': ['LPG Car'],
             'Full Hybrid Petrol LGV': ['Full Hybrid Petrol LGV'],
             'Plug-In Hybrid Petrol LGV': ['Plug-In Hybrid Petrol LGV'],
             'E85 Bioethanol LGV': ['E85 Bioethanol LGV'],
             'LPG LGV': ['LPG LGV'],
             'B100 Rigid HGV': ['B100 Rigid HGV'],
             'B100 Artic HGV': ['B100 Artic HGV'],
             'B100 Bus': ['Biodiesel Buses'],
             'B100 Coach': ['Biodiesel Coaches'],
             'Bus': ['Buses (Not London Buses)', 'Biodiesel Buses'],
             'Coach': ['Coaches', 'Biodiesel Coaches', 'Coach'],
             'Hybrid Bus': ['Hybrid Bus', 'Hybrid Bus - Single Decker', 'Hybrid Bus - Double Decker', 'Hybrid Bus - Articulated']}


euroClassNameVariationsAll = euroClassNameVariations[0]['All'][:]
euroClassTechnologies = list(euroClassNameVariations[0].keys())
for ei in range(1,7):
  euroClassNameVariationsAll.extend(euroClassNameVariations[ei]['All'])
  euroClassTechnologies.extend(euroClassNameVariations[ei].keys())
euroClassNameVariationsAll = list(set(euroClassNameVariationsAll))
euroClassTechnologies = list(set(euroClassTechnologies))
euroClassTechnologies.remove('All')
euroClassTechnologies.sort()
#print(euroClassTechnologies)

EuroClassNameColumns = ["A", "H"]
DefaultEuroColumns = ["B", "I"]
UserDefinedEuroColumns = ["D", "K"]
EuroClassNameColumnsMC = ["B", "H"]
EuroClassNameColumnsHB = EuroClassNameColumnsMC
DefaultEuroColumnsMC = ["C", "I"]
DefaultEuroColumnsHB = DefaultEuroColumnsMC
UserDefinedBusColumn = ["D"]
UserDefinedBusMWColumn = ["E"]
DefaultBusColumn = ["B"]
DefaultBusMWColumn = ["C"]
NameWeightColumn = "A"
DefaultWeightColumn = "B"

VehDetails = {'Petrol Car': {'Fuel': 'Petrol', 'Veh': 'Car', 'Tech': 'Internal Combustion', 'NOxVeh': 'Petrol Car'},
              'Diesel Car': {'Fuel': 'Diesel', 'Veh': 'Car', 'Tech': 'Internal Combustion', 'NOxVeh': 'Diesel Car'},
              'Taxi (black cab)': {'Fuel': 'Diesel', 'Veh': 'Car', 'Tech': 'Internal Combustion', 'NOxVeh': 'Diesel Car'},
              'LGV': {'Fuel': 'Diesel', 'Veh': 'LGV', 'Tech': 'Internal Combustion', 'NOxVeh': 'Diesel LGV'}, # LGV cannot be split between Petrol and diesel (or I haven't figured out a way yet), but the vast majority are Diesel, so use that.
              'Rigid HGV': {'Fuel': 'Diesel', 'Veh': 'Rigid HGV', 'Tech': 'Internal Combustion', 'NOxVeh': 'Rigid HGV'},
              'Artic HGV': {'Fuel': 'Diesel', 'Veh': 'Artic HGV', 'Tech': 'Internal Combustion', 'NOxVeh': 'Artic HGV'},
              'Bus and Coach': {'Fuel': 'Diesel', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'},
              'Bus': {'Fuel': 'Diesel', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'},
              'Coach': {'Fuel': 'Diesel', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'},
              'Motorcycle': {'Fuel': 'Petrol', 'Veh': 'Motorcycle', 'Tech': 'Internal Combustion', 'NOxVeh': 'Motorcycle'},
              'Full Hybrid Petrol Cars': {'Fuel': 'Petrol', 'Veh': 'Car', 'Tech': 'Full Hybrid', 'NOxVeh': 'Petrol Car'},
              'Plug-In Hybrid Petrol Cars': {'Fuel': 'Petrol', 'Veh': 'Car', 'Tech': 'Plug In Hybrid', 'NOxVeh': 'Petrol Car'},
              'Full Hybrid Diesel Cars': {'Fuel': 'Diesel', 'Veh': 'Car', 'Tech': 'Full Hybrid', 'NOxVeh': 'Diesel Car'},
              'Battery EV Cars': {'Fuel': 'Electric', 'Veh': 'Car', 'Tech': 'Battery', 'NOxVeh': 'Petrol Car'},  # Well, it probably doesn't matter.
              'FCEV Cars': {'Fuel': 'Electric', 'Veh': 'Car', 'Tech': 'Fuel Cell', 'NOxVeh': 'Petrol Car'},
              'E85 Bioethanol Cars': {'Fuel': 'Bioethanol', 'Veh': 'Car', 'Tech': 'Internal Combustion', 'NOxVeh': 'Diesel Car'},
              'LPG Cars': {'Fuel': 'LPG', 'Veh': 'Car', 'Tech': 'Internal Combustion', 'NOxVeh': 'Petrol Car'},
              'Full Hybrid Petrol LGV': {'Fuel': 'Petrol', 'Veh': 'LGV', 'Tech': 'Full Hybrid', 'NOxVeh': 'Petrol LGV'},
              'Plug-In Hybrid Petrol LGV': {'Fuel': 'Petrol', 'Veh': 'LGV', 'Tech': 'Plug In Hybrid', 'NOxVeh': 'Petrol LGV'},
              'Battery EV LGV': {'Fuel': 'Electric', 'Veh': 'LGV', 'Tech': 'Battery', 'NOxVeh': 'Petrol LGV'},
              'FCEV LGV': {'Fuel': 'Electric', 'Veh': 'LGV', 'Tech': 'Fuel Cell', 'NOxVeh': 'Petrol LGV'},
              'E85 Bioethanol LGV': {'Fuel': 'Bioethanol', 'Veh': 'LGV', 'Tech': 'Internal Combustion', 'NOxVeh': 'Diesel LGV'},
              'LPG LGV': {'Fuel': 'LPG', 'Veh': 'LGV', 'Tech': 'Internal Combustion', 'NOxVeh': 'Petrol LGV'},
              'B100 Rigid HGV': {'Fuel': 'Biodiesel', 'Veh': 'Rigid HGV', 'Tech': 'Internal Combustion', 'NOxVeh': 'Rigid HGV'},
              'B100 Artic HGV': {'Fuel': 'Biodiesel', 'Veh': 'Artic HGV', 'Tech': 'Internal Combustion', 'NOxVeh': 'Artic HGV'},
              'B100 Bus': {'Fuel': 'Biodiesel', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'},
              'CNG Bus': {'Fuel': 'Compressed Natural Gas', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'},
              'Biomethane Bus': {'Fuel': 'Biomethane', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'},
              'Biogas Bus': {'Fuel': 'Biogas', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'},
              'Hybrid Bus': {'Fuel': 'Diesel', 'Veh': 'Bus and Coach', 'Tech': 'Full Hybrid', 'NOxVeh': 'Bus and Coach'},
              'FCEV Bus': {'Fuel': 'Electric', 'Veh': 'Bus and Coach', 'Tech': 'Fuel Cell', 'NOxVeh': 'Bus and Coach'},
              'B100 Coach': {'Fuel': 'Biodiesel', 'Veh': 'Bus and Coach', 'Tech': 'Internal Combustion', 'NOxVeh': 'Bus and Coach'}}

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

#def readLGVFuelSplit(File='input/NAEI_FuelSplitExtracted.xlsx'):
#  """
#  Function that reads the#
#
#  """
#  FactorsDF = pd.read_excel(File, sheetname='LGV')
#  Factors = {}
#  Years = list(FactorsDF)
#  Years.remove('Road type')
#  Years.remove('Fuel')
#  RoadTypes = FactorsDF['Road type'].unique()
#  Fuels = FactorsDF['Fuel'].unique()
#  for rt in RoadTypes:
#    Factors[rt] = {}
#    FFs = FactorsDF[FactorsDF['Road type'] == rt]
#    for Fuel in Fuels:
#      for Y in Years:
#        Factors[rt][Y] = {}
#        Factors[rt][Y]['Petrol'] = float(FFs[FFs['Fuel'] == 'Petrol'][Y])
#        Factors[rt][Y]['Diesel'] = float(FFs[FFs['Fuel'] == 'Diesel'][Y])
#  return Factors

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

def addNO2(dataframe, Factors='input/NAEI_NO2Extracted.xlsx', mode='Average'):
  """
  Function that adds NO2 emission factors to a data frame that already has NOx
  emission factors.

  The original data frame must have one column called 'NOx (g/km/veh)',
  one column called 'year', and one column called 'vehicle'.
  """

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
  elif version == 8.0:
    vPart = 'EFT2017_v8.0'
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

def extractVersion(fileName, availableVersions=[6.0, 7.0, 7.4, 8.0]):
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

def randomString(N = 6):
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
          if veh in vehiclesToSkip:
            logprint(loggerM, '      skipped', level='debug')
            pass
          else:
            logprint(loggerM, '      included', level='debug')
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
            inputDF.set_value(ri, len(VehSplit)+4, 1) # 1 hour. Not neccesary for g/km output.
            inputDF.set_value(ri, len(VehSplit)+5, 1) # 1 km. Not neccesary either.
  inputData = inputDF.as_matrix()
  return inputData

def logprint(logger, string, level='info'):
  if level.lower() == 'info':
    logfunc = lambda x: logger.info(x)
  elif level.lower() == 'debug':
    logfunc = lambda x: logger.debug(x)
  if logger is not None:
    logfunc(string)
  else:
    print(string)

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

def euroTechs(N):
  return euroClassNameVariations[N].keys()

def euroSearchTerms(N, tech='All'):
  ES = euroClassNameVariations[N][tech]
  return ES

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
    print('The Autohotkey executable file {} could not be found.'.format(ahk_exepath))
    ahk_ahkpathGot = None
  if not path.isfile(ahk_ahkpath):
    ahk_ahkpath_ = workingDir + '\\' + ahk_ahkpath
    if not path.isfile(ahk_ahkpath_):
      print('The Autohotkey file {} could not be found.'.format(ahk_ahkpath))
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
  output = output.drop(details['SourceNameName'], 1)
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
  output = output.pivot_table(index=['year', 'area', 'euro', 'version',
                                     'speed', 'vehicle', 'type', 'mitigation tech'],
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
  return output

def compareArgsEqual(newargs, logfilename):
  searchStr = 'Input arguments parsed as: '
  with open(logfilename, 'r') as f:
    for line in f:
      # We want the last set of commands.
      if searchStr in line:
        oldargs = line[line.find(searchStr)+len(searchStr):-1]
  # Check that they are equal
  oldargs = ast.literal_eval(oldargs)
  newargs = vars(newargs)

  if oldargs != newargs:
    print('')
    print(('You are attempting to continue evaluation based on a different set '
           'of input arguments:'))
    for key in oldargs.keys():
      print('Old: {}, {}'.format(key, oldargs[key]))
      print('New: {}, {}'.format(key, newargs[key]))

    Cont = input('Do you wish to continue. [y/n]')
    if Cont.lower() in ['yes', 'y']:
      pass
    else:
      exit()

def prepareLogger(loggerName, logfilename, pargs, inString):
  loggerName = __name__
  logger = logging.getLogger(loggerName)
  if pargs.loggingmode == 'INFO':
    logger.setLevel(logging.INFO)
  elif pargs.loggingmode == 'DEBUG':
    logger.setLevel(logging.DEBUG)
  else:
    raise ValueError("Logging mode '{}' not understood.".format(pargs.loggingmode))

  fileFormatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
  streamFormatter = logging.Formatter('%(asctime)s - %(message)s')
  logfilehandler = logging.FileHandler(logfilename)
  logfilehandler.setFormatter(fileFormatter)
  logstreamhandler = logging.StreamHandler()
  logstreamhandler.setFormatter(streamFormatter)
  logger.addHandler(logfilehandler)
  logger.addHandler(logstreamhandler)

  logger.info('Program started with command: "{}"'.format(inString))
  logger.info('Input arguments parsed as: {}'.format(vars(pargs)))
  return logger

def getLogger(logger, modName):
  if logger is None:
    return None
  else:
    if 'EFT_Tools' in logger.name:
      return logger.getChild(modName)
    else:
      return logger.getChild('EFT_Tools.{}'.format(modName))

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

  # Euro split!
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
  first = True
  for ci, poltype in enumerate(['NOx', 'PM']):
    for ri in range(len(startrows)):
      logprint(loggerM, '  Dealing with euro proportions for {} - {}.'.format(vehtypes[ri], poltype), level='info')
      ColName = EuroClassNameColumnsDict[vehtypes[ri]][ci]
      ColProp = DefaultEuroColumnsDict[vehtypes[ri]][ci]
      vehRowStarts = details[startrows[ri]]
      vehRowEnds = details[endrows[ri]]
      propdf = getProportions(ws_euro, ColName, ColProp, vehRowStarts,
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
    vehRowStarts = details[startrows[ri]]
    vehRowEnds = details[endrows[ri]]
    propdf = getProportions(ws_euro, ColName, ColProp, vehRowStarts,
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

def getProportions(ws, ColName, ColProp, vehRowStarts,
                   vehRowEnds, mode='Most Vehicles', logger=None):

    # Get the logging details.
  loggerM = getLogger(logger, 'getProportions')

  # Start a pandas dateframe.
  df = pd.DataFrame(columns=['vehicle', 'euroname',
                             'euroclass', 'technology', 'proportion'])
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
        proportion = ws.Range("{}{}".format(ColProp, row)).Value
        if not isinstance(proportion, float):
          logprint(loggerM, 'Bad proportion value "{}" for veh {}, euro {}.'.format(proportion, vehName, euroName), level='info')
          proportion = ws.Range("D{}".format(row)).Value
          if not isinstance(proportion, float):
            print(proportion)
            raise ValueError('Proportion must be a float.')
          else:
            logprint(loggerM, 'Fixed. Proportion value {}.'.format(proportion), level='info')
        logprint(loggerM, 'vehName: {}, euroName: {}, proportion: {}'.format(vehName, euroName, proportion), level='debug')
        got = False
        if mode == 'Weights':
          euroName = weightClassNameVariations[euroName]
          df1 = pd.DataFrame([[vehName, euroName, -99, '--', proportion]],
                               columns=['vehicle', 'euroname',
                                        'euroclass', 'technology', 'proportion'])
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
            df1 = pd.DataFrame([[vehName, euroName, euroI, tech, proportion]],
                               columns=['vehicle', 'euroname',
                                        'euroclass', 'technology', 'proportion'])
            df = df.append(df1, 1)

        if not got:
          raise ValueError("Can't identify euro class from {}.".format(euroName))
  if mode == 'Weights':
    df = df.rename(columns={'euroname': 'weightclass'})
    df = df.drop('euroclass', 1)
    df = df.drop('technology', 1)
  return df




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
  if path.isfile(ahk_exepath):
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
  wb.Worksheets("UserEuro").Select()
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
  (FN, FE) =  path.splitext(fileName)
  if DoBusCoach:
    tempSaveName = fileName.replace(FE, '({}_{}_E{}_{}_{})'.format(location, year, euroClass, busCoach, sizeRow))
  else:
    tempSaveName = fileName.replace(FE, '({}_{}_E{}_{})'.format(location, year, euroClass, sizeRow))
  p = 1
  tempSaveName_ = tempSaveName
  while path.exists('{}.xlsm'.format(tempSaveName)):
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

def combineFiles(directory):
  """
  Combine the created files within a directory. It works by reading the
  directory's log file and reads those files that have been completed.
  """

  # Does the directory exist?
  if not path.isdir(directory):
    raise ValueError('Directory cannot be found.')
  # Search for the log file.
  contents = os.listdir(directory)
  if len(contents) == 0:
    raise ValueError('Directory is empty.')
  go = False
  for content in contents:
    if content[-4:] == '.log':
      logfilename = path.join(directory, content)
      yn = input(('Combine files listed as completed '
                  'in log file {}? [y/n]'.format(logfilename)))
      if yn.upper() != 'Y':
        continue
      else:
        go = True
        break

  first = True
  if go:
    [fname, ext] = os.path.splitext(logfilename)
    fnew = fname+'_combined.csv'
    completed = getCompletedFromLog(logfilename)
    filenames = list(completed['saveloc'])

    for fni, fn in enumerate(filenames):
      if first:
        shutil.copyfile(fn, fnew)
        first = False
      else:
        df = pd.read_csv(fn)
        df.to_csv(fnew, mode='a', header=False, index=False)
  return fnew


def getCompletedFromLog(logfilename, mode='completed'):
  """
  Read the log file to see if any combination of location, year, euroclass,
  and tech have already been completed. Returns completed parameters in a
  pandas dataframe. Can also return combinations marked as skipped.

  logfilename should be the path to a log file created by extractEFT.py
  mode can be 'completed', 'skipped', or 'both'.
  """

  CompletedSearchStr = 'COMPLETED (area, year, euro, tech, saveloc): '
  SkippedSearchStr = 'SKIPPED (area, year, euro, tech, saveloc): '
  ProportionsSearchStr = 'COMPLETED (area, year): '
  if mode == 'completed':
    SearchStrs = [CompletedSearchStr]
  elif mode == 'skipped':
    SearchStrs = [SkippedSearchStr]
  elif mode == 'both':
    SearchStrs = [CompletedSearchStr, SkippedSearchStr]
  elif mode == 'proportions':
    SearchStrs = [ProportionsSearchStr]
  else:
    raise ValueError("mode '{}' is not understood.".format(mode))


  completed = pd.DataFrame(columns=['area', 'year', 'euro', 'tech', 'saveloc'])
  ci = 0
  with open(logfilename, 'r') as logf:
    for line in logf:


      for SearchStr in SearchStrs:
        if SearchStr in line:
          ci += 1
          info = line[line.find(SearchStr)+len(SearchStr):-2]
          infosplt = info.split(',')
          completed.loc[ci] = [infosplt[0].strip(), int(infosplt[1].strip()),
                               int(infosplt[2].strip()), infosplt[3].strip(),
                               infosplt[4].strip()]
          break
  return completed

if __name__ == '__main__':
  # For testing.

  aa = createEFTInput(vBreakdown='Detailed Option 2')
  print(aa.head(30))
