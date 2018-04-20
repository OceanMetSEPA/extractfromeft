# -*- coding: utf-8 -*-
"""
Created on Fri Apr 20 15:58:46 2018

@author: edward.barratt
"""

# Define some global variables. These may need to be augmented if a new EFT
# version is released.
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
versionDetails[7.0]['vehRowStartsHB'] = [377,381,385] # Hybrid buses
versionDetails[7.0]['vehRowEndsHB']   = [380,384,388]
versionDetails[7.0]['busCoachRow']   = [482, 483]
versionDetails[7.0]['SourceNameName'] = 'Source Name'
versionDetails[7.0]['AllLDVName'] = 'All LDVs (g/km)'
versionDetails[7.0]['AllHDVName'] = 'All HDVs (g/km)'
versionDetails[7.0]['AllVehName'] = 'All Vehicles (g/km)'
versionDetails[7.0]['PolName'] = 'Pollutant Name'
versionDetails[7.0]['weightRowStarts'] = [435, 440, 445, 450, 455, 465, 489,
                                          494, 499, 510, 515, 520, 525, 536,
                                          541, 546, 556]
versionDetails[7.0]['weightRowEnds'] =   [437, 442, 447, 452, 462, 469, 491,
                                          496, 501, 512, 517, 522, 527, 538,
                                          543, 553, 560]
versionDetails[7.0]['weightRowNames'] = ['Car', 'Car', 'LGV', 'LGV',
                                         'Rigid HGV', 'Artic HGV',  'Car',
                                         'Car', 'Car', 'Car', 'Car', 'LGV',
                                         'LGV', 'LGV', 'LGV', 'Rigid HGV',
                                         'Artic HGV']
versionDetails[7.0]['weightRowStartsBus'] = [472, 478, 565, 585, 590]
versionDetails[7.0]['weightRowEndsBus'] =   [474, 479, 567, 587, 591]
versionDetails[7.0]['weightRowNamesBus'] = ['Bus', 'Coach', 'Bus', 'Bus',
                                            'Coach']
versionDetails[6.0] = {}
versionDetails[6.0]['vehRowStarts'] = [69, 79, 100, 110, 123, 139, 155, 170]
versionDetails[6.0]['vehRowEnds'] = [75, 87, 106, 119, 134, 150, 166, 181]
versionDetails[6.0]['vehRowStartsMC'] = [186, 192, 198, 204, 210, 216]
versionDetails[6.0]['vehRowEndsMC']   = [191, 197, 203, 209, 215, 221]
versionDetails[6.0]['vehRowStartsHB'] = [377,381,385] # Hybrid buses
versionDetails[6.0]['vehRowEndsHB']   = [380,384,388]
versionDetails[6.0]['busCoachRow']   = [482, 483]
versionDetails[6.0]['SourceNameName'] = 'Source_Name'
versionDetails[6.0]['AllLDVName'] = 'All LDV (g/km)'
versionDetails[6.0]['AllHDVName'] = 'All HDV (g/km)'
versionDetails[6.0]['AllVehName'] = 'All Vehicle (g/km)'
versionDetails[6.0]['PolName'] = 'Pollutant_Name'
versionDetails[6.0]['weightRowStarts'] = [435, 440, 445, 450, 455, 465, 489,
                                          494, 499, 510, 515, 520, 525, 536,
                                          541, 546, 556]
versionDetails[6.0]['weightRowEnds'] =   [437, 442, 447, 452, 462, 469, 491,
                                          496, 501, 512, 517, 522, 527, 538,
                                          543, 553, 560]
versionDetails[6.0]['weightRowNames'] = ['Car', 'Car', 'LGV', 'LGV',
                                         'Rigid HGV', 'Artic HGV',  'Car',
                                         'Car', 'Car', 'Car', 'Car', 'LGV',
                                         'LGV', 'LGV', 'LGV', 'Rigid HGV',
                                         'Artic HGV']
versionDetails[6.0]['weightRowStartsBus'] = [472, 478, 565, 585, 590]
versionDetails[6.0]['weightRowEndsBus'] =   [474, 479, 567, 587, 591]
versionDetails[6.0]['weightRowNamesBus'] = ['Bus', 'Coach', 'Bus', 'Bus',
                                            'Coach']

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
UserDefinedEuroColumnsMC = ["D", "K"]
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
UserDefinedWeightColumn = "D"

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