# -*- coding: utf-8 -*-
"""
Created on Tue Sep 05 15:17:15 2017

This script converts the 14 vehicle split that is used within the EFT into a 10
vehicle split. So that, for example, the proportion of vehicles assigned to
"10HybridCarPetrol" are added to "1Petrol car". It also renames the vehicle
classes so that instead of "1Petrol car" we have "2a. Petrol Cars".

The following changes are made.

1Petrol car               - 2a. Petrol Cars
2Diesel car               - 2b. Diesel Cars
3Taxi (black cab)         - 2c. Taxis (black cab)
4Petrol LGV               - 3a. Petrol LGVs
5Diesel LGV               - 3b. Diesel LGVs
6Rigid                    - 5a. Rigid HGVs
7Artic                    - 5b. Artic HGVs
8Bus and coach            - Split between 4a. Buses and 4b. Coaches, based on
                            whether road is motorway or not. (Controlled
                            through the dictionary PropBuses.)
9Motorcycle               - 1. Motorcycles
10HybridCarPetrol         - 2a. Petrol Cars
11PlugInHybridCarPetrol   - 2a. Petrol Cars
12HybridCarDiesel         - 2b. Diesel Cars
13ElectricCar             - Split between 2a and 2b based on existing 2a-2b split.
14ElectricLGV             - Split between 3a and 3b based on existing 3a-3b split.


@author: edward.barratt
"""
import pandas as pd

inputFileName = 'Cracked\EFT2017_v7.4_ExtractedBasicSplit.xlsx'         # Input file
outputFileName = 'Cracked\EFT2017_v7.4_ExtractedBasicSplit_Recount.csv' # Output file
PropBuses = {'Urban': 0.72, 'Rural': 0.72, 'Motorway': 0} # proportion of "8Bus
    # and coach" that will be assigned as buses. the remainder will be coaches.

# Read the file.
inDF = pd.read_excel(inputFileName)
columnNames = list(inDF)

# Get the available Regions
regions = set(list(inDF['Region']))
rTypes = set(list(inDF['Road Type']))
vehicles = set(list(inDF['Vehicle']))

outDF = pd.DataFrame(columns = list(inDF))


for region in regions:
  inDFB = inDF[inDF['Region'] == region]
  for rType in rTypes:
    inDFC_ = inDFB[inDFB['Road Type'] == rType]
    inDFC = inDFC_.drop('Region', 1)
    inDFC = inDFC.drop('Road Type', 1)

    # All Cars
    PetrolCars = inDFC[inDFC['Vehicle'] == '1Petrol car'].drop('Vehicle', 1).values
    PetrolCars = PetrolCars + inDFC[inDFC['Vehicle'] == '10HybridCarPetrol'].drop('Vehicle', 1).values
    PetrolCars = PetrolCars + inDFC[inDFC['Vehicle'] == '11PlugInHybridCarPetrol'].drop('Vehicle', 1).values
    DieselCars = inDFC[inDFC['Vehicle'] == '2Diesel car'].drop('Vehicle', 1).values
    DieselCars = DieselCars + inDFC[inDFC['Vehicle'] == '12HybridCarDiesel'].drop('Vehicle', 1).values
    ElectricCars = inDFC[inDFC['Vehicle'] == '13ElectricCar'].drop('Vehicle', 1).values
    # Assume electric cars are redistributed between petrol and diesel cars.
    PetDieselRatio = PetrolCars/(PetrolCars + DieselCars)
    PetrolCars = PetrolCars + PetDieselRatio*ElectricCars
    DieselCars = DieselCars + (1.0-PetDieselRatio)*ElectricCars

    # And write to the out data frame.
    PetrolRow = [region, rType, '2a. Petrol Cars']
    PetrolRow.extend(PetrolCars.tolist()[0])
    PetrolRow = pd.DataFrame([PetrolRow], columns=columnNames)
    outDF = outDF.append(PetrolRow)
    DieselRow = [region, rType, '2b. Diesel Cars']
    DieselRow.extend(DieselCars.tolist()[0])
    DieselRow = pd.DataFrame([DieselRow], columns=columnNames)
    outDF = outDF.append(DieselRow)

    # All LGVs
    PetrolLGVs = inDFC[inDFC['Vehicle'] == '4Petrol LGV'].drop('Vehicle', 1).values
    DieselLGVs = inDFC[inDFC['Vehicle'] == '5Diesel LGV'].drop('Vehicle', 1).values
    ElectricLGVs = inDFC[inDFC['Vehicle'] == '14ElectricLGV'].drop('Vehicle', 1).values
    # Assume electric LGVs are redistributed between petrol and diesel LGVs.
    PetDieselRatio = PetrolLGVs/(PetrolLGVs + DieselLGVs)
    PetrolLGVs = PetrolLGVs + PetDieselRatio*ElectricLGVs
    DieselLGVs = DieselLGVs + (1.0-PetDieselRatio)*ElectricLGVs

    # And write to the out data frame.
    PetrolRow = [region, rType, '3a. Petrol LGVs']
    PetrolRow.extend(PetrolLGVs.tolist()[0])
    PetrolRow = pd.DataFrame([PetrolRow], columns=columnNames)
    outDF = outDF.append(PetrolRow)
    DieselRow = [region, rType, '3b. Diesel LGVs']
    DieselRow.extend(DieselLGVs.tolist()[0])
    DieselRow = pd.DataFrame([DieselRow], columns=columnNames)
    outDF = outDF.append(DieselRow)

    # Split bus and coach.
    Both = inDFC[inDFC['Vehicle'] == '8Bus and coach'].drop('Vehicle', 1).values
    Buses = PropBuses[rType]*Both
    BusesRow = [region, rType, '4a. Buses']
    BusesRow.extend(Buses.tolist()[0])
    BusesRow = pd.DataFrame([BusesRow], columns=columnNames)
    outDF = outDF.append(BusesRow)
    Coaches = (1-PropBuses[rType])*Both
    CoachesRow = [region, rType, '4b. Coaches']
    CoachesRow.extend(Coaches.tolist()[0])
    CoachesRow = pd.DataFrame([CoachesRow], columns=columnNames)
    outDF = outDF.append(CoachesRow)

    # Others
    for VC in ["3Taxi (black cab)", "6Rigid", "7Artic", "9Motorcycle"]:
      row = inDFC_[inDFC_['Vehicle'] == VC]
      outDF = outDF.append(row)

# Rename vehicle classes.
outDF.loc[outDF.Vehicle == "3Taxi (black cab)", 'Vehicle'] = "2c. Taxis (black cab)"
outDF.loc[outDF.Vehicle == "6Rigid", 'Vehicle'] = "5a. Rigid HGVs"
outDF.loc[outDF.Vehicle == "7Artic", 'Vehicle'] = "5b. Artic HGVs"
outDF.loc[outDF.Vehicle == "9Motorcycle", 'Vehicle'] = "1. Motorcycles"

outDF = outDF.sort_values(['Region', 'Road Type', 'Vehicle'])
outDF.to_csv(outputFileName, index=False)