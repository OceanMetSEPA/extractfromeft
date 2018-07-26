# ExtractFromEFT #

A suite of tools for interacting with and extracting data from the
[Emission Factor Toolkit](https://laqm.defra.gov.uk/review-and-assessment/tools/emissions-factors-toolkit.html)
(EFT).

See https://oceanmet.atlassian.net/wiki/display/airmod/Updates+to+EFT+-+v7.4# for details.

The EFT files themselves are macro-infested binary excel files, python uses the win32com client
to open and manipulate the files, so unfortunately this code will only work on a machine
running Windows.

## extractEFT.py ##
This group of functions is used to extract single vehicle emission factors from
the EFT. Note that the input EFT file requires some preparation (described in
the documentation below).

#### Usage: extractEFT.py ####
```text
usage: extractEFT.py [-h] [-a [areas [areas ...]]] [-y [year [year ...]]]
                     [-e [euro classes [euro classes ...]]] [-w weight mode]
                     [-t tech mode] [--keeptemp [KEEPTEMP]]
                     [--loggingmode [{INFO,DEBUG}]]
                     input file output directory

Extract emission values from a designated EFT file broken down by location,
year, euro class, technology and weight class. Output will be saved to
individual csv files for each location, year, euro class and technology. A log
file will be created that tracks progress of the processing. All files will be
saved to a directory that should be specified by the user. Ideally an empty
otherwise unused directory should be specified. Processing is very slow,
expect at least 5 minutes for every iteration of location, year, euro class
and technology. If the processing is cancelled by the user (ctrl+c), or
otherwise fails then it can be restarted, so long as the same directory is
specified then the programme will first read through the log file and will not
re-create files that have already been processed.

positional arguments:
  input file            The file to process. This should be a copy of EFT
                        version 7.4 or greater, and it needs a small amount of
                        initial setup. Under Select Pollutants select NOx,
                        PM10 and PM2.5. Under Traffic Format select
                        'Alternative Technologies'. Select 'Emission Rates
                        (g/km)' under 'Select Outputs', and 'Euro
                        Compositions' under 'Advanced Options'. All other
                        fields should be either empty or should take their
                        default values.
  output directory      The directory in which to save output files. If the
                        directory does not exist then it will be created,
                        assuming required permissions, etc.

optional arguments:
  -h, --help            show this help message and exit
  -a [areas [areas ...]]
                        The areas to be processed. One or more of 'England
                        (not London)', 'Northern Ireland', 'Scotland',
                        'Wales', 'all'. Default 'Scotland'.
  -y [year [year ...]]  The year or years to be processed. Default current
                        year.
  -e [euro classes [euro classes ...]]
                        The euro class or classes to be processed. One of more
                        number between 0 and 6, or 99 which will instruct the
                        code to use the default euro breakdown for the year
                        specified. Default 99.
  -w weight mode        The weight mode, either 'all' or 'mix'. 'mix' mode
                        does not split emission factors by weight class,
                        instead using the default mix from the EFT, while
                        'all' mode splits by all possible weight classes.
                        Default 'all'.
  -t tech mode          The technology mode, either 'all' or 'mix'. 'mix' mode
                        does not split emission factors by technology class,
                        instead using the default mix from the EFT , while
                        'all' mode splits by all possible technology classes.
                        By technology we mean the varied additional
                        technologies applied to vehicles of different euro
                        classes to reduce emission factors, e.g. DPF for
                        diesel vehicles, and technology 'c' and 'd' for euro
                        class 6. Default 'all'.
  --keeptemp [KEEPTEMP]
                        Whether to keep or delete temporary files. Boolean.
                        Default False (delete).
  --loggingmode [{INFO,DEBUG}]
                        The logging mode. Either INFO or DEBUG, default INFO.
```

## extractEFT.py helper routines ##
The following programmes are used to make running extractEFT.py easier.

### combineExtracted.py ###
This will merge all of the individual output files for each iteration of year,
technology, etc. into a single .csv file.

#### Usage combineExtracted.py ####
```text
usage: combineExtracted.py [-h] directory

Combine the csv files produced by extractEFT.py in to one large csv file.
Files marked 'COMPLETED' within the log file (which must be present in the
designated directory) will be combined. Other files within the directory, csv
or otherwise, will be ignored.

positional arguments:
  directory   The directory containing the log file for an extractEFT.py
              processing job.

optional arguments:
  -h, --help  show this help message and exit
```

### compressLog.py ###
The log file for extractEFT.py can get very long, and since it is used to determine
which iteration of year, technology, etc, still needs to be processed, this can
slow down the start of processing. compressLog.py strips the log file of all lines
that do not include the key words required by extractEFT.py, while saving the original
log file with a sensible new name for posterity.

#### Usage compressLog.py ####
``` text
usage: compressLog.py [-h] directory

Removes all unneccesary lines from the extractEFT log file within the selected
directory, leaving only lines about 'COMPLETED' or 'SKIPPED' files, since
these lines are key lines used by other processes. Renames the original log
file, so other information is not deiscarded completely.

positional arguments:
  directory   The directory containing the log file for an extractEFT.py
              processing job.

optional arguments:
  -h, --help  show this help message and exit
```

## extractVehProportions.py ##
This function extracts the vehicle fleet split, i.e. the proportion of different
euro classes, weigh classes, etc. from the EFT. Collecting this data is simply
a matter of reading it from the spreadsheet, no macros need to be run, so it is
much quicker than extractEFT.py

### Usage extractVehProportions.py ###
```
usage: extractVehProportions.py [-h] [-a [areas [areas ...]]]
                                [-y [year [year ...]]] [--keeptemp [KEEPTEMP]]
                                [--loggingmode [{INFO,DEBUG}]]
                                input file output directory

Extract the default vehicle euroclass proportion, and the default vehicle
weight class proportion, from the EFT file broken down by location, and year.

positional arguments:
  input file            The file to process. This should be a copy of EFT
                        version 7.4 or greater, and it needs a small amount of
                        initial setup. Under Select Pollutants select NOx,
                        PM10 and PM2.5. Under Traffic Format select
                        'Alternative Technologies'. Select 'Emission Rates
                        (g/km)' under 'Select Outputs', and 'Euro
                        Compositions' under 'Advanced Options'. All other
                        fields should be either empty or should take their
                        default values.
  output directory      The directory in which to save output files. If the
                        directory does not exist then it will be created,
                        assuming required permissions, etc.

optional arguments:
  -h, --help            show this help message and exit
  -a [areas [areas ...]]
                        The areas to be processed. One or more of 'England
                        (not London)', 'Northern Ireland', 'Scotland',
                        'Wales', 'all'. Default 'Scotland'.
  -y [year [year ...]]  The year or years to be processed. Default current
                        year.
  --keeptemp [KEEPTEMP]
                        Whether to keep or delete temporary files. Boolean.
                        Default False (delete).
  --loggingmode [{INFO,DEBUG}]
                        The logging mode. Either INFO or DEBUG, default INFO.
```

## extractVehProportions.py helper routines ##
The following programmes are used to make running extractVehProportions.py easier.

### combineProportions.py ###
This will merge all of the individual output files for each iteration of year,
and location into a single .csv file.

#### Usage combineProportions.py ####
```text
usage: combineProportions.py [-h] directory

Combine the csv files produced by extractVehProportions.py in to one large csv
file.

positional arguments:
  directory   The directory containing the log file for an extractEFT.py
              processing job.

optional arguments:
  -h, --help  show this help message and exit
```

##  shp2EFT.py ##
This programme will take a shape file representing roads with associated traffic counts
and run it through the EFT. It will save a new shapefile with the emission rates for
NOx, NO2, PM10 and PM2.5, for each road, added as new attributes.

### Usage shp2EFT.py ###
```text
usage: shp2EFT.py [-h]
                  [--vehFleetSplit [Vehicle euro class and weight split file.]]
                  [--vehCountNames [Vehicle count field names]]
                  [--trafficFormat [Traffic Format]] [-a [area]] [-y [year]]
                  [--saveloc [output shape file location]]
                  [--speedFieldName [speed field name]]
                  [--classFieldName [class field name]]
                  [--no2file [no2 factor file]] [--keeptemp [keeptemp]]
                  shapefile eftfile

Processes the contents of a shape file through the Emission Factor Toolkit
(EFT).

positional arguments:
  shapefile             The shapefile to be processed. The file must have
                        attributes corresponding to vehicle counts, which can
                        be specified using --vehCountNames.
  eftfile               The EFT file to use. This should be a copy of EFT
                        version 7.4 or greater, and it needs a small amount of
                        initial setup. Under Select Pollutants select NOx,
                        PM10 and PM2.5. Under Traffic Format select 'Detailed
                        Option 2'. Select 'Emission Rates (g/km)' under
                        'Select Outputs', and 'Euro Compositions' and 'Primary
                        NO2 Fraction' under 'Advanced Options'. All other
                        fields should be either empty or should take their
                        default values.

optional arguments:
  -h, --help            show this help message and exit
  --vehFleetSplit [Vehicle euro class and weight split file.]
                        A euro split and weight split proportions file. A
                        template is available in the 'input' directory of the
                        repository as 'ProportionsTemplate.xlsx'. A tool for
                        creating a file from ANPR data is available in the
                        OceanMetSEPA.fleetSplit repository.
  --vehCountNames [Vehicle count field names]
                        The shapefile field names for the vehicles counts.
                        Default "MCYCLE CAR TAXI LGV RHGV_2X RHGV_3X RHGV_4X
                        AHGV_34X AHGV_5X AHGV_6X BUS".
  --trafficFormat [Traffic Format]
                        The traffic format to be used by the EFT. Default
                        'Detailed Option 2'.
  -a [area]             The areas to be processed. One of 'England (not
                        London)', 'Northern Ireland', 'Scotland', 'Wales'.
                        Default 'Scotland'.
  -y [year]             The year to be processed. Default present year.
  --saveloc [output shape file location]
                        Location to save the output shape file. If not
                        assigned one will be created based on the input
                        shapefile.
  --speedFieldName [speed field name]
                        The shapefile field name for the road speed, which
                        itself should be in kmh.
  --classFieldName [class field name]
                        The shapefile field name for the road class. Roads
                        will be processed as Urban roads, unless they are
                        marked, in this field, as 'Motorway' (or any Scottish
                        motorway name, e.g. 'M8') or 'Rural'. Default None,
                        which will set all road to Urban.
  --no2file [no2 factor file]
                        The NOx to NO2 conversion factor file to use. Has no
                        effect for EFT v8.0. Default
                        input/NAEI_NO2Extracted.xlsx.
  --keeptemp [keeptemp]
                        Whether to keep or delete temporary files. Boolean.
                        Default False (delete).
```

## ExtractHiddenSheets.py ##
Most of the internal data within the EFT is saved on hidden sheets and protected
with a password. Well it turns out that passwords in excel documents are only
really useful for blocking access to users using excel, python can read the
sheet's contents no problem.

This script extracts all sheets from the EFT and saves them to a new un-password-protected document. It's
not perfect, you don't get to see the original formatting, and you don't get to see any functions or the
source for the macros, but it's better than nothing.

As a script rather than a function this programme will need to be edited by users when it is required.

##  EFT_Tools.py ##
A range of functions used by other main programmes.