# ExtractFromEFT #

A suite of tools for interacting with and extracting data from the [Emission Factor Toolkit](https://laqm.defra.gov.uk/review-and-assessment/tools/emissions-factors-toolkit.html).

See https://oceanmet.atlassian.net/wiki/display/airmod/Updates+to+EFT+-+v7.4# for details.


## extractEFT.py ##
This group of functions is used to extract single vehicle emission factors from the EFT.

### Usage: extractEFT.py ###
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

##  shp2EFT.py ##

### Usage shp2EFT.py ###
'''python
usage: shp2EFT.py [-h] [--vehCountNames [Vehicle count field names]]
                  [--trafficFormat [Traffic Format]] [-a [area]] [-y [year]]
                  [--saveloc [output shape file location]]
                  [--speedFieldName [speed field name]]
                  [--classFieldName [class field name]]
                  [--no2file [no2 factor file]] [--keeptemp [keeptemp]]
                  shapefile eftfile

Processes the contents of a shape file through the Emission Factor Toolkit
(EFT).

positional arguments:
  shapefile             The shapefile to be processed. This programme is
                        designed to work with shape files produced for the
                        traffic noise modelling project. See details below.
  eftfile               The EFT file to use.

optional arguments:
  -h, --help            show this help message and exit
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
'''