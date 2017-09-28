# ExtractFromEFT #

A suite of tools for interacting with and extracting data from the [Emission Factor Toolkit](https://laqm.defra.gov.uk/review-and-assessment/tools/emissions-factors-toolkit.html).

See https://oceanmet.atlassian.net/wiki/display/airmod/Updates+to+EFT+-+v7.4# for details.


## processEFT.py ##
This group of functions is used to extract single vehicle emission factors from the EFT.

### Usage: processEFT.py ###
```python
usage: processEFT.py [-h] [--version [version number]]
                     [--area [areas [areas ...]]] [--years [year [year ...]]]
                     [--euros [euro classes [euro classes ...]]]
                     [--mode [mode]] [--keeptemp [keeptemp]]
                     [--inputfile [input file]]

Extract emission values from the EFT

optional arguments:
  -h, --help            show this help message and exit
  --version [version number], -v [version number]
                        The EFT version number. One of 6.0, 7.4, 7.0. Default
                        7.0.
  --area [areas [areas ...]], -a [areas [areas ...]]
                        The areas to be processed. One or more of 'England
                        (not London)', 'Northern Ireland', 'Scotland',
                        'Wales'. Default 'all'.
  --years [year [year ...]], -y [year [year ...]]
                        The year or years to be processed. Default 'all'
  --euros [euro classes [euro classes ...]], -e [euro classes [euro classes ...]]
                        The euro class or classes to be processed. One of more
                        number between 0 and 6. Default 0 1 2 3 4 5 6.
  --mode [mode], -m [mode]
                        The mode. One of 'ExtractAll', 'ExtractCarRatio',
                        'ExtractBus'. Default 'ExtractAll'.
  --keeptemp [keeptemp]
                        Whether to keep or delete temporary files. Boolean.
                        Default False (delete).
  --inputfile [input file], -i [input file]
                        The file to process. If set then version will be
                        ignored.
```

For the InputFile the prefilled EFT .xlsb files in the Input directory are designed to assist this process. For 