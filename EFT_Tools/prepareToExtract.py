# -*- coding: utf-8 -*-
"""
extractVersion
prepareToExtract

Created on Fri Apr 20 15:48:23 2018

@author: edward.barratt
"""
import os
import numpy as np

from EFT_Tools import (ahk_ahkpath,
                       ahk_exepath)

workingDir = os.getcwd()

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
  if not os.path.isfile(ahk_exepath):
    print('The Autohotkey executable file {} could not be found.'.format(ahk_exepath))
    ahk_ahkpathGot = None
  if not os.path.isfile(ahk_ahkpath):
    ahk_ahkpath_ = workingDir + '\\' + ahk_ahkpath
    if not os.path.isfile(ahk_ahkpath_):
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
    fN_ = os.path.abspath(fN)
    if not os.path.isfile(fN):
      raise ValueError('Could not find {}.'.format(fN))
    fileNames[fNi] = fN_

  return ahk_ahkpathGot, fileNames, versionNos, versionForOutputs