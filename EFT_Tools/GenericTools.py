# -*- coding: utf-8 -*-
"""
euroSearchTerms
euroTechs
numToLetter
romanNumeral
secondsToString

Created on Fri Apr 20 15:48:23 2018

@author: edward.barratt
"""

import datetime

from EFT_Tools import (euroClassNameVariations)


def euroSearchTerms(N, tech='All'):
  ES = euroClassNameVariations[N][tech]
  return ES

def euroTechs(N):
  return euroClassNameVariations[N].keys()

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

def romanNumeral(N):
  # Could write a function that deals with any, but I only need up to 10.
  RNs = [0, 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
  return RNs[N]

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





# =============================================================================
# def randomString(N = 6):
#   return ''.join(random.choice(string.ascii_uppercase + string.ascii_lowercase + string.digits) for x in range(N))
# =============================================================================

# =============================================================================
# def getInputFile(version, directory='input'):
#   """
#   Return the absolute path to the appropriate file for the selected mode and
#   version. Will return an error if no file is available.
#   """
#
#   # First check that the directory exists.
#   if not path.isdir(directory):
#     raise ValueError('Cannot find directory {}.'.format(directory))
#
#   # Now figure out the file name.
#   if version == 6.0:
#     vPart = 'EFT2014_v6.0.2'
#     ext = '.xls'
#   elif version == 7.0:
#     vPart = 'EFT2016_v7.0'
#     ext = '.xlsb'
#   elif version == 7.4:
#     vPart = 'EFT2017_v7.4'
#     ext = '.xlsb'
#   elif version == 8.0:
#     vPart = 'EFT2017_v8.0'
#     ext = '.xlsb'
#   else:
#     raise ValueError('Version {} is not recognised.'.format(version))
#
#   fname = '{}/{}_empty{}'.format(directory, vPart, ext)
#   # return the absolute paths.
#   fname =  path.abspath(fname)
#
#   # Check that file exists.
#   if not path.exists(fname):
#     raise ValueError('Cannot find file {}.'.format(fname))
#
#   return fname
# =============================================================================

# romanNumeral
# secondsToString