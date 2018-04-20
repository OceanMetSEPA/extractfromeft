# -*- coding: utf-8 -*-
"""
splitSourceNameS
splitSourceNameT
splitSourceNameV

Created on Fri Apr 20 14:38:47 2018

@author: edward.barratt
"""

def splitSourceNameS(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  row['vehicle'] = v
  return int(s[1:])

def splitSourceNameT(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  return t

def splitSourceNameV(row, SourceName='Source Name'):
  s = row[SourceName]
  s, v, t = s.split(' - ')
  return v