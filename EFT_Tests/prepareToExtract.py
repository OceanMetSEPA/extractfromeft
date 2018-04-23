# -*- coding: utf-8 -*-
"""
Tests for the extractEFT programmes.

Created on Mon Apr 23 10:55:37 2018

@author: edward.barratt
"""

import os
import sys
sys.path.insert(0, os.path.abspath('../'))
import unittest

from extractfromeft import EFT_Tools as tools

inputDir = os.path.normpath('C:/Users/edward.barratt/Documents/Development/Python/extractfromeft/input')
inputFiles = {6: 'EFT2014_v6.0.2_empty.xls', 7: 'EFT2016_v7.0_empty.xlsb', 8: 'EFT2017_v8.0_emptyAlternativeTech.xlsb'}


class extractVersion_TestCases(unittest.TestCase):

  def test_finds_correct_versions(self):
    for v, fi in inputFiles.items():
      fpath = os.path.join(inputDir, fi)
      vr, vro = tools.extractVersion(fpath, verbose=False)
      # Version numbers should be identical and correct.
      self.assertEqual(v, vr)
      self.assertEqual(v, vro)

  def test_fails_with_nonsense_in(self):
    self.assertRaises(ValueError, tools.extractVersion, 'ThisIsNotAFileThisIsNotAFileThisIsNot', verbose=False)

  def test_assumes_most_recent(self):
    vr, vro = tools.extractVersion('ThisIsNotAFileThisIsNotAFileThisIsNot', checkExist=False, verbose=False, availableVersions=[6.0, 7.0, 7.4, 8.0])
    self.assertEqual(8.0, vr)
    self.assertEqual('Unknown Version as 8.0', vro)



class prepareToExtract_TestCases(unittest.TestCase):

  def test_returns_correct(self):
    for v, fi in inputFiles.items():
      fpath = os.path.join(inputDir, fi)
      ahk_ahkpathGot, fileNames, versionNos, versionForOutputs = tools.prepareToExtract(fpath, verbose=False)

      # ahk path should exist.
      self.assertTrue(os.path.exists(ahk_ahkpathGot))

      # should be only one filename, and it should exist.
      self.assertEqual(len(fileNames), 1)
      self.assertTrue(os.path.exists(fileNames[0]))



#class splitSourceNameTestCases(unittest.TestCase):
#  testname = '54 - Car -
#
#  def test_return
if __name__ == '__main__':
  unittest.main()