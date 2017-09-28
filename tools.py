#from os import path
import datetime

#homeDir = path.expanduser("~")

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

def extractVersion(fileName, availableVersions=[6.0, 7.0, 7.4]):
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


if __name__ == '__main__':
  # For testing.

  ss = secondsToString(4321)
  print(ss)
  sl = secondsToString(4321, form='long')
  print(sl)
