
from os import path
import json
import datetime

homeDir = path.expanduser("~")

def getSecret(secretname, filename='~.keys/keys'):
  filename = filename.replace('~', homeDir+'/')
  with open(filename) as json_data:
    d = json.load(json_data)
    try:
      secret = d[secretname]
    except KeyError:
      print('No secret named {}'.format(secretname))
      return 'None'
    return secret

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
