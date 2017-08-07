
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


if __name__ == '__main__':
  # For testing.
  ss = secondsToString(4321)
  print(ss)
  sl = secondsToString(4321, form='long')
  print(sl)
