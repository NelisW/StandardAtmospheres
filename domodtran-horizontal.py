import os
import shutil
import subprocess
import time
import numpy as np
from scipy.interpolate import  interp1d
import pandas as pd
import xlsxwriter
# from openpyxl import load_workbook
import scipy.constants as const

import pyradi.ryplot as ryplot
import pyradi.rymodtran as rymodtran
import pyradi.ryplanck as ryplanck
import pyradi.ryutils as ryutils
import pyradi.ryfiles as ryfiles


"""This script creates and run modtran on multiple tape5 files,
for different tape5 input files and horizontal distances.

dir structure:
|dir root
  | file domodtran.py
  | dir scenario 1
    tape5
    | dir 'horizontal'
      | dir alt x
        | dir distance y
          |file tape5 (template for scenario 1)
          |file tape7

The scenario directories must exist, but the other dirs are created.

https://media.readthedocs.org/pdf/pandas_xlsxwriter_charts/latest/pandas_xlsxwriter_charts.pdf
http://xlsxwriter.readthedocs.org/
http://pandas-xlsxwriter-charts.readthedocs.org/en/latest/introduction.html
https://github.com/jmcnamara/XlsxWriter

"""

pathToModtranBin = r'C:\PcModWin5\bin'
wlnum = 1500 #yields about 10 nm wavelength intervals

###############################################################
def createTape5FileDist(tape5base, dist, alt, directory):
  """Reads and existing tape5 and modifies it for new altitude and distance
  Distance in metres
  Altitude in metres
  """

  #read the tape5 base file
  print(tape5base, dist, alt, directory)
  with open(tape5base) as fin:
    lines = fin.readlines()

  #change the altitude to new value
  outlines = []
  for line in lines:
    #the template tape5 has an altitude of '0.305000' to be replaced
    if 'F 7    3    2' in line:
      line = line.replace('F 7    3    2','F 7    1    2')
    if 'F 2    3    2' in line:
      line = line.replace('F 2    3    2','F 2    1    2')
    if 'F 6    3    2' in line:
      line = line.replace('F 6    3    2','F 6    1    2')
    if 'F 1    3    2' in line:
      line = line.replace('F 1    3    2','F 1    1    2')
    if 'F 5    3    2' in line:
      line = line.replace('F 5    3    2','F 5    1    2')
    if 'F 4    3    2' in line:
      line = line.replace('F 4    3    2','F 4    1    2')
    if 'F 3    3    2' in line:
      line = line.replace('F 3    3    2','F 3    1    2')

    if '0.305000' in line and '135.000000' in line:
      line = line.replace('0.305000','{:08.3f}'.format(alt))
      line = line.replace('135.000000   0.00000','135.000000  {:08.3f}'.format(dist))
      line = line.replace('135.000000','{:08.3f}'.format(90))
    outlines.append(line)

  #create new directory and write file to new dir
  dirname = os.path.join('.',directory,'horizontal','{:.0f}'.format(alt),'{:.2f}'.format(dist))
  if not os.path.exists(dirname):
      os.makedirs(dirname)

  filename = os.path.join(dirname, 'tape5')
  with open(filename,'w') as fout:
    fout.writelines(outlines)

  return dirname

#################################################################
def runModtran(datadir):
  #copy the tape5 file to C:\PcModWin5\bin
  filename = os.path.join(datadir, 'tape5')
  shutil.copy2(filename, pathToModtranBin)

  #run modtran on the tape5 file in its bin directory
  p = subprocess.Popen(os.path.join(pathToModtranBin, 'OntarMod5_3_2.exe'),
    shell=True, stdout=None, stderr=None, cwd=pathToModtranBin)
  while  p.poll() == None:
    time.sleep(0.5)

  #copy the tape6/7 files back to appropriate directory
  shutil.copy2(os.path.join(pathToModtranBin, 'tape6'), datadir)
  shutil.copy2(os.path.join(pathToModtranBin, 'tape7'), datadir)



##########################################################################################################

#each base tape5 file is in its own directory
alts = [0., 1., 2., 5., 10., 20.] # altitude in km

dirs = ['ExtremeHotLowHumidity', 'ExtremeHumidity',
        'MidLatMaritimeSummer','MidLatMaritimeWinter',
        'ScandinavianSummer','ScandinavianWinter',
        'TropicalDesert','TropicalRural',
        'TropicalUrban','USStdNavyMarVis23km']

#do for these distances (in km)
distances = [1.0]

# dirs = ['MidLatMaritimeSummer']

for dir in dirs:
 for alt in alts:
  for distance in distances:
   pass
   dirname = createTape5FileDist(os.path.join(dir,'tape5'), distance, alt, dir)
   runModtran(dirname)

