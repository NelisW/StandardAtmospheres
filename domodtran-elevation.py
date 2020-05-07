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
for different tape5 input files and altitudes.
The tape7 file is read and processed to create new data files:
 - transmittance from sun to earth
 - path radiance from sun to earth, no terrain contribution
and in wavelength domain with specified spectral intervals.

All spectral integrals are performed in this code and the 
wideband results are stored in 'atmos-elevation-angles.xlsx'

dir structure:
|dir root
  | file domodtran.py
  | dir scenario 1
    tape5
    | dir 'elev'
      | dir alt x
        | dir elev y
          |file tape5 (template for scenario 1)
          |file tape7

The scenario directories must exist, but the altitude dirs are created.

https://media.readthedocs.org/pdf/pandas_xlsxwriter_charts/latest/pandas_xlsxwriter_charts.pdf
http://xlsxwriter.readthedocs.org/
http://pandas-xlsxwriter-charts.readthedocs.org/en/latest/introduction.html
https://github.com/jmcnamara/XlsxWriter

"""

pathToModtranBin = r'C:\PcModWin5\bin'
wlnum = 1500 #yields about 10 nm wavelength intervals

###############################################################
def createTape5FileElev(tape5base, elev, alt, directory):
  """Reads and existing tape5 and modifies it for new altitude and elevation
  Elevation in degrees
  Altitude in metres
  """

  #read the tape5 base file
  print(tape5base, alt, directory)
  with open(tape5base) as fin:
    lines = fin.readlines()

  #change the altitude to new value
  outlines = []
  for line in lines:
    #the template tape5 has an altitude of '0.305000' to be replaced
    if '0.305000' in line and '135.000000' in line:
      line = line.replace('0.305000','{:08.3f}'.format(alt/1000.))
      line = line.replace('135.000000','{:08.3f}'.format(elev))
    outlines.append(line)

  #create new directory and write file to new dir
  dirname = os.path.join('.',directory,'elev','{:.0f}'.format(alt),'{:.2f}'.format(elev))
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
  p = subprocess.Popen(os.path.join(pathToModtranBin, 'Mod5.2.0.0.exe'), 
    shell=True, stdout=None, stderr=None, cwd=pathToModtranBin)
  while  p.poll() == None:
    time.sleep(0.5)

  #copy the tape6/7 files back to appropriate directory
  shutil.copy2(os.path.join(pathToModtranBin, 'tape6'), datadir)
  shutil.copy2(os.path.join(pathToModtranBin, 'tape7'), datadir)

###########################################################################
def calcEffective(ilines, alt, dir, specranges, dffilename):
 dirname = os.path.join('.',dir,'elev','{:.0f}'.format(alt))
 lfiles = ryfiles.listFiles(dirname, patterns='tape7')
 # print(dirname)
 dfCols = ['Atmo','Altitude','Zenith','SpecBand','ToaWattTot','ToaWatt','BoaWattTot','BoaWatt','LpathWatt',
            'ToaQTot','ToaQ','BoaQTot','BoaQ','LpathQ','effTauSun','effTau300']
 #open the excel file, read in, append to the dataframe and later save to the same file
 if os.path.exists(dffilename): 
  df = pd.read_excel(dffilename)
 else:          
  df = pd.DataFrame(columns=dfCols)

 for filename in lfiles:
  print(filename)
  elev = float(filename.split('\\')[3])
  #read transmittance and path radiance from tape7
  dataset = rymodtran.loadtape7(filename, ['FREQ', 'TOT_TRANS', 'TOA_SUN', 'TOTAL_RAD'] )
  dataset[:,2] *= 1e4 # convert from /cm2 to /m2
  dataset[:,3] *= 1e4 # but only do this once for data set
  for key in specranges:
   tape7 = dataset.copy()
   # select only lines in the spectral ranges 
   select = np.all([ tape7[:,0] >= 1e4/specranges[key][1], tape7[:,0] <= 1e4/specranges[key][0]], axis=0)
   # print(tape7[select])
   tape7s = tape7[select]
   #now integrate the path and TOA radiances. now in W/(m2.sr.cm-1)
   TOAwatttot = np.trapz(tape7[:,2],  tape7[:,0])
   TOAwatt =    np.trapz(tape7s[:,2], tape7s[:,0])
   BOAwatttot = np.trapz(tape7[:,1] *tape7[:,2], tape7[:,0])
   BOAwatt =    np.trapz(tape7s[:,1]*tape7s[:,2], tape7s[:,0])
   Lpathwatt =  np.trapz(tape7s[:,3], tape7s[:,0])
   # convert radiance terms to photon rates, must to this while it is still spectral
   # then integrate after conversion
   conv = tape7[:,0] * const.h * const.c * 1e2
   tape7[:,2] /= conv
   tape7[:,3] /= conv
   tape7s = tape7[select]
   #now integrate the path and TOA radiances. now in (q/s)/(m2.sr.cm-1)
   TOAqtot = np.trapz(tape7[:,2],  tape7[:,0])
   TOAq =    np.trapz(tape7s[:,2], tape7s[:,0])
   BOAqtot = np.trapz(tape7[:,1] *tape7[:,2], tape7[:,0])
   BOAq =    np.trapz(tape7s[:,1]*tape7s[:,2], tape7s[:,0])
   Lpathq =  np.trapz(tape7s[:,3], tape7s[:,0])
   LSun = ryplanck.planck(tape7s[:,0],6000.,'en') 
   L300 = ryplanck.planck(tape7s[:,0],300.,'en')
   effTauSun = np.trapz(LSun * tape7s[:,1], tape7s[:,0]) / np.trapz(LSun, tape7s[:,0])
   effTau300 = np.trapz(L300 * tape7s[:,1], tape7s[:,0]) / np.trapz(L300, tape7s[:,0])

   # print(elev,key, BOAwatttot, BOAwatt, TOAwatttot, TOAwatt, Lpathwatt)
   # print(elev,key, BOAqtot, BOAq, TOAqtot, TOAq, Lpathq)

   df = df.append(pd.DataFrame([[dir,alt,elev,key,
    TOAwatttot,TOAwatt,BOAwatttot,BOAwatt,Lpathwatt,TOAqtot,TOAq,BOAqtot,BOAq,Lpathq,
    effTauSun,effTau300
    ]], columns=dfCols))
   ilines += 1

 writer = pd.ExcelWriter(dffilename)
 df.to_excel(writer,'Sheet1')  
 pd.DataFrame(specranges).to_excel(writer,'SpecRanges')
 writer.save()

 return(ilines)


##########################################################################################################

#each base tape5 file is in its own directory
specranges = {}
with open('StandardSpectralRanges.txt','rt') as fin:
    lines = fin.readlines()
    for line in lines:
        linelst = line.rstrip().split()
        specranges[linelst[0]] = [float(linelst[1]),float(linelst[2])]
print(specranges)

alts = [0, 30000] # altitude in km

dirs = ['ExtremeHotLowHumidity', 'ExtremeHumidity',
        'MidLatMaritimeSummer','MidLatMaritimeWinter',
        'ScandinavianSummer','ScandinavianWinter',
        'TropicalDesert','TropicalRural',
        'TropicalUrban','USStdNavyMarVis23km']   

#do for these elevations (in deg)
elevs = np.hstack((np.linspace(0,70,15),np.linspace(72,80,9),np.linspace(81,90,19),
        np.linspace(90.5,100.,20),np.linspace(102,110,9),np.linspace(110,180,15)))


# dirs = ['MidLatMaritimeSummer']
# elevs = [0, 45]
ilines = 0
for dir in dirs:
 for alt in alts:
  for elev in elevs:
   pass
   # dirname = createTape5FileElev(os.path.join(dir,'tape5'), elev, alt, dir)
   # runModtran(dirname)
  ilines = calcEffective(ilines, alt, dir, specranges,'atmos-elevation-angles.xlsx')

print('Number of lines written to file: {}'.format(ilines))
print('Number of points in data set: {}'.format(len(specranges) * len(elevs) * len(dirs) * len(alts)))

