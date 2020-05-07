import os
import shutil
import subprocess
import time
import numpy as np
from scipy.interpolate import  interp1d
import pandas as pd
import xlsxwriter
# from openpyxl import load_workbook

import pyradi.ryplot as ryplot
import pyradi.rymodtran as rymodtran
import pyradi.ryutils as ryutils


"""This script creates and run modtran on multiple tape5 files, 
for different tape5 input files and altitudes.
The tape7 file is read and processed to create a new data file
containing the transmittance normalised to 1 km path length
and in wavelength domain with specified spectral intervals.
The resulting output filenames have the form 'dirname-altm.1km'
and the file contains two columns: 
wavelength in um and transmittance (over 1 km path length)
dir structure:
|dir root
  | file domodtran.py
    | dir scenario 1
         |dir alt 1
         |dir alt 2
         |dir alt 3
         |file tape5 template for scenario 1
    | dir scenario 2
         |dir alt 1
         |dir alt 2
         |dir alt 3
         |file tape5 template for scenario 2
The scenario directories must exist, but the altitude dirs are created.

https://media.readthedocs.org/pdf/pandas_xlsxwriter_charts/latest/pandas_xlsxwriter_charts.pdf
http://xlsxwriter.readthedocs.org/
http://pandas-xlsxwriter-charts.readthedocs.org/en/latest/introduction.html
https://github.com/jmcnamara/XlsxWriter

"""

pathToModtranBin = r'C:\PcModWin5\bin'
slantAngle = 45.0 * np.pi / 180.
wlnum = 1500 #yields about 10 nm wavelength intervals

def createTape5File(tape5base, alt, directory):

  #read the tape5 base file
  print(tape5base, alt, directory)
  with open(tape5base) as fin:
    lines = fin.readlines()

  #change the altitude to new value
  outlines = []
  for line in lines:
    #the template tape5 has an altitude of '0.305000' to be replaced
    outlines.append(line.replace('0.305000','{:08.3f}'.format(alt/1000.)))

  #create new directory and write file to new dir
  dirname = os.path.join('.',directory,'{}'.format(alt))
  if not os.path.exists(dirname):
      os.makedirs(dirname)

  filename = os.path.join(dirname, 'tape5')
  with open(filename,'w') as fout:
    fout.writelines(outlines)

def runModtran(alt, directory):
  datadir = os.path.join('.',directory,'{}'.format(alt))
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

  #read transmittance from tape7
  tape7 = rymodtran.loadtape7(os.path.join(datadir,'tape7'), 
    ['FREQ', 'DEPTH'] )

  #rescale transmittance to 1 km path length
  tape7[:,1] = np.exp(- tape7[:,1] * 1.0e3 / (alt / (np.cos(slantAngle)) ))

  #convolve transmittance to get to lower wavenumber resolution
  tape7[:,1],  windowfn = ryutils.convolve(tape7[:,1], 1, 1, 8)

  #convert to wavelength scale by interpolation
  wl = np.linspace(1.0e4/tape7[-1,0], 1.0e4/tape7[0,0], wlnum)
  interpfunT = interp1d(1.0e4/tape7[:,0], tape7[:,1], bounds_error=False, fill_value=0.0)
  twl = interpfunT(wl)

  #save to new file
  filename = os.path.join(datadir,'{}-{}m.1km'.format(directory, alt))
  with open(filename, 'wt') as fout:
    fout.write('scenario {}, altitude {} m\n'.format(directory, alt))
    np.savetxt(fout,np.hstack((wl.reshape(-1,1),twl.reshape(-1,1))))


def plotTau(alts, dirs):
  #first plot the different altitudes for each atmosphere
  for i,directory in enumerate(dirs):
    p = ryplot.Plotter(i, 1, 1, figsize=(12,6))
    for alt in alts:
      datadir = os.path.join('.',directory,'{}'.format(alt))
      filename = '{}-{}m.1km'.format(directory, alt)
      data = np.loadtxt(os.path.join(datadir,filename), skiprows=1)
      p.plot(1, data[:,0], data[:,1],
        '{} {} 1 km transmittance, 135 deg zenith'.format(directory, directory),
        'Wavelength $\mu$m','Transmittance', label=['{} m'.format(alt)],legendAlpha=0.5)
    p.saveFig(os.path.join('.',directory,'{}.png'.format(directory)))
    
  #now plot the different altitudes for each altitude
  for alt in alts:
    p = ryplot.Plotter(i, 1, 1, figsize=(12,6))
    for i,directory in enumerate(dirs):
      datadir = os.path.join('.',directory,'{}'.format(alt))
      filename = '{}-{}m.1km'.format(directory, alt)
      data = np.loadtxt(os.path.join(datadir,filename), skiprows=1)
      p.plot(1, data[:,0], data[:,1],
        '{} m altitude, 135 deg zenith 1 km transmittance'.format(alt),
        'Wavelength $\mu$m','Transmittance', label=['{} {}'.format(directory, directory)],
        legendAlpha=0.5)
    p.saveFig(os.path.join('.','AllScen-{}m.png'.format(alt)))

def writeXLS(alts, dirs):
  #first write the different altitudes for each atmosphere
  sheetname = 'tau'
  for i,directory in enumerate(dirs):
      filename = os.path.join(directory,'{}.xlsx'.format(directory))

      # Create an new Excel file and add a worksheet.
      writer = pd.ExcelWriter(filename, engine='xlsxwriter')

      #create and write the data
      for i,alt in enumerate(alts):
        datadir = os.path.join('.',directory,'{}'.format(alt))
        filename = '{}-{}m.1km'.format(directory, alt)
        data = np.loadtxt(os.path.join(datadir,filename), skiprows=1)
        if i==0:
          outdata = data
        else:
          outdata = np.hstack((outdata, data[:,1].reshape(-1,1)))
      data = pd.DataFrame(outdata)
      data.to_excel(writer, sheet_name=sheetname, startrow=4, startcol=1, header=False, index=False)

      # Write the headers, get the workbook handle from writer
      # worksheet = workbook.add_worksheet()
      workbook = writer.book
      worksheet = workbook.worksheets()[0]
      worksheet.write('C2', 'Altitude')
      worksheet.write('B3', 'Wavelength')
      worksheet.write('B4', u'\u00B5m')
      for i,alt in enumerate(alts):
        worksheet.write(3, i+2, alt)
        worksheet.write(2, i+2, 10* int(round(alt * 0.328083)))
      worksheet.write(3, i+2+1, 'm')
      worksheet.write(2, i+2+1, 'ft')

##########################################################################################################

#each base tape5 file is in its own directory

dirs = ['ExtremeHotLowHumidity', 'ExtremeHumidity',
        'MidLatMaritimeSummer','MidLatMaritimeWinter',
        'ScandinavianSummer','ScandinavianWinter',
        'TropicalDesert','TropicalRural',
        'TropicalUrban','USStdNavyMarVis23km']       
#do for these altitudes  (in m)
alts = [305, 1524, 3048, 4572, 6096, 7620, 9144, 10668, 12192, 13716, 14326, 15240]

#dirs = ['USStdNavyMarVis23km']
# alts = [305]

for dir in dirs:
 for alt in alts:
   createTape5File(os.path.join(dir,'tape5'), alt, dir)
   runModtran(alt, dir)

plotTau(alts, dirs)

writeXLS(alts, dirs)
