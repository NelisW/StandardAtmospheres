"""This file provides Python-style definitions of the standard atmospheres.
This file can be included in other python scripts
"""

#each base tape5 file is in its own directory
dirs = ['ExtremeHotLowHumidity', 'ExtremeHumidity',
        'MidLatMaritimeSummer','MidLatMaritimeWinter',
        'ScandinavianSummer','ScandinavianWinter',
        'TropicalDesert','TropicalRural',
        'TropicalUrban','USStdNavyMarVis23km']
dirs = ['USStdNavyMarVis23km']
        
#altitudes  (in m)
alts = [305, 1524, 3048, 4572, 6096, 7620, 9144, 10668, 12192, 13716, 14326, 15240]
              
atmospheres = {
    u'ExtremeHotLowHumidity':['44 C, 30% RH (18.1), 77 km Vis Desert'],
    u'ExtremeHumidity'      :['35 C, 95% RH (37.9), 23 km Vis Rural'],
    u'MidLatMaritimeSummer' :['21 C, 76% RH (14), 23 km Vis Maritime'],
    u'MidLatMaritimeWinter' :['-1 C, 77% RH (3), 10 km Vis Maritime'],
    u'ScandinavianSummer'   :['14 C, 75% RH (9), 31 km Vis Maritime'],
    u'ScandinavianWinter'   :['-15.9 C, 80% RH (1), 31 km Vis Maritime'],
    u'TropicalDesert'       :['26.6 C, 75% RH (18), 75 km Vis Desert'], 
    u'TropicalRural'        :['26.6 C, 75% RH (18), 23 km Vis Rural'],
    u'TropicalUrban'        :['26.6 C, 75% RH (18), 5 km Vis Urban'],
    u'USStdNavyMarVis23km'  :['15 C, 46% RH (5.9), 23 km Vis Navy Maritime, 7.2 m/s'],
    }
