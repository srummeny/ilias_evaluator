#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb  8 10:29:55 2022

@author: rummeny
"""

import pandas as pd

bv1 = pd.read_excel('2021w_ETG_Bonuspunkte_pub_v1.xlsx',
                     sheet_name='Sheet1', header=5)

bv2 = pd.read_excel('2021w_ETG_Bonuspunkte_pub.xlsx',
                     sheet_name='Sheet1', header=5)
df = bv2.copy()
for i in range(len(bv2)):
    if any(bv1['Matrikelnummer']==bv2['Matrikelnummer'][i]):
        print('is in!')
    else:
        print(bv2['Matrikelnummer'][i], 'is not in!')
        df = df.drop(i)
sel=bv1['Summe']!=df['Summe']
print(sel[sel])
