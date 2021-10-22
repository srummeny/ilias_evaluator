#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 15 14:37:36 2021

@author: rummeny
"""
import pandas as pd

#####
old_pra = pd.read_excel('20210330_Liste Bonuspunkte Praktika WS17-WS20.xlsx', sheet_name='aktuelle Tabelle')
old_pra = old_pra.rename(columns={'Matrikelnr.':'Matrikelnummer'})
old_pra = old_pra.drop(columns=['Unnamed: 6', 'Unnamed: 7'])
pra_2021s = course_data[course_data[['V1','V2','V3']].columns[[1,3,5]]]
pra_2021s = pra_2021s.rename(columns={('V1','Note'):'V1', ('V2', 'Note'):'V2', ('V3','Note'):'V3'})
# filter rows with all nan
pra_2021s = pra_2021s[~pra_2021s.isnull().all(axis=1)]
pra_2021s = pra_2021s.replace(['BE', 'NB'], [1, 0]).fillna(0)
pra_2021s = pra_2021s.merge(members.loc[pra_2021s.index, 'Matrikelnummer'], left_index=True, right_index=True)
pra_2021s['Semester']='2021s'
pra_2021s['Summe'] = pra_2021s[['V1', 'V2', 'V3']].sum(axis=1)
pra_2021s = pra_2021s[['Semester', 'Matrikelnummer', 'V1', 'V2', 'V3', 'Summe']]
pra = old_pra.append(pra_2021s, ignore_index=True)
pra = pra.sort_values(by = ['Semester', 'Matrikelnummer'],ascending=[False, True])
pra.to_excel('2021s_Bonuspunkte_Praktika.xlsx',index=False)
