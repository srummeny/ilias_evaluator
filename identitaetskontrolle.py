#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jan 31 16:01:35 2022

@author: rummeny
"""

import pandas as pd
import numpy as np

import_dir = '2021w_ETG_Members/Identitaetskontrolle/'
SR = pd.read_excel(import_dir+'20220125_Kohortenaufteilung_ETG_full_Anwesenheitskontrolle_SR.xlsx', header=0, sheet_name='Sheet1')
KM = pd.read_excel(import_dir+'Kohortenaufteilung_ETG_Probeklausur_20220125_KM.xlsx', header=0, sheet_name='Sheet1')
TK = pd.read_excel(import_dir+'Kohortenaufteilung_ETG_Probeklausur_20220125_TK.xlsx', header=0, sheet_name='Sheet1')
RG = pd.read_excel(import_dir+'Kohortenaufteilung_ETG_Probeklausur_20220125_full_RG.xlsx', header=0, sheet_name='Sheet1')
sr_c = 'Kontrolle'
km_c = 'Breakout-Session'
tk_c = 'Identitätskontrolle'
rg_c = 'Geprüft'
df = TK.merge(SR[['Matrikelnummer', sr_c]], how='outer')
df = df.merge(KM[['Matrikelnummer', km_c]], how='outer') 
df = df.merge(RG[['Matrikelnummer', rg_c]], how='outer') 
df = df.drop(columns=['Unnamed: 15'])
print(df[tk_c].dropna()[df[tk_c].dropna()=='Nein'])
df.loc[df[tk_c].dropna()[df[tk_c].dropna()=='Nein'].index,tk_c]=np.nan
print(df[km_c].dropna()[df[km_c].dropna().str.len()>1])
df.loc[df[km_c].dropna()[df[km_c].dropna().str.len()>1].index,km_c]=np.nan
print(df[km_c].dropna()[df[km_c].dropna()=='x'])
df.loc[df[km_c].dropna()[df[km_c].dropna()=='x'].index,km_c]=np.nan
df['Identitaetsnachweis'] = df[[sr_c,km_c, tk_c, rg_c]].any(axis=1)
df = df.loc[df['Matrikelnummer'].dropna().index]
df['Matrikelnummer']=df['Matrikelnummer'].astype(int)
for i in range(len(df)):
    j = members.index[members['Matrikelnummer']==df['Matrikelnummer'][i]]
    members.loc[j, 'Identitaetsnachweis'] = df['Identitaetsnachweis'][i]
print('Is there a Participant without Note? :',members['Identitaetsnachweis'][df1['Note'][df1['Note'].isna()].index].any())

EigErk = pd.read_excel('2021w_ETG_Members/2022125_Eigenstaendigkeitserklaerungen_Probepruefung.xlsx', header=5, sheet_name='Tabelle1')
EigErk['Eigenstaendigkeitserklaerung'] = EigErk['Bewertung']=='bestanden'
for i in range(len(EigErk)):
    j = members.index[members['Benutzername']==EigErk['Benutzername'][i]]
    members.loc[j, 'Eigenstaendigkeitserklaerung'] = EigErk['Eigenstaendigkeitserklaerung'][i]
members['Eigenstaendigkeitserklaerung'] = members['Eigenstaendigkeitserklaerung'].fillna(False)
