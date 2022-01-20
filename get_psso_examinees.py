"""
Script of getting all signed examinees of PSSO-System
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import datetime as dt
import pandas as pd
import glob

cohorts = ['A', 'B', 'C']
cohorts_normal = cohorts[0]
now = dt.datetime.now()
member_dir = '2021w_Members/psso-2022-01-19/'
psso_identifier = 'prf'
NTA_identifier = 'NTA'
psso_member_export = now.strftime('%d%m%Y')+'_Kohortenaufteilung_ETG_full_SR.xlsx'
psso_import = []
nta = []

# find all psso member lists in directory
for i in range(len(glob.glob(member_dir+'/*.xls'))):
    if psso_identifier in glob.glob(member_dir+'/*.xls')[i]:
        psso_import.append(glob.glob(member_dir+'/*.xls')[i])
    else:
        print('### Skipped file:', glob.glob(member_dir+'/*.xls')[i])
# find NTA list
for i in range(len(glob.glob(member_dir+'/*.ods'))):
    if NTA_identifier in glob.glob(member_dir+'/*.ods')[i]:
        nta.append(glob.glob(member_dir+'/*.ods')[i])
    else:
        print('### Skipped file:', glob.glob(member_dir+'/*.ods')[i])

# read psso member list and rename columns
psso_members_origin = ev.import_psso_members(psso_import)
nta = pd.read_excel(nta[0], header=4)
psso_members = psso_members_origin.rename(columns={'mtknr':'Matrikelnummer', 
                                                   'sortname':'Name', 
                                                   'nachname':'Nachname', 
                                                   'vorname':'Vorname'})
psso_members['Matrikelnummer'] = psso_members['Matrikelnummer'].astype(int)   

# check if any Matrikelnummer occurs more than once in member list
if len(psso_members['Matrikelnummer'].value_counts()[psso_members['Matrikelnummer'].value_counts()>1]) > 0:
    print('double entry!')
    print(psso_members['Matrikelnummer'].value_counts()[psso_members['Matrikelnummer'].value_counts()>1])
    print('please check and remove double entry!')

psso_members['Kohorte'] = 'A'
for i in nta['Matrikel'].dropna().index:
    sel = psso_members['Matrikelnummer']==nta['Matrikel'].dropna().astype(int)[i]
    psso_members.loc[sel,'Kohorte'] = nta['Kohorte'][i]
    print('added Kohorte', psso_members.loc[sel,'Kohorte'].values[0], 
          'for member', psso_members.loc[sel,'Matrikelnummer'].values[0])
psso_members = psso_members.sort_values(['Kohorte', 'Matrikelnummer'], ignore_index=True)
print('Kohorten:')
print(psso_members['Kohorte'].value_counts())
psso_members.to_excel(member_dir+psso_member_export, index=False, na_rep='N/A')
