"""
Script of getting all signed examinees of PSSO-System
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import datetime as dt
import pandas as pd
import glob

### Important entries

cohorts = ['A'] #, 'B', 'C']
member_dir = '2022s_ETG_Members/psso-2022-06-21/'

###

cohorts_normal = cohorts[0]
now = dt.datetime.now()
psso_identifier = 'prf'
NTA_identifier = 'NTA'
psso_member_export = now.strftime('%Y%m%d')+'_Kohortenaufteilung_ETG_full_SR.xlsx'
psso_import = []
nta = []
kohorten = ['A', 'B', 'C', 'D', 'E', 'F']

# find all psso member lists in directory
for i in range(len(glob.glob(member_dir+'/*.xls'))):
    if psso_identifier in glob.glob(member_dir+'/*.xls')[i]:
        psso_import.append(glob.glob(member_dir+'/*.xls')[i])
    else:
        print('### Skipped file:', glob.glob(member_dir+'/*.xls')[i])
# find NTA list
#for i in range(len(glob.glob(member_dir+'/*.ods'))):
#    if NTA_identifier in glob.glob(member_dir+'/*.ods')[i]:
#        nta.append(glob.glob(member_dir+'/*.ods')[i])
#    else:
#        print('### Skipped file:', glob.glob(member_dir+'/*.ods')[i])

# read psso member list and rename columns
psso_members_origin = ev.import_psso_members(psso_import)
#nta = pd.read_excel(nta[0], header=4)
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

psso_members['Kohorte'] = 'A'           #'-normal'
#for i in nta['Matrikel'].dropna().index:
#    sel = psso_members['Matrikelnummer']==nta['Matrikel'].dropna().astype(int)[i]
#    psso_members.loc[sel,'Kohorte'] = nta['Kohorte'][i]
#    print('added Kohorte', psso_members.loc[sel,'Kohorte'].values[0], 
#          'for member', psso_members.loc[sel,'Matrikelnummer'].values[0])
#psso_members = psso_members.sort_values(['Kohorte', 'Matrikelnummer'], ignore_index=True)
#k = sum(psso_members['Kohorte'] == '-normal')
#j = 1
#while k//j >= 150:
#    j += 1
#rest = k%i
#i_k = []
#i_prev = 0
#for i in range(j):
#    if rest >=1:
#        i_k.append(range(i_prev, (i+1)*k//j+1))
#        i_prev = (i+1)*k//j+1
#        rest -=1
#    else:
#        i_k.append(range(i_prev, (i+1)*k//j))
#        i_prev = (i+1)*k//j
#    psso_members['Kohorte'].loc[i_k[-1]] = kohorten[i]
#    if i_k[-1][-1]+1 == k:
#        print('----------------------------')
#        print('Kohorte', kohorten[i],':')
#        print('DONE: All',k,'normal members are assigned to a cohort')
#
#    else:
#        num = sum(psso_members['Kohorte'] == '-normal')
#        print('----------------------------')
#        print('Kohorte', kohorten[i],':')
#        print('NOT DONE: still',num, 'normal members are not assigned to a cohort')
print('Kohorten:')
print(psso_members['Kohorte'].value_counts())
psso_members.to_excel(member_dir+psso_member_export, index=False, na_rep='N/A')
# psso_members_old = pd.read_excel('2021w_ETG_Members/psso-2022-01-24/20220124_Probeprüfung_Kohortenaufteilung_ETG_full_SR.xlsx')
# df = psso_members.merge(psso_members_old, on=['Matrikelnummer'],how='outer', indicator=False)
