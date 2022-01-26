"""
Script of ETG Praktikum evaluation
Author: Silvan Rummeny
"""
import pandas as pd
import numpy as np
import ilias_evaluator as ev

# General constants
result_identifier = '_results'
ff_pool_identifier = 'Formelfrage'
sc_pool_identifier = 'SingleChoice'
export_prefix = '2021w_ETG_'
name_marker = 'Ergebnisse von Testdurchlauf '   # 'Ergebnisse von Testdurchlauf 1 für '
run_marker = 'dummy_text'   # run marker currently not used
tasks = ['Formelfrage', 'Single Choice', 'Lückentextfrage', 'Hotspot/Imagemap', 'Freitext eingeben']
res_marker_ft = "Ergebnis"
var_marker = '$v'
res_marker = '$r'
marker = [run_marker, tasks, var_marker, res_marker, res_marker_ft] 

# Specific constants for members
mem_dir = '2021w_ETG_Members/'
psso_import = []
ilias_mem = pd.read_excel(mem_dir+'2022_01_20_13-351642682110_member_export_2520764.xlsx', 
                          sheet_name='Mitglieder')
ilias_mem = ilias_mem.loc[ilias_mem['Matrikelnummer'].dropna().index]
# Specific constants for Praktikum
# What Notes by what total percentage points?
pra_scheme = pd.Series(data= [0,    50], 
                       index=['NB','BE'])
pra_experiment = [1, 2, 3]
pra_dir = '2021w_ETG_Praktikum/'

# test data
[pra_ilias_result, pra_pool_ff, pra_pool_sc] = ev.get_excel_files(pra_experiment, pra_dir)

# read bonus list from old Praktika 
old_praktika = pd.read_excel(pra_dir+'2021s_Bonuspunkte_Praktika.xlsx', 
                            sheet_name='Sheet1')
print('Praktika import OK')

# read psso member list
#psso_members_origin = ev.import_psso_members(psso_import)
members = ilias_mem
members['Matrikelnummer'] = pd.to_numeric(members['Matrikelnummer'])
members['Name_'] = np.nan       # members['Name'].str.replace("'","")
members['Benutzername'] = np.nan
members['E-Mail'] = np.nan
members['Bonus_ZT'] = np.nan
members['Bonus_Pra'] = np.nan
members['Bonus_Pkt'] = np.nan
members['Kohorte'] = np.nan
members['Exam_Pkt'] = np.nan
members['Ges_Pkt'] = np.nan
members['ILIAS_Pkt'] = np.nan
members['Note'] = np.nan
for i in range(len(members)): 
# add a space behind the komma of the name
    vorname = members.loc[i, 'Vorname']
    nachname = members.loc[i, 'Nachname']
    members.loc[i, 'Name_'] = nachname + ', ' + vorname
# get Benutzername and Email from ilias_mem
    mtknr_sel = ilias_mem['Matrikelnummer'].astype(int)==members['Matrikelnummer'][i]
    if sum(mtknr_sel)==0: 
        # psso member is not an ilias member
        continue
    members.loc[i,'Benutzername'] = ilias_mem['Benutzername'][mtknr_sel].values.item()
    members.loc[i,'E-Mail'] = ilias_mem['E-Mail'][mtknr_sel].values.item()
## remove ' from Names to get ILIAS equivalent names
members['Name_'] = members['Name_'].str.replace("'", "")
print('PSSO member import OK')

i_lev1 = []
i_lev2 = []
subtitles = ['Ges_Pkt', 'Note']
for n in pra_experiment:
    i_lev1 += ['V'+str(n)]*len(subtitles)
    i_lev2 += subtitles
c_tests = pd.MultiIndex.from_arrays([i_lev1, i_lev2], names = ['test', 'parameter'])
course_data = pd.DataFrame(index=members.index, columns=c_tests)

## disable here
########### LOOP of evaluating all considered Praktikum experiments ############
praktikum = []
for pra in range(len(pra_experiment)):
    praktikum.append([])
    for sub in range(len(pra_ilias_result[pra])):
        print('started evaluating Praktikum test', pra_ilias_result[pra][sub][21:])
        praktikum[pra].append(ev.Test(members, marker, pra_experiment[pra],
                                 pra_ilias_result[pra][sub]))
        print("process ILIAS data...")
        praktikum[pra][sub].process_ilias()
    
print("evaluate pra bonus...")
[members, course_data]= ev.evaluate_praktika(members, pra_prev=old_praktika, 
                                             pra_tests=praktikum, 
                                             d_course=course_data, 
                                             semester_name = '2021w')
print("Done")
