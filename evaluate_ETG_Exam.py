"""
Script of ETG Exam evaluation
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import pandas as pd
import numpy as np
import glob

version = '2.0, 29.06.2021'
print ("Tool zur externen Bewertung von ILIAS Formelfragen-Tests")
print ("Version", version)
print ("(c) by Eberhard Waffenschmidt, TH-Köln")
print ("edited by Silvan Rummeny, TH-Köln")

# What Notes by what total percentage points? 
scheme = pd.Series(data= [0,    44,   49,   53,   58,   62,   67,   71,   76,   80,   84],
                   index=["5,0","4,0","3,7","3,3","3,0","2,7","2,3","2,0","1,7","1,3","1,0"]) 
considered_tests = ['Kohorte A', 'Kohorte B', 'Kohorte C']
import_dir = '2021w_ETG_Probeklausur/'
result_file_identifier = '_results'
pool_FF_file_identifier = 'Formelfrage'
pool_SC_file_identifier = 'SingleChoice'
result_import = []
import_pool_FF = []
import_pool_SC = []
import_members = import_dir+'2021_06_28_10-011624867289_member_export_2341446.xlsx'
                          # '2021_06_14_08-401623652810_member_export_2341446.xlsx' # old file
export_prefix = '2021w_ETG_Probeklausur_'
Filename_Export = export_prefix+'exp_Ergebnisse.xlsx'
Filename_Export_public = export_prefix+'exp_Ergebnisse_pub.xlsx'
Filename_Export_PSSO = export_prefix+'exp_Ergebnissse_psso.xlsx'
IRT_Frame_Name = export_prefix+'irt_frame.xlsx'

name_marker = 'Ergebnisse von Testdurchlauf '   # 'Ergebnisse von Testdurchlauf 1 für '
run_marker = 'dummy_text'   # run marker currently not used
tasks = ['Formelfrage', 'Single Choice', 'Lückentextfrage', 'Hotspot/Imagemap', 'Freitext eingeben']
res_marker_ft = "Ergebnis"
var_marker = '$v'
res_marker = '$r'
marker = [run_marker, tasks, var_marker, res_marker, res_marker_ft] 
# find all import data and pools in directory
for j in considered_tests:
    if len(glob.glob(import_dir+str(j)+'/*.xlsx')) > 3:
        print('### WARNING: There may be to much import data files in', 
              import_dir+str(j), '###')
    for i in range(len(glob.glob(import_dir+str(j)+'/*.xlsx'))):
        if result_file_identifier in glob.glob(import_dir+str(j)+'/*.xlsx')[i]:
            result_import.append(glob.glob(import_dir+str(j)+'/*.xlsx')[i])
        if pool_FF_file_identifier in glob.glob(import_dir+str(j)+'/*.xlsx')[i]:
            import_pool_FF.append(glob.glob(import_dir+str(j)+'/*.xlsx')[i])
        if pool_SC_file_identifier in glob.glob(import_dir+str(j)+'/*.xlsx')[i]:
            import_pool_SC.append(glob.glob(import_dir+str(j)+'/*.xlsx')[i])
#### disable here
# add a space behind the komma of the name
print('PSSO member import OK')
# read bonus list from Praktika 
praktika = pd.read_excel('ETG_SS21_ZT/20210614_ETG_SS21_Bonuspunkte.xlsx', 
                            sheet_name='Liste Bonuspunkte - Praktika')
print('Praktika import OK')
# find all psso member lists in directory
# TODO: read all members from .xlsx
# TODO: import_bonus = '20210526_Übersicht Bonuspunkte.xlsx'
members = pd.read_excel('ETG_SS21_ZT/ETG_SS21_ZT_exp_Ergebnisse.xlsx', 
                      sheet_name='Sheet1')
members = members.rename(columns={'Bonuspunkte':'Bonus_Pkt'})
members['Name_'] = members['Name'].str.replace("'","")
members['Exam_Pkt'] = np.nan
members['Ges_Pkt'] = np.nan
members['Note'] = np.nan
# compare ILIAS members and psso_members
# ilias course members which are missing in psso --> ignore in evaluation
members = members.loc[members['Matrikelnummer'].dropna().index]
sel_ilias = [members['Matrikelnummer'].astype(int).values[i] not in psso_members['Matrikelnummer'].astype(int).values for i in range(len(members))]
# psso members which are missing in ilias course --> add to evaluation
sel_psso = [psso_members['Matrikelnummer'].astype(int).values[i] not in members['Matrikelnummer'].astype(int).values for i in range(len(psso_members))]
# get praktikum bonus of members missing in ilias course
missing_members = psso_members.loc[sel_psso].copy()
missing_members.loc[:,'Bonus_ZT'] = 0.0
missing_members.loc[:,'Bonus_Pra'] = 0.0
missing_members.loc[:,'Bonus_Pkt'] = 0.0
[missing_members, course_data]= ev.evaluate_bonus(missing_members, praktika)
members = pd.concat([members, missing_members], ignore_index=True)
print('Bonus import of members OK')
######### LOOP of evaluating several cohorts (c) of the exam ############
exam = []
for c in range(len(considered_tests)):
    print('started evaluating exam,', considered_tests[c])
    exam.append(ev.Test(members, marker, considered_tests[c],
                                      result_import[c], import_pool_FF[c], 
                                      import_pool_SC[c]))
    print("process ILIAS data...")
    exam[c].process_d_ilias()
    print("process task pools and evaluate...")
    exam[c].process_pools()
    
print("evaluate exam...")
[members, all_entries] = ev.evaluate_exam(members, exam, scheme, max_pts=41) 
print ('export results as excel...')
drop_columns = ['Name_']
members.drop(drop_columns, axis=1).to_excel(Filename_Export, index=False, na_rep='N/A')
members['n_answers'] = np.nan
for p in range(len(members['Note'].dropna().index)):
    members.loc[p, 'n_answers'] = sum(exam[0].entries.loc[3, pd.IndexSlice[:,'R']].str.len()>0)
# members['Note'].hist()
# members['n_answers'].dropna().value_counts()
# members['n_answers'].hist()
print ('### done! ###')
