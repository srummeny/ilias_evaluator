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

# What Notes by what total percentage points (referenced without Bonus)? 
scheme = pd.Series(data= [0,    50,   55,   60,   65,   70,   75,   80,   85,   90,   95],
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

# Specific constants for members
# read psso member list

psso_members = pd.read_excel('2021w_ETG_Members/psso-2022-01-24/20220124_Kohortenaufteilung_ETG_full_SR.xlsx', 
                             sheet_name='Sheet1')
members = psso_members
members['Matrikelnummer'] = pd.to_numeric(members['Matrikelnummer'])
members['Benutzername'] = np.nan
members['bearbeitete Fragen'] = np.nan
members['Bearbeitungsdauer'] = np.nan
members['Startzeit'] = np.nan
members['Bonus_Pkt'] = np.nan
members['Kohorte'] = np.nan
members['Exam_Pkt'] = np.nan
members['Ges_Pkt'] = np.nan
members['ILIAS_Pkt'] = np.nan
members['Note'] = np.nan

print('PSSO member import OK')

import_bonus = pd.read_excel('2021w_ETG_Bonuspunkte_pub.xlsx', header=5, sheet_name='Sheet1')
members['Bonus_Pkt'] = import_bonus['Summe']

print('Bonus import of members OK')

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
######### LOOP of evaluating several cohorts (c) of the exam ############
exam = []
# Log-Dataframe for occurance of Errors
errorlog = pd.DataFrame(columns=['Kohorte', 'Matrikelnummer', 'Task', 'formula', 'var', 'input_res', 'tol', 'Error', 'Points'])
# Log-Dataframe for occurance of Differences between ILIAS result points and ilias_evaluator points
difflog = pd.DataFrame(columns=['Kohorte', 'Matrikelnummer', 'Points_ILIAS', 'Points', 'diff'])      
for c in range(len(considered_tests)):
    print('started evaluating exam,', considered_tests[c])
    exam.append(ev.Test(members, marker, considered_tests[c],
                                      result_import[c], import_pool_FF[c], 
                                      import_pool_SC[c], daIr=True))
    print("process ILIAS data...")
    exam[c].process_ilias()
    print("process task pools and evaluate...")
    exam[c].process_pools()
    errorlog = errorlog.append(exam[c].errorlog)
    
print("evaluate exam...")
[members, all_entries] = ev.evaluate_exam(members, exam, scheme, max_pts=41) 
sel = members['Exam_Pkt'].dropna()!=members['ILIAS_Pkt'].dropna()
df = members.loc[sel[sel].index]
log = pd.DataFrame({'Kohorte':df['Kohorte'],
                    'Matrikelnummer':df['Matrikelnummer'],
                    'Points_ILIAS':df['ILIAS_Pkt'],
                    'Points':df['Exam_Pkt']})
log['diff'] = log['Points'] - log['Points_ILIAS']
difflog = difflog.append(log)
print ('export results as excel...')
members.loc[members['Note'].dropna().index].to_excel(Filename_Export, index=False, na_rep='N/A')
members['n_answers'] = np.nan
for p in members['Note'].dropna().index:
    members.loc[p, 'n_answers'] = sum(exam[0].ent.loc[3, pd.IndexSlice[:,'R']].str.len()>0)
# members['Note'].hist()
# members['n_answers'].dropna().value_counts()
# members['n_answers'].hist()
print ('### done! ###')
