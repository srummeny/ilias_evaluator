"""
Script of ETG Bonus evaluation
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
scheme = pd.Series(data= [0,    70], 
                   index=['NB','BE'])
considered_tests = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
import_dir = 'ETG_SS21_ZT/'
result_id = '_results'
ff_pool_id = 'Formelfrage'
sc_pool_id = 'SingleChoice'
result_import = []
import_pool_FF = []
import_pool_SC = []
import_members = import_dir+'2021_06_28_10-011624867289_member_export_2341446.xlsx'
                          # '2021_06_14_08-401623652810_member_export_2341446.xlsx' # old file
export_prefix = 'ETG_SS21_ZT_'
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
    for i in range(len(glob.glob(import_dir+str(j)+'/*.xlsx'))):
        dir_i = glob.glob(import_dir+str(j)+'/*.xlsx')[i]
        if result_id in dir_i or ff_pool_id in dir_i or sc_pool_id in dir_i:
            if result_id in dir_i:
                result_import.append(glob.glob(import_dir+str(j)+'/*.xlsx')[i])
            elif ff_pool_id in dir_i:
                import_pool_FF.append(glob.glob(import_dir+str(j)+'/*.xlsx')[i])
            elif sc_pool_id in dir_i:
                import_pool_SC.append(glob.glob(import_dir+str(j)+'/*.xlsx')[i])
        else:
            result_import.append(None)
### disable here

# read ILIAS-members as DataFrame members
# containing: Name, Matrikelnummer, Gesamtpunkte, Note, Bonuspunkte, etc.
members = pd.read_excel (import_members, sheet_name='Mitglieder')
print ("ILIAS members OK")
# Create new columns with full name
# Full name as used in ILIAS_results
members['Name'] = members['Nachname']+', '+members['Vorname']
# Full name as used in ILIAS_data (usually without "'" in Name)
members['Name_'] = members['Name'].str.replace("'","")
# Create DataFrame course_data for Test result aggregation
i_lev1 = []
i_lev2 = []
subtitles = ['Ges_Pkt', 'Note']
for n in considered_tests:
    i_lev1 += ['ZT'+str(n)]*len(subtitles)
    i_lev2 += subtitles
c_tests = pd.MultiIndex.from_arrays([i_lev1, i_lev2], names = ['test', 'parameter'])
course_data = pd.DataFrame(index=members.index, columns=c_tests)
members['Bonus_ZT'] = 0.0
members['Bonus_Pra'] = 0.0
members['Bonus_Pkt'] = 0.0
# read bonus list from Praktika 
praktika = pd.read_excel('ETG_SS21_ZT/20210614_ETG_SS21_Bonuspunkte.xlsx', 
                            sheet_name='Liste Bonuspunkte - Praktika')
print('Praktika import OK')
########## LOOP of evaluating several intermediate tests ############
intermediate_tests = []
for zt in range(len(considered_tests)):
    print('started evaluating intermediate test', considered_tests[zt])
    intermediate_tests.append(ev.Test(members, marker, considered_tests[zt],
                                      result_import[zt], import_pool_FF[zt], 
                                      import_pool_SC[zt]))
    print("process ILIAS data...")
    intermediate_tests[zt].process_d_ilias()
    print("process task pools and evaluate...")
    intermediate_tests[zt].process_pools()
    
print("evaluate bonus...")
[members, course_data]= ev.evaluate_bonus(members, praktika, 
                                          tests=intermediate_tests, 
                                          course_data=course_data,  
                                          scheme=scheme)
course_export = course_data
course_export.columns = course_export.columns.map('_'.join)
members = pd.concat([members[members.columns[:-3]],course_export, 
                     members[members.columns[-3:]]], axis=1)
print ('export results as excel...')
drop_columns = ['Rolle/Status', 'Nutzungsvereinbarung akzeptiert', 'Name_']
members.drop(drop_columns, axis=1).to_excel(Filename_Export, index=False, na_rep='N/A')
print ('### done! ###')
