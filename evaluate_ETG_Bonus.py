"""
Script of ETG Bonus evaluation
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import pandas as pd
import numpy as np

# General constants
result_identifier = '_results'
ff_pool_identifier = 'Formelfrage'
sc_pool_identifier = 'SingleChoice'
export_prefix = '2021w_ETG'
Filename_Export_detailed = export_prefix+'exp_Bonus_det.xlsx'
Filename_Export_public = export_prefix+'exp_Bonus_pub.xlsx'
name_marker = 'Ergebnisse von Testdurchlauf '   # 'Ergebnisse von Testdurchlauf 1 für '
run_marker = 'dummy_text'   # run marker currently not used
tasks = ['Formelfrage', 'Single Choice', 'Lückentextfrage', 'Hotspot/Imagemap', 'Freitext eingeben']
res_marker_ft = "Ergebnis"
var_marker = '$v'
res_marker = '$r'
marker = [run_marker, tasks, var_marker, res_marker, res_marker_ft] 

# Specific constants for members
# read psso member list
psso_members = pd.read_excel('2021w_ETG_Members/psso-2022-01-19/20220121_Kohortenaufteilung_ETG_full_SR.xlsx', 
                             sheet_name='Sheet1')
members = psso_members
members['Matrikelnummer'] = pd.to_numeric(members['Matrikelnummer'])
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

print('PSSO member import OK')

# Specific constants for intermediate tests (Zwischentests=zt)
# What Notes by what total percentage points?
zt_scheme = pd.Series(data= [0,    70], 
                   index=['NB','BE'])
zt_test = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
zt_dir = '2021w_ETG_Zwischentests/'

# Specific constants for Praktikum
# What Notes by what total percentage points?
pra_scheme = pd.Series(data= [0,    50], 
                       index=['NB','BE'])
pra_experiment = [1, 2, 3]
pra_dir = 'ETG_SS21_Praktikum/'

# test data
[zt_ilias_result, zt_pool_ff, zt_pool_sc] = ev.get_excel_files(zt_test, zt_dir)

# read bonus list from Praktika 
# praktika = pd.read_excel('2021w_ETG_Praktika_Bonus.xlsx', 
#                         sheet_name='Tabelle1')
print('Praktika import OK')

i_lev1 = []
i_lev2 = []
subtitles = ['Ges_Pkt', 'Note']
for n in zt_test:
    i_lev1 += ['ZT'+str(n)]*len(subtitles)
    i_lev2 += subtitles
c_tests = pd.MultiIndex.from_arrays([i_lev1, i_lev2], names = ['test', 'parameter'])
course_data = pd.DataFrame(index=members.index, columns=c_tests)

## disable here
######### LOOP of evaluating all considered intermediate tests ############
intermediate_tests = []
for zt in range(len(zt_test)):
    intermediate_tests.append([])
    for sub in range(len(zt_ilias_result[zt])):
        print('started evaluating intermediate test', zt_test[zt])
        intermediate_tests[zt].append(ev.Test(members, marker, zt_test[zt],
                                          zt_ilias_result[zt][sub], ff=zt_pool_ff[zt][sub], sc=zt_pool_sc[zt][sub]))
        print("process ILIAS data...")
        intermediate_tests[zt][sub].process_ilias()
        print("process task pools and evaluate...")
        intermediate_tests[zt][sub].process_pools()
    
print("evaluate zt bonus...")
[members, course_data]= ev.evaluate_intermediate_tests(members, 
                                                       zt_tests=intermediate_tests, 
                                                       d_course=course_data,  
                                                       scheme=zt_scheme,
                                                       tests_p_bonus=2)

###### EVALUATE TOTAL BONUS ###########
members = ev.evaluate_bonus(members)

######## Export for lecturer (detailed, not anonymous) ###########
#TODO: export Bonus with details 

######## Export for participants (short, anonymous) ############
#TODO. export Bonus without details

print("Excel Export OK")
print ('### done! ###')
