"""
Script of ETG Bonus evaluation
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import pandas as pd
import numpy as np

### Important entries

export_prefix = '2022s_ETG'
# read psso member list
psso_members = pd.read_excel('2022s_ETG_Members/psso-2022-06-21/20220712_Kohortenaufteilung_ETG_full_SR.xlsx', 
                             sheet_name='Sheet1')
print('PSSO member import OK')
zt_dir = '2022s_ETG_Zwischentests/'
# read bonus list from Praktika 
pra = pd.read_excel('2022s_ETG_Pra_Bonus.xlsx',
                    sheet_name='Sheet1')
print('Praktika import OK')
first_print_line = 'Bonuspunkte Elektrische Energietechnik (ETG), SoSe 22'      # WiSe 21/22
# Specific constants for intermediate tests (Zwischentests=zt)
# What Notes by what total percentage points?
zt_scheme = pd.Series(data= [0,    70], 
                   index=['NB','BE'])
zt_test = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
# Specific constants for Praktikum
# What Notes by what total percentage points?
pra_experiment = [1, 2, 3]

###

# General constants
result_identifier = '_results'
ff_pool_identifier = 'Formelfrage'
sc_pool_identifier = 'SingleChoice'
Filename_Export_public = export_prefix+'_Bonuspunkte_pub.xlsx'
name_marker = 'Ergebnisse von Testdurchlauf '   # 'Ergebnisse von Testdurchlauf 1 für '
run_marker = 'dummy_text'   # run marker currently not used
tasks = ['Formelfrage', 'Single Choice', 'Lückentextfrage', 'Hotspot/Imagemap', 'Freitext eingeben']
res_marker_ft = "Ergebnis"
var_marker = '$v'
res_marker = '$r'
marker = [run_marker, tasks, var_marker, res_marker, res_marker_ft] 

# Specific constants for members
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

# test data
[zt_ilias_result, zt_pool_ff, zt_pool_sc] = ev.get_excel_files(zt_test, zt_dir)

i_lev1 = []
i_lev2 = []
subtitles = ['Ges_Pkt', 'Note']
for n in zt_test:
    i_lev1 += ['ZT'+str(n)]*len(subtitles)
    i_lev2 += subtitles
for n in pra_experiment:
    i_lev1 += ['V'+str(n)]*len(subtitles)
    i_lev2 += subtitles
c_tests = pd.MultiIndex.from_arrays([i_lev1, i_lev2], names = ['test', 'parameter'])
course_data = pd.DataFrame(index=members.index, columns=c_tests)

## disable here
######### LOOP of evaluating all considered intermediate tests (zt) ############
zt = []
# Log-Dataframe for occurance of Errors
errorlog = pd.DataFrame(columns=['Test', 'Matrikelnummer', 'Task', 'formula', 'var', 'input_res', 'tol', 'Error', 'Points'])
# Log-Dataframe for occurance of Differences between ILIAS result points and ilias_evaluator points
difflog = pd.DataFrame(columns=['Test', 'Matrikelnummer', 'Task', 'formula', 'var', 'input_res', 'tol', 'Points_ILIAS', 'Points', 'diff'])      
for t in range(len(zt_test)):
    zt.append([])
    for sub in range(len(zt_ilias_result[t])):
        print('started evaluating intermediate test', zt_test[t])
        zt[t].append(ev.Test(members, marker, zt_test[t],
                                          zt_ilias_result[t][sub], ff=zt_pool_ff[t][sub], sc=zt_pool_sc[t][sub]))
        print("process ILIAS data...")
        zt[t][sub].process_ilias()
        print("process task pools and evaluate...")
        zt[t][sub].process_pools()
        errorlog = errorlog.append(zt[t][sub].errorlog)
        difflog = errorlog.append(zt[t][sub].difflog)
    
print("evaluate zt bonus...")
[members, course_data]= ev.evaluate_intermediate_tests(members, 
                                                       zt_tests=zt, 
                                                       d_course=course_data,  
                                                       scheme=zt_scheme,
                                                       tests_p_bonus=3)
[members, course_data]= ev.evaluate_praktika(members, 
                                             pra_prev = pra,
                                             d_course=course_data,
                                             tests_p_bonus=1, 
                                             semester_name=export_prefix[0:5])

###### EVALUATE TOTAL BONUS ###########
members = ev.evaluate_bonus(members)

#%%
######## Export for lecturer (detailed, not anonymous) ###########
# TODO: Export hier noch nötig?

######## Export for participants (short, anonymous) ############
Bonus_exp_pub = members[['Matrikelnummer','Bonus_ZT', 'Bonus_Pra', 'Bonus_Pkt']]
Bonus_exp_pub = Bonus_exp_pub.rename(columns={'Bonus_ZT':'Boni durch Zwischentests',
                                      'Bonus_Pra':'Boni durch Praktika',
                                      'Bonus_Pkt':'Summe'})
Bonus_exp_pub = Bonus_exp_pub.sort_values(by=['Matrikelnummer'])
writer = pd.ExcelWriter(Filename_Export_public)
Bonus_exp_pub.to_excel(writer ,index=False, na_rep='N/A', startrow=5)
workbook = writer.book
worksheet = writer.sheets['Sheet1']
title_format = workbook.add_format({'bold': True, 'font_size':16})
align = workbook.add_format({'align':'left'})
worksheet.write_string(0,0, first_print_line, title_format)
worksheet.write_string(1,0,'1 bestandenes Praktikum = 1 Bonuspunkt')
worksheet.write_string(2,0,'3 bestandene Zwischentests = 1 Bonuspunkt')
worksheet.write_string(3,0,'Die Summe der Bonuspunkte kann nur max. 5 Punkte betragen')
worksheet.write_string(4,0,'N/A = nicht teilgenommen')
worksheet.set_column(0,0,15, align)
worksheet.set_column(1,1,25, align)
worksheet.set_column(2,2,20, align)
worksheet.set_column(3,3,7, align)
writer.save()
print("Excel Export OK")
print ('### done! ###')
