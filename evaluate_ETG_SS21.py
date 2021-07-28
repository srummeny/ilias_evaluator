"""
Script of ETG Bonus evaluation
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import pandas as pd
import numpy as np
import glob

version = '2.1, 27.07.2021'
print ("Tool zur externen Bewertung von ILIAS Formelfragen-Tests")
print ("Version", version)
print ("(c) by Eberhard Waffenschmidt, TH-Köln")
print ("edited by Silvan Rummeny, TH-Köln")

# General constants
import_members = '2021_06_28_10-011624867289_member_export_2341446.xlsx'
               # '2021_06_14_08-401623652810_member_export_2341446.xlsx' # old file
psso_identifier = 'prf'
result_identifier = '_results'
ff_pool_identifier = 'Formelfrage'
sc_pool_identifier = 'SingleChoice'
export_prefix = 'ETG_SS21_bonus_'
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
mem_dir = 'ETG_SS21_Members/'
psso_import = []
ilias_mem = pd.read_excel(mem_dir+'2021_06_28_10-011624867289_member_export_2341446.xlsx', 
                          sheet_name='Mitglieder')
ilias_mem = ilias_mem.loc[ilias_mem['Matrikelnummer'].dropna().index]

# Specific constants for intermediate tests (Zwischentests=zt)
# What Notes by what total percentage points?
zt_scheme = pd.Series(data= [0,    70], 
                   index=['NB','BE'])
zt_test = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
zt_dir = 'ETG_SS21_ZT/'

# Specific constants for Praktikum
# What Notes by what total percentage points?
pra_scheme = pd.Series(data= [0,    50], 
                       index=['NB','BE'])
pra_experiment = [1, 2, 3]
pra_dir = 'ETG_SS21_Praktikum/'

# Specific constants for Exam
exam_scheme = pd.Series(data= [0,    44,   49,   53,   58,   62,   67,   71,   76,   80,   84],
                        index=["5,0","4,0","3,7","3,3","3,0","2,7","2,3","2,0","1,7","1,3","1,0"]) 
exam_cohort = ['Kohorte A', 'Kohorte F']
exam_dir = 'ETG_SS21_Exam/'

# find all import data and pools in directory
# psso members
for i in range(len(glob.glob(mem_dir+'/*.xls'))):
    if psso_identifier in glob.glob(mem_dir+'/*.xls')[i]:
        psso_import.append(glob.glob(mem_dir+'/*.xls')[i])
    else:
        print('### PSSO import - skipped file:', glob.glob(mem_dir+'/*.xls')[i])

# test data
[zt_ilias_result, zt_pool_ff, zt_pool_sc] = ev.get_excel_files(zt_test, zt_dir)
[pra_ilias_result, pra_pool_ff, pra_pool_sc] = ev.get_excel_files(pra_experiment, pra_dir)
[exam_ilias_result, exam_pool_ff, exam_pool_sc] = ev.get_excel_files(exam_cohort, exam_dir)

# read bonus list from old Praktika 
old_praktika = pd.read_excel('ETG_SS21_ZT/20210614_ETG_SS21_Bonuspunkte.xlsx', 
                            sheet_name='Liste Bonuspunkte - Praktika')
print('Praktika import OK')
# read psso member list
psso_members_origin = ev.import_psso_members(psso_import)
members = psso_members_origin.rename(columns={'mtknr':'Matrikelnummer', 
                                              'sortname':'Name', 
                                              'nachname':'Nachname', 
                                              'vorname':'Vorname'})
members['Matrikelnummer'] = pd.to_numeric(members['Matrikelnummer'])
members['Name_'] = np.nan       # members['Name'].str.replace("'","")
members['Benutzername'] = np.nan
members['E-Mail'] = np.nan
members['Bonus_ZT'] = np.nan
members['Bonus_Pra'] = np.nan
members['Bonus_Pkt'] = np.nan
members['Exam_Pkt'] = np.nan
members['Ges_Pkt'] = np.nan
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
print('PSSO member import OK')

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

### disable here
########## LOOP of evaluating all considered intermediate tests ############
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
                                                       scheme=zt_scheme)
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
                                             d_course=course_data)
# Exception for Zeinab Mohammad: "doch im Testdurchlauf 2 bestanden, erfolgloser Testdurchlauf 3 wurde gewertet
course_data.loc[96, ('V2', 'Ges_Pkt')] = 2
course_data.loc[96, ('V2', 'Note')] = 'BE'

###### EVALUATE TOTAL BONUS ###########
members = ev.evaluate_bonus(members)
#
########### LOOP of evaluating all considered exam cohorts ############
#exam = []
#for c in range(len(exam_cohort)):
#    print('started evaluating exam,', exam_cohort[c])
#    exam.append(ev.Test(members, marker, exam_cohort[c],
#                        exam_ilias_result[c], ff=exam_pool_ff[c], sc=exam_pool_sc[c]))
#    print("process ILIAS data...")
#    exam[c].process_ilias()
#    print("process task pools and evaluate...")
#    exam[c].process_pools()
#    
#print("evaluate exam...")
#[members, all_entries] = ev.evaluate_exam(members, exam, exam_scheme, max_pts=41) 
#print ('export results as excel...')
#drop_columns = ['Name_']
#members.drop(drop_columns, axis=1).to_excel(Filename_Export, index=False, na_rep='N/A')
#members['n_answers'] = np.nan
#for p in range(len(members['Note'].dropna().index)):
#    members.loc[p, 'n_answers'] = sum(exam[0].entries.loc[3, pd.IndexSlice[:,'R']].str.len()>0)
# members['Note'].hist()
# members['n_answers'].dropna().value_counts()
# members['n_answers'].hist()
print ('### done! ###')
                 
#course_export = course_data
#course_export.columns = course_export.columns.map('_'.join)
#members = pd.concat([members[members.columns[:-3]],course_export, 
#                     members[members.columns[-3:]]], axis=1)
#print ('export results as excel...')
#drop_columns = ['Rolle/Status', 'Nutzungsvereinbarung akzeptiert', 'Name_']
#members.drop(drop_columns, axis=1).to_excel(Filename_Export, index=False, na_rep='N/A')
#print ('### done! ###')
