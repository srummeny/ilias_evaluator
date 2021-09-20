"""
Script of Exam evaluation
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import pandas as pd
import numpy as np
import glob

print ("Tool zur externen Bewertung von ILIAS Formelfragen-Tests")
print ("Version", ev.__version__)
print ("(c) by Eberhard Waffenschmidt, TH-Köln")
print ("edited by Silvan Rummeny, TH-Köln")

# General constants
import_members = '2021_06_28_10-011624867289_member_export_2341446.xlsx'
               # '2021_06_14_08-401623652810_member_export_2341446.xlsx' # old file
               
# File identifiers
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

# Specific constants for Exam
exam_scheme = pd.Series(data= [0,    44,   49,   53,   58,   62,   67,   71,   76,   80,   84],
                        index=["5,0","4,0","3,7","3,3","3,0","2,7","2,3","2,0","1,7","1,3","1,0"]) 
exam_cohort = ['Kohorte A', 'Kohorte B', 'Kohorte C', 'Kohorte D', 'Kohorte E', 'Kohorte F']
exam_dir = 'ETG_SS21_Exam/'

# find all import data and pools in directory
# psso members
print ('# Start File Import #')
for i in range(len(glob.glob(mem_dir+'/*.xls'))):
    if psso_identifier in glob.glob(mem_dir+'/*.xls')[i]:
        psso_import.append(glob.glob(mem_dir+'/*.xls')[i])
    else:
        print('### PSSO import - skipped file:', glob.glob(mem_dir+'/*.xls')[i])

# test data
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
# remove ' from Names to get ILIAS equivalent names
members['Name_'] = members['Name_'].str.replace("'", "")
print('PSSO member import OK')

########## LOOP of evaluating all considered exam cohorts ############
exam = []
for c in range(len(exam_cohort)):
    print('started evaluating exam,', exam_cohort[c])
    exam.append(ev.Test(members, marker, exam_cohort[c],
                        exam_ilias_result[c][0], ff=exam_pool_ff[c][0], sc=exam_pool_sc[c][0]))
    print("process ILIAS data...")
    exam[c].process_ilias()
    print("process task pools and evaluate...")
    exam[c].process_pools()
    
print("evaluate exam...")
[members, all_entries] = ev.evaluate_exam(members, exam, exam_scheme, max_pts=41) 

a = members['Ges_Pkt'].value_counts().sort_index()
b = members['ILIAS_Pkt'].value_counts().sort_index()
ab = pd.merge(a, b, how='outer', on=a.index)
ab[['ILIAS_Pkt','Ges_Pkt']].plot.bar()

print ('### done! ###')
