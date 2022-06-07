"""
Script of ETG Exam evaluation
Author: Silvan Rummeny
"""
import ilias_evaluator as ev
import pandas as pd
import numpy as np
import glob

version = '2.1, 23.05.2022'
print ("Tool zur externen Bewertung von ILIAS Formelfragen-Tests")
print ("Version", version)
print ("(c) by Silvan Rummeny, TH-Köln")
print ("based on related work of Eberhard Waffenschmidt, TH-Köln")

# What Notes by what total percentage points (referenced without Bonus)? 
scheme = pd.Series(data= [0,    50,   55,   60,   65,   70,   75,   80,   85,   90,   95],
                   index=["5,0","4,0","3,7","3,3","3,0","2,7","2,3","2,0","1,7","1,3","1,0"]) 
considered_tests = ['Kohorte A'] #, 'Kohorte B', 'Kohorte C', 'Kohorte D', 'Kohorte E', 'Kohorte F']
import_dir = '2022s_Demo_Klausur/'
result_file_identifier = '_results'
pool_FF_file_identifier = 'Formelfrage'
pool_SC_file_identifier = 'SingleChoice'
result_import = []
import_pool_FF = []
import_pool_SC = []

export_prefix = '2022s_Demo_Klausur_'
Filename_Export = export_prefix+'exp_Ergebnisse.xlsx'
Filename_Export_public = export_prefix+'exp_Ergebnisse_pub.xlsx'
Filename_Export_PSSO = export_prefix+'exp_Ergebnissse_psso.xlsx'
Filename_Export_review_detailed = export_prefix+'exp_results_review_det.xlsx'
Filename_Export_review_public = export_prefix+'exp_results_review.xlsx'

name_marker = 'Ergebnisse von Testdurchlauf '   # 'Ergebnisse von Testdurchlauf 1 für '
run_marker = 'dummy_text'   # run marker currently not used
tasks = ['Formelfrage', 'Single Choice', 'Lückentextfrage', 'Hotspot/Imagemap', 'Freitext eingeben']
res_marker_ft = "Ergebnis"
var_marker = '$v'
res_marker = '$r'
marker = [run_marker, tasks, var_marker, res_marker, res_marker_ft] 

# Specific constants for members
# read psso member list

psso_members = pd.read_excel('2022s_Demo_Members/20220523_Kohortenaufteilung_Demo_full_SR.xlsx', 
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
members['% über Bestehensgrenze'] = None
members['Identitaetsnachweis'] = np.nan
members['Eigenstaendigkeitserklaerung'] = np.nan
members['Note'] = np.nan

print('PSSO member import OK')

import_bonus = pd.read_excel('2022s_Demo_Bonuspunkte_pub.xlsx', header=5, sheet_name='Sheet1')
bonus_mrg = pd.merge(members['Matrikelnummer'], import_bonus, how='left', on='Matrikelnummer') 
members['Bonus_Pkt'] = bonus_mrg['Summe']

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
# members = ev.get_originality_proof(members)
[members, all_entries] = ev.evaluate_exam(members, exam, scheme, max_pts=10) 
sel = members['Exam_Pkt'].dropna()!=members['ILIAS_Pkt'].dropna()
df = members.loc[sel[sel].index]
log = pd.DataFrame({'Kohorte':df['Kohorte'],
                    'Matrikelnummer':df['Matrikelnummer'],
                    'Points_ILIAS':df['ILIAS_Pkt'],
                    'Points':df['Exam_Pkt']})
log['diff'] = log['Points'] - log['Points_ILIAS']
difflog = difflog.append(log)
#%%
print ('export results as excel...')
members[['Identitaetsnachweis','Eigenstaendigkeitserklaerung']] = members[['Identitaetsnachweis','Eigenstaendigkeitserklaerung']].replace([True, False],['Ja','Nein'])
# sort members by Matrikelnummer
members = members.sort_values(by=['Matrikelnummer'])
# consider only members with Note
members = members.loc[members['Note'].dropna().index]

######### Export for PSSO ############
exp_psso = members[members.columns[0:13]].rename(columns={'Matrikelnummer':'mtknr',
                                                          'Name':'sortname',
                                                          'Nachname':'nachname',
                                                          'Vorname':'vorname'})   
exp_psso.to_excel(Filename_Export_PSSO, index=False)

######## Export for lecturer (detailed, not anonymous) ###########
cols_detailed = ['Matrikelnummer', 'Name', 'Benutzername', 'stg', 
                 'pversuch', 'Kohorte', 'ILIAS_Pkt', 'Exam_Pkt', 'Bonus_Pkt', 'Ges_Pkt', 'bewertung']
exp_detailed = members[cols_detailed].rename(columns={'stg':'Studiengang',
                                                      'pversuch':'Prüfungsversuch',
                                                      'Kohorte':'Exam_Kohorte',
                                                      'ILIAS_Pkt':'Exam_ILIAS_Pkt',
                                                      'bewertung':'Bewertung'})  
exp_detailed.to_excel(Filename_Export, index=False)

######## Export for participants (short, anonymous) ############
cols_public = ['Matrikelnummer', 'Note']
exp_public = members[cols_public]
exp_public.to_excel(Filename_Export_public, index=False)

######## Export for Exam review (detailed and short) ############
exam_export = all_entries.copy()
exam_export.columns = exam_export.columns.map('_'.join)
exam_export = exam_export[~exam_export['A1_Title'].isnull()]
idx = exam_export.index
review = pd.concat([members.loc[idx,['Matrikelnummer','Name']], exam_export, members.loc[idx,['Bonus_Pkt','Ges_Pkt','% über Bestehensgrenze','Identitaetsnachweis','Eigenstaendigkeitserklaerung','Note']]],axis=1)
for i in range(11):
    review = review.drop(columns=['A'+str(i+1)+'_ID',
                                  'A'+str(i+1)+'_Type', 
                                  'A'+str(i+1)+'_Pkt_ILIAS'])
    review = review.rename(columns={'A'+str(i+1)+'_R':'A'+str(i+1)+'_Student',
                                    'A'+str(i+1)+'_R_ref':'A'+str(i+1)+'_Musterloesung',
                                    'A'+str(i+1)+'_Pkt':'A'+str(i+1)+'_Punkte'})
review = review.rename(columns={'Bonus_Pkt':'Bonuspunkte', 'Ges_Pkt':'Gesamtpunkte'})
exp_review_detailed = review.copy()
# exp_review_detailed = exp_review_detailed.sort_values(by=['Matrikelnummer'])
exp_review_detailed = exp_review_detailed.transpose()
exp_review_detailed.to_excel(Filename_Export_review_detailed, header=False, index=True, na_rep='N/A')

exp_review_public = review.copy()
for i in range(11):
    exp_review_public = exp_review_public.drop(columns=['A'+str(i+1)+'_Title',
                                                        'A'+str(i+1)+'_Formula',
                                                        'A'+str(i+1)+'_Var',
                                                        'A'+str(i+1)+'_Tol'])
exp_review_public = exp_review_public.drop(columns=['Name'])
exp_review_public = exp_review_public.sort_values(by=['Matrikelnummer'])
exp_review_public = exp_review_public.transpose()
exp_review_public = exp_review_public.reset_index()
writer = pd.ExcelWriter(Filename_Export_review_public)
exp_review_public.to_excel(writer , index=False, header=False, na_rep='N/A', startrow=5)
# add all different formats
title = writer.book.add_format({'bold': True, 'font_size':16, 'font_color':'#ff0000', 'fg_color':'#ffff00'})
subtitle = writer.book.add_format({'italic': True, 'font_color': '#00b050', 'fg_color':'#ffff00'})
remark = writer.book.add_format({'bold': True, 'fg_color':'#ffff00'})
matnr = writer.book.add_format({'bold': True, 'fg_color':'#b2b2b2', 'border': 1, 'align':'left'})
idx = writer.book.add_format({'bold':True})
ax_stud = writer.book.add_format({'fg_color':'#b7b3ca', 'border': 1, 'align':'left'})
ax_muster = writer.book.add_format({'fg_color':'#b2b2b2', 'border': 1, 'align':'left'})
ax_pkt = writer.book.add_format({'fg_color':'#ffdbb6', 'border': 1, 'align':'left'})
ax_stud_i = writer.book.add_format({'bold': True, 'fg_color':'#b7b3ca', 'border': 1, 'align':'left'})
ax_muster_i = writer.book.add_format({'bold': True, 'fg_color':'#b2b2b2', 'border': 1, 'align':'left'})
ax_pkt_i = writer.book.add_format({'bold': True, 'fg_color':'#ffdbb6', 'border': 1, 'align':'left'})                                    
footer = writer.book.add_format({'fg_color':'#ffff00', 'border': 1, 'align':'left'})
footer_i = writer.book.add_format({'bold': True, 'fg_color':'#ffff00', 'border': 1, 'align':'left'})
note = writer.book.add_format({'bold': True, 'fg_color':'#81d41a', 'border': 1, 'align':'left'})                         
writer.sheets['Sheet1'].write_string(0,0,'Ergebnisse der Demo-Klausur vom 23.05.2022, Demo-Modul, SoSe 22')
writer.sheets['Sheet1'].set_row(0,cell_format=title)
writer.sheets['Sheet1'].write_string(1,0,'A#_Student ist ihre getaetigte Antwort', subtitle)
writer.sheets['Sheet1'].set_row(1,cell_format=subtitle)
writer.sheets['Sheet1'].write_string(2,0,'A#_Musterloesung ist die Richtige Antwort', subtitle)
writer.sheets['Sheet1'].set_row(2,cell_format=subtitle)
writer.sheets['Sheet1'].write_string(3,0,'A#_Punkte sind die resultierenden Punkte aus der jeweilgen Aufgabe', subtitle)
writer.sheets['Sheet1'].set_row(3,cell_format=subtitle)
writer.sheets['Sheet1'].write_string(4,0,'Die Bestehensgrenze betraegt für alle Teilnehmer 21 Punkte', remark)
writer.sheets['Sheet1'].set_row(4,cell_format=remark)

# set columns width in pxl

writer.sheets['Sheet1'].set_column(1, len(exp_review_public.columns),12)
writer.sheets['Sheet1'].set_row(5,cell_format=matnr)
for i in range(11):
#    writer.sheets['Sheet1'].write_blank(0, 6+i*3, cell_format=ax_stud.set_bold())
    writer.sheets['Sheet1'].write_string(6+i*3,0, exp_review_public['index'][1+i*3], cell_format=ax_stud_i)
    writer.sheets['Sheet1'].write_string(7+i*3,0, exp_review_public['index'][2+i*3], cell_format=ax_muster_i)
    writer.sheets['Sheet1'].write_string(8+i*3,0, exp_review_public['index'][3+i*3], cell_format=ax_pkt_i)
    writer.sheets['Sheet1'].set_row(6+i*3,cell_format=ax_stud)
    writer.sheets['Sheet1'].set_row(7+i*3,cell_format=ax_muster)
    writer.sheets['Sheet1'].set_row(8+i*3,cell_format=ax_pkt)
writer.sheets['Sheet1'].set_column(0,0,27, cell_format=idx)
writer.sheets['Sheet1'].write_string(9+10*3,0, exp_review_public['index'][4+10*3], cell_format=footer_i)
writer.sheets['Sheet1'].write_string(10+10*3,0, exp_review_public['index'][5+10*3], cell_format=footer_i)
writer.sheets['Sheet1'].write_string(11+10*3,0, exp_review_public['index'][6+10*3], cell_format=footer_i)
writer.sheets['Sheet1'].write_string(12+10*3,0, exp_review_public['index'][7+10*3], cell_format=footer_i)
writer.sheets['Sheet1'].write_string(13+10*3,0, exp_review_public['index'][8+10*3], cell_format=footer_i)
writer.sheets['Sheet1'].set_row(9+10*3,cell_format=footer)
writer.sheets['Sheet1'].set_row(10+10*3,cell_format=footer)
writer.sheets['Sheet1'].set_row(11+10*3,cell_format=footer)
writer.sheets['Sheet1'].set_row(12+10*3,cell_format=footer)
writer.sheets['Sheet1'].set_row(13+10*3,cell_format=footer)
writer.sheets['Sheet1'].set_row(14+10*3,cell_format=note)
writer.sheets['Sheet1'].repeat_rows(5)
writer.save()

# TODO: export errorlog and difflog

### Plot comparison of Exam_Pkt and ILIAS_Pkt #####
a = members['Exam_Pkt'].value_counts().sort_index()
b = members['ILIAS_Pkt'].value_counts().sort_index()
ab = pd.merge(a, b, how='outer', on=a.index)
ab[['ILIAS_Pkt','Exam_Pkt']].plot.bar()
