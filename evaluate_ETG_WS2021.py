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
import_members = '2021_08_15_10-061629014791_member_export_37420.xlsx'
               # '2021_06_28_10-011624867289_member_export_2341446.xlsx'
               # '2021_06_14_08-401623652810_member_export_2341446.xlsx' # old file
psso_identifier = 'psso_alle'
result_identifier = '_results'
ff_pool_identifier = 'Formelfrage'
sc_pool_identifier = 'SingleChoice'
export_prefix = 'ETG_WS2021_'
Filename_Export_detailed = export_prefix+'exp_Ergebnisse_det.xlsx'
Filename_Export_public = export_prefix+'exp_Ergebnisse_pub.xlsx'
Filename_Export_PSSO = export_prefix+'exp_Ergebnissse_psso.xlsx'
Filename_Export_review_detailed = export_prefix+'exp_results_review_det.xlsx'
Filename_Export_review_public = export_prefix+'exp_results_review.xlsx'
IRT_Frame_Name = export_prefix+'irt_frame.xlsx'
name_marker = 'Ergebnisse von Testdurchlauf '   # 'Ergebnisse von Testdurchlauf 1 für '
run_marker = 'dummy_text'   # run marker currently not used
tasks = ['Formelfrage', 'Single Choice', 'Lückentextfrage', 'Hotspot/Imagemap', 'Freitext eingeben']
res_marker_ft = "Ergebnis"
var_marker = '$v'
res_marker = '$r'
marker = [run_marker, tasks, var_marker, res_marker, res_marker_ft] 

# Specific constants for members
mem_dir = 'ETG_WS2021_Members/'
psso_import = []
ilias_mem = pd.read_excel(mem_dir+import_members, 
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
exam_cohort = ['Exam']
exam_dir = 'ETG_WS2021_Exam/'

# find all import data and pools in directory
# psso members
for i in range(len(glob.glob(mem_dir+'/*.xls'))):
    if psso_identifier in glob.glob(mem_dir+'/*.xls')[i]:
        psso_import.append(glob.glob(mem_dir+'/*.xls')[i])
    else:
        print('### PSSO import - skipped file:', glob.glob(mem_dir+'/*.xls')[i])

# test data
# [zt_ilias_result, zt_pool_ff, zt_pool_sc] = ev.get_excel_files(zt_test, zt_dir)
# [pra_ilias_result, pra_pool_ff, pra_pool_sc] = ev.get_excel_files(pra_experiment, pra_dir)
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
## remove ' from Names to get ILIAS equivalent names
members['Name_'] = members['Name_'].str.replace("'", "")
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

## disable here
######### LOOP of evaluating all considered intermediate tests ############
#intermediate_tests = []
#for zt in range(len(zt_test)):
#    intermediate_tests.append([])
#    for sub in range(len(zt_ilias_result[zt])):
#        print('started evaluating intermediate test', zt_test[zt])
#        intermediate_tests[zt].append(ev.Test(members, marker, zt_test[zt],
#                                          zt_ilias_result[zt][sub], ff=zt_pool_ff[zt][sub], sc=zt_pool_sc[zt][sub]))
#        print("process ILIAS data...")
#        intermediate_tests[zt][sub].process_ilias()
#        print("process task pools and evaluate...")
#        intermediate_tests[zt][sub].process_pools()
#    
#print("evaluate zt bonus...")
#[members, course_data]= ev.evaluate_intermediate_tests(members, 
#                                                       zt_tests=intermediate_tests, 
#                                                       d_course=course_data,  
#                                                       scheme=zt_scheme)
############ LOOP of evaluating all considered Praktikum experiments ############
#praktikum = []
#for pra in range(len(pra_experiment)):
#    praktikum.append([])
#    for sub in range(len(pra_ilias_result[pra])):
#        print('started evaluating Praktikum test', pra_ilias_result[pra][sub][21:])
#        praktikum[pra].append(ev.Test(members, marker, pra_experiment[pra],
#                                 pra_ilias_result[pra][sub]))
#        print("process ILIAS data...")
#        praktikum[pra][sub].process_ilias()
#    
#print("evaluate pra bonus...")
#[members, course_data]= ev.evaluate_praktika(members, pra_prev=old_praktika, 
#                                             pra_tests=praktikum, 
#                                             d_course=course_data)
# Exception for Zeinab Mohammad: "doch im Testdurchlauf 2 bestanden, erfolgloser Testdurchlauf 3 wurde gewertet
course_data.loc[96, ('V2', 'Ges_Pkt')] = 2
course_data.loc[96, ('V2', 'Note')] = 'BE'

###### EVALUATE TOTAL BONUS ###########
members = ev.evaluate_bonus(members)

########## LOOP of evaluating all considered exam cohorts ############
exam = []
for c in range(len(exam_cohort)):
    print('started evaluating exam,', exam_cohort[c])
    exam.append(ev.Test(members, marker, exam_cohort[c],
                        exam_ilias_result[c][0], ff=exam_pool_ff[c][0], daIr=True))
    print("process ILIAS data...")
    exam[c].process_ilias()
    print("process task pools and evaluate...")
    exam[c].process_pools()
    
print("evaluate exam...")
[members, all_entries] = ev.evaluate_exam(members, exam, exam_scheme, max_pts=41) 
members['bewertung']=members['Note'].fillna('PNE')

#### Check number of Taxonomies in Exam #####
tax = []
for i in range(42):
    tax.append('%02d' % i)
tax[0] = 'Bitte ignorieren!'
tax[-1] = 'SC'
taxonomies = pd.DataFrame(index=members.index, columns=tax)
n_tax = pd.Series(index=range(11), dtype=object).fillna(0)
n_taxs = pd.DataFrame(index=range(11))
run_max = pd.Series(index=members.index, dtype=object)
for m in members['Note'].dropna().index.values:
    run = 0
    runs = []
    for i in range(len(tax)):
        if tax[i]=='SC': 
            taxonomies.loc[m, tax[i]] = len(tax)-taxonomies.loc[m].sum()
        else:
            taxonomies.loc[m, tax[i]] = sum(all_entries.loc[m,pd.IndexSlice[:,'Title']].str.startswith(str(tax[i])))
        if taxonomies.loc[m, tax[i]]==0:
            run += 1
        else:
            if run==0:
                continue
            else:
                runs.append(run)
                run = 0
    n_taxs[m] = n_tax.add(taxonomies.loc[m].value_counts().sort_index(), fill_value=0)
    run_max[m] = max(runs)
    # plt.plot(taxonomies.loc[m].value_counts().sort_index())
# v_max, n_0, n_1, n_2, n_3, n_4

######### Export for PSSO ############
exp_psso = members[members.columns[0:13]].rename(columns={'Matrikelnummer':'mtknr',
                                                          'Name':'sortname',
                                                          'Nachname':'nachname',
                                                          'Vorname':'vorname'})
exp_psso.to_excel(Filename_Export_PSSO, index=False)

######## Export for lecturer (detailed, not anonymous) ###########
cols_detailed = [0,13,14,15,10,7,16,17,18,19,20,22,21,4]
exp_detailed = members[members.columns[cols_detailed]].rename(columns={'Name_':'Name',
                                                                       'stg':'Studiengang',
                                                                       'pversuch':'Prüfungsversuch',
                                                                       'Bonus_Pkt':'Bonus_Ges',
                                                                       'Kohorte':'Exam_Kohorte',
                                                                       'ILIAS_Pkt':'Exam_ILIAS_Pkt',
                                                                       'bewertung':'Bewertung'})
exp_detailed = exp_detailed.sort_values(by=['Matrikelnummer'])    
exp_detailed.to_excel(Filename_Export_detailed, index=False)

######## Export for participants (short, anonymous) ############
cols_public = [0,23]
exp_public = members[members.columns[cols_public]].rename(columns={'bewertung':'Bewertung'})
exp_public = exp_public.sort_values(by=['Matrikelnummer'])
exp_public = exp_public[~exp_public['Note'].isnull()]
exp_public.to_excel(Filename_Export_public, index=False)

######## Export for Exam review (detailed and short) ############
# TODO: complete export for Exam review
exam_export = all_entries.copy()
exam_export.columns = exam_export.columns.map('_'.join)
exam_export = exam_export[~exam_export['A1_Title'].isnull()]
idx = exam_export.index
review = pd.concat([members.loc[idx,['Matrikelnummer','Name_']], exam_export, members.loc[idx,['Bonus_Pkt','Ges_Pkt','Note']]],axis=1)
for i in range(42):
    review = review.drop(columns=['A'+str(i+1)+'_ID',
                                  'A'+str(i+1)+'_Type', 
                                  'A'+str(i+1)+'_Pkt_ILIAS'])
review = review.rename(columns={'Name_':'Name','Bonus_Pkt':'Bonuspunkte', 'Ges_Pkt':'Gesamtpunkte'})
exp_review_detailed = review.copy()
exp_review_detailed.index = exp_review_detailed['Matrikelnummer']
exp_review_detailed = exp_review_detailed.sort_values(by=['Matrikelnummer'])
exp_review_detailed = exp_review_detailed.transpose()
exp_review_detailed.to_excel(Filename_Export_review_detailed, header=False)

exp_review_public = review.copy()
for i in range(42):
    exp_review_public = exp_review_public.drop(columns=['A'+str(i+1)+'_Title',
                                                        'A'+str(i+1)+'_Formula',
                                                        'A'+str(i+1)+'_Var',
                                                        'A'+str(i+1)+'_Tol'])
exp_review_public = exp_review_public.drop(columns=['Name'])
exp_review_public = exp_review_public.sort_values(by=['Matrikelnummer'])
exp_review_public = exp_review_public.transpose()
exp_review_public.to_excel(Filename_Export_review_public, header=False)
#print ('export results as excel...')
#drop_columns = ['Rolle/Status', 'Nutzungsvereinbarung akzeptiert', 'Name_']
#members.drop(drop_columns, axis=1).to_excel(Filename_Export, index=False, na_rep='N/A')
#print ('### done! ###')

print("Excel Export OK")

#### Plot comparison of Exam_Pkt and ILIAS_Pkt #####
#a = members['Exam_Pkt'].value_counts().sort_index()
#b = members['ILIAS_Pkt'].value_counts().sort_index()
#ab = pd.merge(a, b, how='outer', on=a.index)
#ab[['ILIAS_Pkt','Exam_Pkt']].plot.bar()

print ('### done! ###')
