"""
ILIAS Evaluator: A tool for evaluating and post-correction of tests and exams 
done in the ILIAS or ILIAS E-Assessment System

The tool is based on 
https://github.com/P4ckP4ck/ILIAS_KlausurAuswertung/Bewerte ILIAS-Testergebnisse V1_5.py
first created by Eberhard Waffenschmidt, TH Köln 

For export of the tests and exams done in ILIAS or ILIAS E-Assessment System 
please use the tool
https://github.com/TPanteleit/ILIAS---Test-Generator

This tool consists of: 
    - classes:
        - Test
    - general methods:
        - eval_ILIAS
        - import_psso_members
        - evaluate_bonus
        - evaluate_exam

This tool is capable of: 
    - read and process ILIAS test results
    - evaluate intermediate tests of a course regarding bonus points
    - evaluate exam of a course
    - tasks with multiple answeres can be considered in evaluation
    - determine course note for each participant
    - export detailed results (e.g. for exam review for students)
TODO:
    - find better solution for processing bonus by bonus tests and Praktikum
    - complete automated statistical analysis of tests or an exam 

This tool is limited as follows: 
    - active evaluation only for task types: Formelfrage, Single Choice
    - passive ILIAS result import only for tasks which are in ILIAS import data
TODO: 
    - feasible extension for active evaluation for task types: Multiple Choice, 
        Freitextaufgabe
@author: Silvan Rummeny"""

__version__ = '2.0'
__author__ = 'srummeny'

import pandas as pd
import numpy as np
from math import *


class Test:
    """
    class of a test or exam to get evaluated  
    """
    def __init__(self, 
                 members: pd.DataFrame, 
                 marker: list, 
                 name: int or str, 
                 ilias_export: str, 
                 ff: str,
                 sc: str):
        """
        Parameters
        -----------
        members: pd.DataFrame
            DataFrame of all course members incl. Name, Matrikelnr., etc. 
        marker: list of str
            list of marker used to identify variables, results, etc. in 
            self.d_ilias
        name: int or str
            Number or name of test (e.g. number of intermediate test or name of
            exam cohort)
        ilias_export: str
            path of the ILIAS result and data import file
        ff: str
            path of Formelfrage task pool
        sc: str
            path of SingleChoice task pool
        """
        self.members = members
        self.marker = marker
        self.name = name
    ## 1. read task pool data
        self.ff = pd.read_excel(ff, sheet_name='Formelfrage - Database')
        self.sc = pd.read_excel(sc, sheet_name='SingleChoice - Database')
        print ("excel task pools OK")
        print ("read ILIAS-data...")
        df = pd.ExcelFile(ilias_export)  
    ## 2. Get aggregated results and data of ILIAS
        self.r_ilias = df.parse(df.sheet_names[0])
        # drop all empty rows of ILIAS_results until first name appears 
        while not self.r_ilias.loc[self.r_ilias.index[0]].any():
            self.r_ilias = self.r_ilias.drop(index=self.r_ilias.index[0])
        self.r_ilias = self.r_ilias.set_index('Name')
        # get important test parameters from aggregated ILIAS data
        self.n_tasks = int(self.r_ilias['Gesamtzahl der Fragen'][0])
        self.max_pts = self.r_ilias['Maximal erreichbare Punktezahl'][0]
    ## 3. Get ILIAS data of every single participant (detailed)
        # save sheet data of each participant 
        self.d_ilias = []  
        for i in range(len(df.sheet_names[1:])):   
            self.d_ilias.append(df.parse(df.sheet_names[i+1], header=None,
                    ignore_index=True, names=[df.sheet_names[i+1], 'values']))
        # initialize row finder of init and valid run of participant
        self.row_finder = pd.DataFrame(index = range(len(self.d_ilias)), 
                                columns = ['i_mem', 'row_init', 'row_valid'])
    ## 4. Initialize self. entries containing all task details of the participant
        # create MultiIndex 
        i_tasks = []
        i_sub = []
        subtitles = ['ID', 'Title', 'Type', 'Formula', 'Var', 'Tol', 'R', 
                     'R_ref', 'Pkt', 'Pkt_ILIAS']
        for n in range(self.n_tasks):
            i_tasks += ['A'+str(n+1)]*len(subtitles)
            i_sub += subtitles
        c_ent = pd.MultiIndex.from_arrays([i_tasks, i_sub], names = ['task', 
                                                                  'parameter'])
        self.ent = pd.DataFrame(index=self.members.index, columns = c_ent)


    def process_d_ilias(self):
        """ processes ILIAS Data for the test and saves it in self.ent for
        all members
        
        Mainly used Parameters of class Test
        ----------------
        self.d_ilias: list of pd.DataFrame
            ILIAS Data of entries, variables and results of every single member
            for a single Test
        self.r_ilias: pd.DataFrame
            Aggregated ILIAS results of every single member
        self.ent: pd.DataFrame
            Empty Output Dataframe for calling all single entries, variables 
            and results of all members and tests
        self.members: pd.DataFrame
            DataFrame of exported ILIAS members to get the i_mem for self.ent
        self.marker: list
            marker list to identify runs, tasks, variables or results
        """
    ## 1. Define marker to find important elements in txt
        run_marker = self.marker[0]     # run marker
        tasks = self.marker[1]          # task marker
        var_marker = self.marker[2]     # variable marker
        res_marker = self.marker[3]     # result marker
        res_marker_ft = self.marker[4]  # result marker of Freitextaufgabe
        # iterate every test participant
        for p in range(len(self.d_ilias)):  
    ## 2. Find name of participant and match it with self.members
            name = self.d_ilias[p].columns[0] 
            # print(p, name)
            # match name with self.members['Name_']
            if self.members['Name_'].str.contains(name).any():
                # check if name is in  self.members['Name']
                p_sel = self.members['Name_'].str.contains(name)
            else:
                # check if  self.members['Name'] is in name
                p_sel = [self.members['Name_'].values[i] in name for i in \
                         range(len(self.members))]
            # get member index in self.members 
            if any(p_sel):
                i_mem = self.members['Name_'][p_sel].index.values.item()
            else: 
                print('### Participant',p, name,'skipped, because it is not in ILIAS member list! ###')
                continue
            # (re-)set i_run and i_task 
            i_run = 0 
            i_task = 0
    ## 3. Get ILIAS data of participant and iterate every row to extract it
            # skip empty rows
            i_data = self.d_ilias[p][name].dropna().index.values 
            for i in i_data:   # iterate every Excel Cell 
                txt = self.d_ilias[p][name][i]
                """
                TODO: consider task IDs
                # id_title = title[0:title.find(" ")]      # FragenID
                # self.ent.loc[i_mem, (a_t, 'ID')]= id_title
                """
                # is there a new run?
                if txt.startswith(run_marker):    
                    i_run += 1
                    i_task = 1  # tasks start with integer 1
                # is there a new task?
                elif any([txt == tasks[j] for j in range(len(tasks))]):
                    i_task += 1
                    a_t = 'A'+str(i_task)
                    title = self.d_ilias[p]['values'][i]
                    self.ent.loc[i_mem, (a_t, 'Type')]= txt
                    self.ent.loc[i_mem, (a_t, 'Title')]= title
                    self.ent.loc[i_mem, (a_t, 'Var')]= [None]*15
                    self.ent.loc[i_mem, (a_t, 'R')]= []   # [None]*10
                # is there a new variable or result?
                elif (txt.startswith(var_marker) or
                    txt.startswith(res_marker) or
                    txt.startswith(res_marker_ft)):
                    # if there is a value for variable or result available
                    if ~self.d_ilias[p]['values'].isna()[i]:
                        if txt.startswith(var_marker):
                            var = self.d_ilias[p]['values'][i]
                            v_i = int(txt.replace(var_marker, ''))-1
                            self.ent.loc[i_mem, (a_t, 'Var')][v_i] = var
                        elif (txt.startswith(res_marker) or 
                              txt.startswith(res_marker_ft)):
                            r = self.d_ilias[p]['values'][i]
                            self.ent.loc[i_mem, (a_t, 'R')].append(r)
                    """
                    TODO: implement consideration of arbitrary result inputs
                    # r_i = int(txt.replace(res_marker, ''))-1
                    """
                else: 
                    # catch selected Single-Choice-Answeres (no marker used)
                    if self.d_ilias[p]['values'][i] == 1:
                        self.ent.loc[i_mem, (a_t, 'R')].append(txt)
    ## 4. Create self.row_finder of valid results according to ILIAS 
            try: # try to find row of Name in r_ilias (identical)
                row = self.r_ilias.index.get_loc(self.members['Name'][i_mem])
            except: # try to find row of Name in r_ilias (containing)
                names = self.r_ilias.index.dropna()
                name_sel = [self.members['Name'][i_mem] in names[i] for i in \
                            range(len(names))] 
                row = self.r_ilias.index.get_loc(names[name_sel].values.item())
            self.row_finder.loc[p, 'i_mem'] = i_mem
            self.row_finder.loc[p, 'row_init'] = row
        # find row of valid run of all participants
        for i in range(len(self.row_finder)):   
            row = self.row_finder['row_init'][i]
            if np.isnan(row):   # skip rows containing nan
                continue
            else:
                # find number of valid run in ILIAS_results
                i_val = self.r_ilias['Bewerteter Durchlauf'][row].astype(int)-1
                if i_val > 0: # only if first run =! valid run
                    try: # try to find row of next participant in ILIAS_results
                        next_row = self.row_finder['row_init'][i+1] 
                    except: # there is no next participant in ILIAS_results
                        next_row = None
                    # get row of valid_run according to i_val
                    i_run = self.r_ilias['Durchlauf'][row:next_row].values
                    sel_v = [float(i_val+1) == i_run[i] for i in range(len(i_run))]
                    i_bew = np.where(sel_v)[0].item()
                    self.row_finder.loc[i, 'row_valid'] = row + i_bew
                else: # first run == the valid run
                    self.row_finder.loc[i, 'row_valid'] = row  

    def process_Pools(self):
        """ process task pools, evaluate results and returns self.ent with 
        evaluated results
        
        Mainly used Parameters of class Test:
        ----------------
        self.ff: pd.DataFrame
            DataFrame of Formelfrage task pool
        self.sc: pd.DataFrame
            DataFrame of SingleChoice task pool
        self.ent: pd.DataFrame
            Dataframe containing all single entries, variables and results of 
            all members
        self.marker: list
            marker list to identify runs, tasks, variables or results
        """ 
        # iterate all participating members
        participants = self.ent.any(axis=1)
        for m in self.ent.index[participants].to_list(): 
            sel_m = self.row_finder['i_mem'] == float(m)
    ## 1. Get row in r_ilias of participants valid run
            row_r = self.row_finder['row_valid'][sel_m].values.item()
            if self.members['Benutzername'].isna()[m]:
                # print(row_r)
                self.members.loc[m,'Benutzername'] = self.r_ilias.iloc[row_r,0]
            # iterate every tasks
            for t in range(self.ent.columns.levshape[0]): 
                P = 0   # set task points to zero (default)
                a_t = 'A'+str(t+1)  # get task header
                input_res = []
    ## 2. Get task in task pools and/or in r_ilias
                # is task title in ff-task pool?
                sel_ff = self.ff['question_title']==self.ent[(a_t, 'Title')][m]
                # is task title in sc-task pool?
                sel_sc = self.sc['question_title']==self.ent[(a_t, 'Title')][m]
                # is the task title in represented in r_ilias?
                sel_c = self.r_ilias.columns == self.ent[(a_t, 'Title')][m]
                # proof if task title is unique
                if (len(self.ff[sel_ff]) > 1 or len(self.sc[sel_sc]) > 1):
                    print('### Task title "', self.ent[(a_t, 'Title')][m],
                          '" is not unique! ###')
                    continue    
            ## if task type is Formelfrage
                if any(sel_ff):
                    sel_formula = self.ff.loc[sel_ff, self.ff.columns.str.contains('formula')]
                    # initialize empty lists in following parameters
                    self.ent[(a_t, 'Formula')][m] = []
                    self.ent[(a_t, 'Tol')][m] = []
                    self.ent[(a_t, 'Pkt')][m] = []
                    self.ent[(a_t, 'R_ref')][m] = []
                    # iterate number of formulas/results of that task
                    for n in range(sum([sel_formula.iloc[0] != ' '][0])): 
                        formula = self.ff['res'+str(n+1)+'_formula'][sel_ff].astype(str).item()
                        tol = self.ff['res'+str(n+1)+'_tol'][sel_ff].item()
                        Var = self.ent[(a_t, 'Var')][m]
                        self.ent[(a_t, 'Formula')][m].append(formula)
                        self.ent[(a_t, 'Tol')][m].append(tol)
                        if (~self.ent[(a_t, 'Var')].isna()[m] and # if var not NaN
                            len(self.ent[(a_t, 'R')][m])>=n+1):   # if R-list is long enough
        ## 2.a evaluate Formelfrage task
                            R_ref = eval_ILIAS(formula, var=Var, res=input_res)
                            if R_ref == None: 
                                print('### Result of Member', str(m), 
                                      ', Task', str(t+1), 'is None! ###')
                            elif R_ref == 'not_valid':
                                # if there is a formula error, decide in favour of participant
                                P += self.ff['res'+str(n+1)+'_points'][sel_ff].item()
                            else:
                                input_res.append(R_ref)
                                self.ent[(a_t, 'R_ref')][m].append(R_ref)
                                R_min = min(R_ref*(1+tol/100),R_ref*(1-tol/100))
                                R_max = max (R_ref*(1+tol/100),R_ref*(1-tol/100)) 
                                R_entry = self.ent[(a_t, 'R')][m][n]
                                if type(R_entry)==str: # if the entered result = str (e.g. when fractions are entered) 
                                    R_entry = eval(R_entry)
                                if (R_entry >= R_min) and (R_entry <= R_max): # check correct result
                                    P += self.ff['res'+str(n+1)+'_points'][sel_ff].item()
                        else: # no result or vars available
                            continue
                    self.ent.loc[m, (a_t, 'Pkt')] = P
                    # if there are ILIAS_results for single task available
                    if any(sel_c): 
                        # get Points achieved according to ILIAS
                        self.ent.loc[m, (a_t, 'Pkt_ILIAS')] = self.r_ilias.iloc[row_r, sel_c].values.item()
            ## if task type is SingleChoice
                elif any(sel_sc):
                    # if there are ILIAS_results for single task available
                    if any(sel_c):
                        # get Points achieved according to ILIAS
                        self.ent.loc[m,(a_t, 'Pkt_ILIAS')] = self.r_ilias.iloc[row_r, sel_c].values.item()
                        # set Points to Points achieved according to ILIAS
                        self.ent.loc[m,(a_t, 'Pkt')] = self.ent[(a_t, 'Pkt_ILIAS')][m]
                    else: 
                        try: 
                            # get number of answeres
                            sel_text = self.sc.loc[sel_sc, self.sc.columns.str.contains('text')]
                            self.ent[(a_t, 'Formula')][m] = []
                            self.ent[(a_t, 'Tol')][m] = []
                            self.ent[(a_t, 'Pkt')][m] = []
                            self.ent[(a_t, 'R_ref')][m] = []
                            # iterate number of answer texts / results of that task
                            for n in range(sum([sel_text.iloc[0] != ' '][0])):
                            # get the correct answer and save it in self.ent
                                if self.sc['response_'+str(n+1)+'_pts'][sel_sc].values.item() > 0:
                                    text = self.sc['response_'+str(n+1)+'_text'][sel_sc].values.item()
                                    pts = self.sc['response_'+str(n+1)+'_pts'][sel_sc].values.item()
                                    self.ent[(a_t, 'Formula')][m].append(pts)
                                    self.ent[(a_t, 'R_ref')][m].append(text)
                                 
                            for ref in range(len(self.ent[(a_t, 'R_ref')][m])):
                                for r in range(len(self.ent[(a_t, 'R')][m])):
                                    if self.ent[(a_t, 'R')][m][r] in self.ent[(a_t, 'R_ref')][m][ref]:
        ## 2.b evaluate Single Choice task
                                        P += self.ent[(a_t, 'Formula')][m][ref]
                            self.ent.loc[m, (a_t, 'Pkt')] = P
                        except: 
                            print('### skipped Single Choice task', a_t, 'of member', m, '###')
                            self.ent.loc[m,(a_t, 'Pkt')] = 0
            ## if task is task for test correction
                elif self.ent[(a_t, 'Title')][m] == "Bitte ignorieren!":
                    self.ent.loc[m,(a_t, 'Pkt')] = 0
            ## if no task is identified
                else: 
                    # find not identified tasks in ILIAS_results
                    if any(sel_c): 
                        # get Points achieved according to ILIAS
                        self.ent.loc[m,(a_t, 'Pkt_ILIAS')] = self.r_ilias.iloc[row_r, sel_c].values.item()
                        self.ent.loc[m,(a_t, 'Pkt')] = self.ent[(a_t, 'Pkt_ILIAS')][m]
                    # contradiction! task doesn't exist
                    else:
                        print ("### Member", m, ", task", t+1, ",",
                               self.ent[(a_t, 'Title')][m],
                               "doesn't exist in task pool or ILIAS_results! ###")
        ## if there are no ILIAS results for single task available, take the aggregated ones
            if not any(sel_c):
                self.members.loc[m, 'Pkt_ILIAS'] = self.r_ilias.loc[
                        self.r_ilias.index[row_r], 'Testergebnis in Punkten']
         
def eval_ILIAS (formula_ILIAS: str, 
                var: list = [], 
                res: list = []):
    """Reformatting and calculation of ILIAS formula with given var and res 
    inputs, returning calculated result
    @author: srummeny, 21.6.21 (edited Version of E. Waffenschmidt, 3.9.2020)
              
    Parameters
    -----------
    formula_ILIAS: str
        formula string in ILIAS format
    var: list of float
        variable values as used as input in formula
    res: list of float
        result values as used as input in formula
    """
## 1. Reformat formula_ILIAS to formula_py  
    formula_py = formula_ILIAS.lower()
    for v in range(len(var)):
        var_ILIAS = '$v'+str(v+1)
        var_py = 'var['+str(v)+']'
        formula_py = formula_py.replace(var_ILIAS, var_py)
    for r in range(len(res)):
        res_ILIAS = '$r'+str(r+1)
        res_py = 'res['+str(r)+']'
        formula_py = formula_py.replace(res_ILIAS, res_py)
    # reformat math functions
    formula_py = formula_py.replace(",", ".")
    formula_py = formula_py.replace("^", "**")
    formula_py = formula_py.replace("arcsin", "asin")
    formula_py = formula_py.replace("arcsinh", "asinh")
    formula_py = formula_py.replace("arccos", "acos")
    formula_py = formula_py.replace("arccosh", "acosh")
    formula_py = formula_py.replace("arctan", "atan")
    formula_py = formula_py.replace("arctanh", "atanh")
    formula_py = formula_py.replace("ln", "log")
    formula_py = formula_py.replace("log", "log10")
## 2. Calculation of results, based on formula, var and res
    try: # successful calculation
        result = eval(formula_py)
    except ZeroDivisionError: # case for formula error
        result = 'not_valid'
        print('### ZeroDivisionError ocurred ###')
    except: # if not, result is None
        result = None
        print('### Formula', formula_py,'with var=', var, 'and res=', res,
              'could not be solved! ###')
    return result

def import_psso_members(psso_import: list):
    """import and concatenation of all psso members. Return one complete psso 
    member list, which is used for evaluation of an exam or course 
    
    Parameters
    ---------------
    psso_import: list of str
        path str list of the files containing psso_members 
    """
    # init psso_members by importing first psso_import file
    psso_members = pd.read_excel(psso_import[0], skiprows=3)
    # drop last row (summation row)
    psso_members = psso_members.drop(index=psso_members.index[-1])
    # import all left psso_import files
    if len(psso_import)>1:
        for i in range(len(psso_import)-1):
            new_members = pd.read_excel(psso_import[i+1], skiprows=3)
            new_members = new_members.drop(index=new_members.index[-1])
            psso_members = pd.concat([psso_members, new_members], axis=0, 
                                     ignore_index=True)
    return psso_members

def evaluate_bonus(members: pd.DataFrame, 
                   praktika: pd.DataFrame, 
                   bonus_tests: list = None,
                   bonus_ges: pd.DataFrame = None, 
                   scheme: pd.Series = None,
                   tests_p_bonus: int = 2, 
                   max_bonus: int = 5):
    """evaluation of bonus of a course and returns members['Bonus_Pkt'] 
    and full DataFrame of bonus_ges
    
    Parameters
    -------------
    members: pd.DataFrame
        DataFrame of all course members incl. Name, Matrikelnr., etc.
    praktika: pd.DataFrame
        DataFrame of bonus achieved from Praktika
    bonus_tests: list of class Test
        List of evaluated bonus tests
    bonus_ges: pd.DataFrame
        empty DataFrame containing, Ges_Pkt and Note from each bonus test
    scheme: pd.Series
        scheme for bonus test evaluation containing note str as index and 
        corresponding percentage limits as values 
    tests_p_bonus: int
        number of bonus tests to get 1 bonus point
    max_bonus: int
        maximum achievable bonus points
    """
    # iterate every member
    for p in members.index.to_list(): 
        # add Bonus_ZT of member only when there are bonus tests available
        if bonus_tests is not None: 
    ## 1. Evaluate bonus tests
            for t in bonus_tests:    
                ZT_i = 'ZT'+str(t.name)
                try:
                    # determine total points of bonus test
                    bonus_ges.loc[p, [(ZT_i,'Ges_Pkt')]] = sum(t.ent.loc[p,
                                                      pd.IndexSlice[:,'Pkt']]) 
                    # get bonus test note
                    if bonus_ges[(ZT_i,'Ges_Pkt')][p]/t.max_pts >= \
                                                            scheme.iloc[1]/100:
                        bonus_ges.loc[p,[(ZT_i,'Note')]] = scheme.index[1]
                    else:
                        bonus_ges.loc[p,[(ZT_i,'Note')]] = scheme.index[0]
                except: # evaluation of bonus test failed
                    print('### skipped Member', p, members.loc[p,'Name'], 
                          'in test', t.name, '###')
            test_res = bonus_ges.loc[p,pd.IndexSlice[:,'Note']].value_counts()
            if (test_res.index=='BE').any():
                # Get bonus by bonus tests by considering tests_p_bonus
                members.loc[p,'Bonus_ZT'] = floor(test_res['BE']/tests_p_bonus)
    ## 2. Evaluate bonus by Praktikum
        if any(praktika['Matrikelnr.']==members['Matrikelnummer'][p]):
            # match values by Matrikelnummer
            sel = praktika['Matrikelnr.']==members['Matrikelnummer'][p]
            members.loc[p,'Bonus_Pra'] = max(praktika['Summe'][sel])
    ## 3. Determine total bonus
            members.loc[p,'Bonus_Pkt'] = min(max_bonus, 
                                        np.nansum([members.loc[p,'Bonus_Pra'], 
                                                   members.loc[p,'Bonus_ZT']]))
        else: # member got no bonus by Praktikum 
            members.loc[p,'Bonus_Pkt'] = min(max_bonus, 
                                             members.loc[p,'Bonus_ZT'])
    return members, bonus_ges


def evaluate_exam(members: pd.DataFrame, 
                  exam: list,
                  scheme: pd.Series, 
                  max_pts: int = 40):
    """evaluation of the final exam and course note and returns members['Note']
    and all_entries (all exam cohort data unified)
    
    Parameters
    -------------
    members: pd.DataFrame
        DataFrame of all course members incl. Name, Matrikelnr., etc. 
    exam: list of class Test
        List of evaluated cohorts of the exam
    scheme: pd.Series
        scheme for exam evaluation containing note str as index and 
        corresponding percentage limits as values 
    max_pts: int
        maximum achievable exam points, used as reference for note scheme
    """
    # iterate every cohort
    for c in range(len(exam)): 
        participants = exam[c].ent.any(axis=1)
        # iterate every participant
        for p in exam[c].ent.index[participants].to_list(): 
            try: # to get the exam note
    ## 1. Determine reached Pkt by exam
                members.loc[p,'Exam_Pkt'] = sum(exam[c].ent.loc[p,
                                                       pd.IndexSlice[:,'Pkt']])
    ## 2. Determine reached Pkt including bonus
                members.loc[p,'Ges_Pkt'] = members.loc[p,'Exam_Pkt'] + \
                                           members.loc[p,'Bonus_Pkt']
    ## 3. Evaluate course note of participant
                n_sel = members.loc[p, 'Ges_Pkt']/max_pts*100 >= scheme
                members.loc[p,'Note'] = n_sel[n_sel==True].index[-1]
            except: 
                print('### Exam and note evaluation failed for Member', p, 
                      members.loc[p,'Name'], 'in test', exam[c].name, '###')
    ## 4. Create pd.DataFrame containing all entries of all cohorts together
        if c==0:
            all_entries = exam[c].ent.copy()
        else:
            # overwrite all_entries (usually nan-rows) by values of exam[c].ent
            all_entries.update(exam[c].ent)
    return members, all_entries
