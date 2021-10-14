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
    - tasks with multiple answers can be considered in evaluation
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

__version__ = '2.1'
__author__ = 'srummeny'

import pandas as pd
import numpy as np
from math import *
import glob


class Test:
    """
    class of a test or exam to get evaluated  
    """

    def __init__(self,
                 members: pd.DataFrame,
                 marker: list,
                 name: int or str,
                 ilias_export: str,
                 ff: str=None,
                 sc: str=None,
                 daIr: bool=False):
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
            path of Formelfrage task pool, default is None
        sc: str
            path of SingleChoice task pool, default in None
        daIr: bool
            document aggregated ILIAS results? - default is False
        """
        self.members = members
        self.marker = marker
        self.name = name
        # 1. read task pool data
        if ff is not None:
            self.ff = pd.read_excel(ff, sheet_name='Formelfrage - Database')
        else:
            self.ff = None
        if sc is not None:
            self.sc = pd.read_excel(sc, sheet_name='SingleChoice - Database')
        else: 
            self.sc = None
        print("excel task pools OK")
        print("read ILIAS-data...")
        df = pd.ExcelFile(ilias_export)
        # 2. Get aggregated results and data of ILIAS
        self.r_ilias = df.parse(df.sheet_names[0])
        self.doc_aggr_ILIAS_results = daIr
        # drop all empty rows of ILIAS_results until first name appears 
        while not self.r_ilias.loc[self.r_ilias.index[0]].any():
            self.r_ilias = self.r_ilias.drop(index=self.r_ilias.index[0])
        self.r_ilias = self.r_ilias.set_index('Name')
        # get important test parameters from aggregated ILIAS data
        self.n_tasks = int(self.r_ilias['Gesamtzahl der Fragen'].dropna()[0])
        self.max_pts = self.r_ilias['Maximal erreichbare Punktezahl'][0]
        # 3. Get ILIAS data of every single participant (detailed)
        # save sheet data of each participant 
        self.d_ilias = []
        for i in range(len(df.sheet_names[1:])):
            self.d_ilias.append(df.parse(df.sheet_names[i + 1], header=None, ignore_index=True,
                                         names=[df.sheet_names[i + 1], 'values']))
        # initialize row finder of init and valid run of participant
        self.row_finder = pd.DataFrame(index=range(len(self.d_ilias)), columns=['i_mem', 'row_init', 'row_valid'])
        # 4. Initialize self. entries containing all task details of the participant
        # create MultiIndex 
        i_tasks = []
        i_sub = []
        subtitles = ['ID', 'Title', 'Type', 'Formula', 'Var', 'Tol', 'R',
                     'R_ref', 'Pkt', 'Pkt_ILIAS']
        for n in range(self.n_tasks):
            i_tasks += ['A' + str(n + 1)] * len(subtitles)
            i_sub += subtitles
        c_ent = pd.MultiIndex.from_arrays([i_tasks, i_sub], names=['task', 'parameter'])
        self.ent = pd.DataFrame(index=self.members.index, columns=c_ent)

    def process_ilias(self):
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
        # 1. Define marker to find important elements in txt
        run_marker = self.marker[0]  # run marker
        tasks = self.marker[1]  # task marker
        var_marker = self.marker[2]  # variable marker
        res_marker = self.marker[3]  # result marker
        res_marker_ft = self.marker[4]  # result marker of Freitextaufgabe
        # iterate every test participant
        for p in range(len(self.d_ilias)):
            name = self.d_ilias[p].columns[0]
            # print(p, name)
            # 2. Find name of participant and match it with self.members
            # match name with self.members['Name_']
            if self.members['Name_'].str.contains(name).any():
                # check if name is in  self.members['Name_']
                p_sel = self.members['Name_'].str.contains(name)
            else:
                # check if  self.members['Name_'] is in name
                p_sel = [self.members['Name_'].values[i] in name for i in range(len(self.members))]
            # get member index in self.members 
            if any(p_sel):
                i_mem = self.members['Name_'][p_sel].index.values.item()
            else:
                print('### Skipped participant', p, name, ', because it is not in PSSO member list! ###')
                continue
            # (re-)set i_run and i_task 
            i_run = 0
            i_task = 0
            # 3. Get ILIAS data of participant and iterate every row to extract it
            # skip empty rows
            i_data = self.d_ilias[p][name].dropna().index.values
            for i in i_data:  # iterate every Excel Cell
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
                    a_t = 'A' + str(i_task)
                    title = self.d_ilias[p]['values'][i]
                    self.ent.loc[i_mem, (a_t, 'Type')] = txt
                    self.ent.loc[i_mem, (a_t, 'Title')] = title
                    self.ent.loc[i_mem, (a_t, 'Var')] = [None] * 15
                    self.ent.loc[i_mem, (a_t, 'R')] = []  # [None]*10
                    self.R_sc = [[],[]]
                # is there a new variable or result?
                elif (txt.startswith(var_marker) or
                      txt.startswith(res_marker) or
                      txt.startswith(res_marker_ft)):
                    # if there is a value for variable or result available
                    if ~self.d_ilias[p]['values'].isna()[i]:
                        if txt.startswith(var_marker):
                            var = self.d_ilias[p]['values'][i]
                            v_i = int(txt.replace(var_marker, '')) - 1
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
                    self.ent.loc[i_mem, (a_t, 'R')] = [[],[]]
                    pts = self.d_ilias[p]['values'][i]
                    self.R_sc[0].append(pts)
                    self.R_sc[1].append(txt) 
                    self.ent.loc[i_mem, (a_t, 'R')] = self.R_sc
            # 4. Create self.row_finder of valid results according to ILIAS
            try:  # try to find row of Name in r_ilias (identical)
                row = self.r_ilias.index.get_loc(name)
            except KeyError:  # try to find row of Name in r_ilias (containing)
                if name== "Tiendo Nzako, Elito Dauvillier":
                    name = "Tiendo Nzako, Elito D'auvillier"
                names = self.r_ilias.index.dropna()
                name_sel = [name in names[i] for i in range(len(names))]
                row = self.r_ilias.index.get_loc(names[name_sel].values.item())
            self.row_finder.loc[p, 'i_mem'] = i_mem
            self.row_finder.loc[p, 'row_init'] = row
        # find row of valid run of all participants
        for i in range(len(self.row_finder)):
            row = self.row_finder['row_init'][i]
            if np.isnan(row):  # skip rows containing nan
                continue
            else:
                # find number of valid run in ILIAS_results
                i_val = self.r_ilias['Bewerteter Durchlauf'][row].astype(int) - 1
                if i_val > 0:  # only if first run =! valid run
                    # get row of valid_run according to i_val
                    i_run = self.r_ilias['Durchlauf'][row:].values
                    j=0
                    while float(i_val + 1) != i_run[j]:
                        j += 1
                    self.row_finder.loc[i, 'row_valid'] = row + j
                else:  # first run == the valid run
                    self.row_finder.loc[i, 'row_valid'] = row

    def process_pools(self):
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
        # iterate all participating members and skip not participating members
        participants = self.ent.any(axis=1)
        for m in self.ent.index[participants].to_list():
            sel_m = self.row_finder['i_mem'] == float(m)
        # 1. Get row in r_ilias of participants valid run
            row_r = self.row_finder['row_valid'][sel_m].values.item()
            row_i = self.row_finder['row_init'][sel_m].values.item()
            # get ilias benutzername from r_ilias if it is NaN
            if self.members['Benutzername'].isna()[m]:
                # print(row_r)
                self.members.loc[m, 'Benutzername'] = self.r_ilias.iloc[row_r, 0]
            # iterate every tasks
            for t in range(self.ent.columns.levshape[0]):
                pkt = 0  # set task points to zero (default)
                a_t = 'A' + str(t + 1)  # get task header
                input_res = []
        # 2. Get task in task pools and/or in r_ilias
                # is task title in ff-task pool?
                sel_ff = self.ff['question_title'] == self.ent[(a_t, 'Title')][m]
                # proof if task title is unique
                if len(self.ff[sel_ff]) > 1:
                    print('### Task title "', self.ent[(a_t, 'Title')][m],
                              '" is not unique! ###')
                # set default sel_sc to a list of False
                sel_sc = [False, False]
                if self.sc is not None:
                    # is task title in sc-task pool?
                    sel_sc = self.sc['question_title'] == self.ent[(a_t, 'Title')][m]
                    sel_sc12 = None
                    # proof if task title is unique
                    if len(self.sc[sel_sc]) > 1:
                        print('### Task title "', self.ent[(a_t, 'Title')][m],
                              '" is not unique! ###')
                        # added special case for exam of 13.09.2021
                        if len(self.sc[sel_sc])==2:
                            sel_sc12 = [sel_sc.copy(), sel_sc.copy()]
                            sel_sc12[0][8] = False
                            sel_sc12[1][7] = False
                        else:
                            continue
                # is the task title in represented in r_ilias?
                sel_c = self.r_ilias.columns == self.ent[(a_t, 'Title')][m]
                
            # if task type is Formelfrage
                if any(sel_ff):
                    sel_formula = self.ff.loc[sel_ff, self.ff.columns.str.contains('formula')]
                    # initialize empty lists in following parameters
                    self.ent[(a_t, 'Formula')][m] = []
                    self.ent[(a_t, 'Tol')][m] = []
                    self.ent[(a_t, 'Pkt')][m] = []
                    self.ent[(a_t, 'R_ref')][m] = []
                    # iterate number of formulas/results of that task
                    for n in range(sum([sel_formula.iloc[0] != ' '][0])):
                        formula = self.ff['res' + str(n + 1) + '_formula'][sel_ff].astype(str).item()
                        tol = self.ff['res' + str(n + 1) + '_tol'][sel_ff].item()
                        prec = self.ff['res' + str(n + 1) + '_prec'][sel_ff].item()
                        if not np.isnan(prec):
                            prec = int(prec)
                        var = self.ent[(a_t, 'Var')][m]
                        self.ent[(a_t, 'Formula')][m].append(formula)
                        self.ent[(a_t, 'Tol')][m].append(tol)
                        if (~self.ent[(a_t, 'Var')].isna()[m] and  # if var not NaN
                                len(self.ent[(a_t, 'R')][m]) >= n + 1):  # if R-list is long enough
                            r_ref = eval_ilias(formula, var=var, res=input_res)
        # 2.a evaluate Formelfrage task
                            if r_ref is None:
                                print('### Result of Member', str(m),
                                      ', Task', str(t + 1), 'is None! ###')
                            elif r_ref == 'not_valid':
                                # if there is a formula error, decide in favour of participant
                                pkt += self.ff['res' + str(n + 1) + '_points'][sel_ff].item()
                            else:
                                input_res.append(r_ref)
                                self.ent[(a_t, 'R_ref')][m].append(r_ref)
                                # find r_min/r_max according to tol
                                r_min = min(r_ref * (1 + tol / 100), r_ref * (1 - tol / 100))
                                r_max = max(r_ref * (1 + tol / 100), r_ref * (1 - tol / 100))
                                # set r_min/r_max to minimal/maximal value considering prec
                                r_min = min(floor(r_min * 10 ** prec) / 10 ** prec, r_min)
                                r_max = max(round(r_max, prec), r_max)
                                r_entry = self.ent[(a_t, 'R')][m][n]
                                # if the entered result = str (e.g. when fractions are entered)
                                if type(r_entry) == str:
                                    r_entry = eval(r_entry)
                                # update r_entry considering prec
                                r_ent = floor(r_entry * 10 ** prec) / 10 ** prec
                                if (r_ent >= r_min) and (r_ent <= r_max):  # check correct result
                                    pkt += self.ff['res' + str(n + 1) + '_points'][sel_ff].item()
                        else:  # no result or vars available
                            continue
                    self.ent.loc[m, (a_t, 'Pkt')] = pkt
                    # if there are ILIAS_results for single task available
                    if any(sel_c):
                        # get Points achieved according to ILIAS
                        self.ent.loc[m, (a_t, 'Pkt_ILIAS')] = self.r_ilias.iloc[row_r, sel_c].values.item()
            # if task type is SingleChoice
                elif any(sel_sc):
                    # if there are ILIAS_results for single task available
                    if any(sel_c):
                        # get Points achieved according to ILIAS
                        self.ent.loc[m, (a_t, 'Pkt_ILIAS')] = self.r_ilias.iloc[row_r, sel_c].values.item()
                        # set Points to Points achieved according to ILIAS
                        self.ent.loc[m, (a_t, 'Pkt')] = self.ent[(a_t, 'Pkt_ILIAS')][m]
                    else:
                        # get number of answeres                        
                        # # if the single choice title is ambiguous
                        if sel_sc12 is not None:
                            pkts = [0, 0]
                            for sel_sc in sel_sc12:
                                sel_pts = self.sc.loc[sel_sc, self.sc.columns.str.contains('_pts')]
                                self.ent[(a_t, 'Formula')][m] = [[],[]]
                                self.ent[(a_t, 'Tol')][m] = []
                                self.ent[(a_t, 'Pkt')][m] = []
                                self.ent[(a_t, 'R_ref')][m] = []
                                # iterate number of answer texts / results of that task
                                for n in range(len(sel_pts.squeeze().dropna())):
                                    # get the correct answer and save it in self.ent
                                    text = self.sc['response_' + str(n + 1) + '_text'][sel_sc].values.item()
                                    pts = self.sc['response_' + str(n + 1) + '_pts'][sel_sc].values.item()
                                    self.ent[(a_t, 'Formula')][m][0].append(pts)
                                    self.ent[(a_t, 'Formula')][m][1].append(text)
                                    # pick the correct answer
                                    if pts > 0:
                                        self.ent[(a_t, 'R_ref')][m].append(text)
                                # if there is no text available for Single-Choice options
                                if any(self.ent[(a_t, 'Formula')][m][1][i]==' ' for i in range(len(self.ent[(a_t, 'Formula')][m][1]))):
                                    # compare pts patterns in Formula and R, e.g. [0, 0, 1] vs. [0, 0, 1] --> correct!
                                    if self.ent[(a_t, 'Formula')][m][0] == self.ent[(a_t, 'R')][m][0]:
                                        # select the correct single choice task of the abiguous ones and evaluate this one
                                        ind = [i for i in range(len(sel_sc12)) if all(sel_sc==sel_sc12[i])]
                                        pkts[ind[0]] = sum(self.ent[(a_t, 'Formula')][m][0])
                                # if there is text available for Single-Choice options
                                else:
                                    for ref in range(len(self.ent[(a_t, 'R_ref')][m])):
                                        # filter correct answers
                                        results = [r for i, r in enumerate(self.ent[(a_t, 'R')][m][1]) if self.ent[(a_t, 'R')][m][0][i]>1]
                                        for r in range(len(results)):
                                            if self.ent[(a_t, 'R')][m][r] in self.ent[(a_t, 'R_ref')][m][ref]:
                                                ind = [i for i in range(len(sel_sc12)) if all(sel_sc==sel_sc12[i])]
                                                pkts[ind[0]] += self.ent[(a_t, 'Formula')][m][ref]
        # 2.b evaluate Single Choice task
                            pkt = max(pkts) 
                            self.ent.loc[m, (a_t, 'Pkt')] = pkt
                # if the single choice title is unique
                        else:
                            sel_pts = self.sc.loc[sel_sc, self.sc.columns.str.contains('_pts')]
                            self.ent[(a_t, 'Formula')][m] = [[],[]]
                            self.ent[(a_t, 'Tol')][m] = []
                            self.ent[(a_t, 'Pkt')][m] = []
                            self.ent[(a_t, 'R_ref')][m] = []
                            # iterate number of answer texts / results of that task
                            for n in range(len(sel_pts.squeeze().dropna())):
                                # get the correct answer and save it in self.ent
                                text = self.sc['response_' + str(n + 1) + '_text'][sel_sc].values.item()
                                pts = self.sc['response_' + str(n + 1) + '_pts'][sel_sc].values.item()
                                self.ent[(a_t, 'Formula')][m][0].append(pts)
                                self.ent[(a_t, 'Formula')][m][1].append(text)
                                # pick the correct answer
                                if pts > 0:
                                    self.ent[(a_t, 'R_ref')][m].append(text)
                            # if there is no text available for Single-Choice options
                            if any(self.ent[(a_t, 'Formula')][m][1][i]==' ' for i in range(len(self.ent[(a_t, 'Formula')][m][1]))):
                                # compare pts patterns in Formula and R, e.g. [0, 0, 1] vs. [0, 0, 1] --> correct!
                                if self.ent[(a_t, 'Formula')][m][0] == self.ent[(a_t, 'R')][m][0]:
                                    pkt = sum(self.ent[(a_t, 'Formula')][m][0])
                            # if there is text available for Single-Choice options
                            else:
                                for ref in range(len(self.ent[(a_t, 'R_ref')][m])):
                                    # filter correct answers
                                    results = [r for i, r in enumerate(self.ent[(a_t, 'R')][m][1]) if self.ent[(a_t, 'R')][m][0][i]>1]
                                    for r in range(len(results)):
                                        if self.ent[(a_t, 'R')][m][r] in self.ent[(a_t, 'R_ref')][m][ref]:
                                            pkt += self.ent[(a_t, 'Formula')][m][ref]
    # 2.b evaluate Single Choice task
                        self.ent.loc[m, (a_t, 'Pkt')] = pkt
            # if task is task for test correction
                elif self.ent[(a_t, 'Title')][m] == "Bitte ignorieren!":
                    self.ent.loc[m, (a_t, 'Pkt')] = 0
            # if no task is identified
                else:
                    # find not identified tasks in ILIAS_results
                    if any(sel_c):
                        # get Points achieved according to ILIAS
                        self.ent.loc[m, (a_t, 'Pkt_ILIAS')] = self.r_ilias.iloc[row_r, sel_c].values.item()
                        self.ent.loc[m, (a_t, 'Pkt')] = self.ent[(a_t, 'Pkt_ILIAS')][m]
                    # contradiction! task doesn't exist
                    else:
                        print("### Member", m, ", task", t + 1, ",",
                              self.ent[(a_t, 'Title')][m],
                              "doesn't exist in task pool or ILIAS_results! ###")
            # take the aggregated ILIAS results
            if self.doc_aggr_ILIAS_results:
                self.members.loc[m, 'ILIAS_Pkt'] = self.r_ilias.loc[self.r_ilias.index[row_i], 'Testergebnis in Punkten']


def eval_ilias(formula_ilias: str,
               var=None,
               res=None):
    """Reformatting and calculation of ILIAS formula with given var and res 
    inputs, returning calculated result
    @author: srummeny, 21.6.21 (edited Version of E. Waffenschmidt, 3.9.2020)
              
    Parameters
    -----------
    formula_ilias: str
        formula string in ILIAS format
    var: list of float
        variable values as used as input in formula
    res: list of float
        result values as used as input in formula
    """
# 1. Reformat formula_ilias to formula_py
    if res is None:
        res = []
    if var is None:
        var = []
    formula_py = formula_ilias.lower()
    for v in range(len(var)):
        var_ilias = '$v' + str(v + 1)
        var_py = 'var[' + str(v) + ']'
        formula_py = formula_py.replace(var_ilias, var_py)
    for r in range(len(res)):
        res_ilias = '$r' + str(r + 1)
        res_py = 'res[' + str(r) + ']'
        formula_py = formula_py.replace(res_ilias, res_py)
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
# 2. Calculation of results, based on formula, var and res
    try:  # successful calculation
        result = eval(formula_py)
    except ZeroDivisionError:  # case for formula error
        result = 'not_valid'
        print('### ZeroDivisionError ocurred ###')
    except SyntaxError:  # if not, result is None
        result = None
        print('### Formula', formula_py, 'with var=', var, 'and res=', res,
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
    if len(psso_import) > 1:
        for i in range(len(psso_import) - 1):
            new_members = pd.read_excel(psso_import[i + 1], skiprows=3)
            new_members = new_members.drop(index=new_members.index[-1])
            psso_members = pd.concat([psso_members, new_members], axis=0,
                                     ignore_index=True)
    return psso_members

def get_excel_files(considered_tests: list,
                    import_dir: str, 
                    identifier: list=['_results', 'Formelfrage', 'SingleChoice']):
    """ algorithm to collect all excel input data for several considered tests
    
    Parameters
    -------------------
    considered_tests: list
        list of names of the considered tests (e.g. [1, 2, 4, 7] or ['Test1, Test_xy, Test_final'])
    import_dir: str
        directory containing import data
    identifier: list
        list of identifiers for ILIAS result files ('_results'), Formelfrage or Single Choice task pools
    """
    # init outputs
    result_files = []
    pool_ff_files = []
    pool_sc_files = []
    for j in considered_tests:
        j_i = considered_tests.index(j)
        result_files.append([])
        pool_ff_files.append([])
        pool_sc_files.append([])
    # iterate all found files
        for i in range(len(glob.glob(import_dir+str(j)+'/*.xlsx'))):
            file = glob.glob(import_dir+str(j)+'/*.xlsx')[i]
            print(file)
            if identifier[0] in file or identifier[1] in file or identifier[2] in file:
                if identifier[0] in file:
                    result_files[j_i].append(file)
                elif identifier[1] in file:
                    pool_ff_files[j_i].append(file)
                elif identifier[2] in file:
                    pool_sc_files[j_i].append(file)
            else:
                result_files[j_i].append(None)
    return result_files, pool_ff_files, pool_sc_files


def evaluate_intermediate_tests(members: pd.DataFrame,
                                zt_tests: list = None,
                                d_course: pd.DataFrame = None,
                                scheme: pd.Series = None,
                                tests_p_bonus: int = 2):
    """evaluation of intermediate test bonus of a course and returns members['Bonus_ZT'] 
    and full DataFrame of bonus_ges
    
    Parameters
    -------------
    members: pd.DataFrame
        DataFrame of all course members incl. Name, Matrikelnr., etc.
    zt_tests: list of class Test
        List of evaluated intermediate tests
    d_course: pd.DataFrame
        empty DataFrame containing, Ges_Pkt and Note from each bonus test
    scheme: pd.Series
        scheme for intermediate test evaluation containing note str as index and 
        corresponding percentage limits as values 
    tests_p_bonus: int
        number of bonus tests to get 1 bonus point
    """
# 1.a Evaluate intermediate tests
    zt_filter = [col for col in d_course.columns.get_level_values(0) if col.startswith('ZT')]
    if zt_tests is not None:
        # iterate every intermediate test
        for zt in range(len(zt_tests)):
            # iterate every test of one intermediate test
            for t in zt_tests[zt]:
                zt_i = 'ZT' + str(t.name)
                # iterate every participating member
                for p in t.row_finder['i_mem'].dropna().values:
                    try:
                        # determine total points of bonus test
                        d_course.loc[p, [(zt_i, 'Ges_Pkt')]] = sum(t.ent.loc[p, pd.IndexSlice[:, 'Pkt']])
                        # get bonus test note
                        if d_course[(zt_i, 'Ges_Pkt')][p] / t.max_pts >= \
                                scheme.iloc[1] / 100:
                            d_course.loc[p, [(zt_i, 'Note')]] = scheme.index[1]
                        else:
                            d_course.loc[p, [(zt_i, 'Note')]] = scheme.index[0]
                    except TypeError:  # evaluation of bonus test failed
                        print('### skipped Member', p, members.loc[p, 'Name'], 'in test', t.name, '###')
                    test_res = d_course.loc[p, pd.IndexSlice[zt_filter, 'Note']].value_counts()
                    if (test_res.index == 'BE').any():
                        # Get bonus by bonus tests by considering tests_p_bonus
                        members.loc[p, 'Bonus_ZT'] = floor(test_res['BE'] / tests_p_bonus)
    return members, d_course


def evaluate_praktika(members: pd.DataFrame,
                      pra_prev: pd.DataFrame,
                      pra_tests: list = None,
                      d_course: pd.DataFrame = None,
                      scheme: pd.Series = None,
                      tests_p_bonus: int = 1):
    """evaluation of praktikum bonus of a course and returns members['Bonus_Pra'] 
    and full DataFrame of bonus_ges
    
    Parameters
    -------------
    members: pd.DataFrame
        DataFrame of all course members incl. Name, Matrikelnr., etc.
    pra_prev: pd.DataFrame
        DataFrame of bonus achieved from Praktika of previous semesters
    pra_tests: list of class Test
        List of evaluated Praktika tests
    d_course: pd.DataFrame
        empty DataFrame containing, Ges_Pkt and Note from each bonus test
    scheme: pd.Series
        scheme for Praktika test evaluation containing note str as index and 
        corresponding percentage limits as values 
    tests_p_bonus: int
        number of bonus tests to get 1 bonus point
    """        
    """
    TODO: apply adjustments for praktikum evaluation:
        - write new praktikum in old praktikum
    """
# 1.a Evaluate Praktika tests
    pra_filter = [col for col in d_course.columns.get_level_values(0) if col.startswith('V')]
    if pra_tests is not None:
        # iterate every experiment 
        for exp in range(len(pra_tests)):
            # iterate every test of one experiment
            for t in pra_tests[exp]:
                # iterate every participating member of test
                for p in t.row_finder['i_mem'].dropna().values:
                    pra_i = 'V' + str(t.name)
                    # do an own evaluation, when a scheme is given 
                    if scheme is not None: 
                        try:
                            # determine total points of bonus test
                            d_course.loc[p, [(pra_i, 'Ges_Pkt')]] = sum(t.ent.loc[p, pd.IndexSlice[:, 'Pkt']])
                            # get bonus test note
                            if d_course[(pra_i, 'Ges_Pkt')][p] / t.max_pts >= \
                                    scheme.iloc[1] / 100:
                                d_course.loc[p, [(pra_i, 'Note')]] = scheme.index[1]
                            else:
                                d_course.loc[p, [(pra_i, 'Note')]] = scheme.index[0]
                        except TypeError:  # evaluation of bonus test failed
                            print('### skipped Member', p, members.loc[p, 'Name'], 'in praktikum experiment', t.name, '###')
                    # if scheme is None, take the evaluation of the ILIAS system
                    else: 
                        sel_m = t.row_finder['i_mem'] == float(p)
                        # 1. Get row in r_ilias of participants valid run))
                        row_i = t.row_finder['row_init'][sel_m].values.item() 
                        d_course.loc[p, [(pra_i, 'Ges_Pkt')]] = t.r_ilias.loc[t.r_ilias.index[row_i], 'Testergebnis in Punkten']
                        if t.r_ilias.loc[t.r_ilias.index[row_i], 'Testergebnis als Note']=='bestanden': 
                            d_course.loc[p, [(pra_i, 'Note')]] = 'BE'
                        else:
                            d_course.loc[p, [(pra_i, 'Note')]] = 'NB'
                test_res = d_course.loc[p, pd.IndexSlice[pra_filter, 'Note']].value_counts()
                if (test_res.index == 'BE').any():
                # Get bonus by bonus tests by considering tests_p_bonus
                    members.loc[p, 'Bonus_Pra'] = floor(test_res['BE'] / tests_p_bonus)
    else:
# 1.b Or get bonus by old Praktikum
        # iterate every member
        for p in members.index.to_list():
            if any(pra_prev['Matrikelnr.'] == members['Matrikelnummer'][p]):
                # match values by Matrikelnummer
                sel = pra_prev['Matrikelnr.'] == members['Matrikelnummer'][p]
                members.loc[p, 'Bonus_Pra'] = max(pra_prev['Summe'][sel])
    return members, d_course


def evaluate_bonus(members: pd.DataFrame,
                   max_bonus: int = 5):
    """evaluation of total bonus points of a course and returns members['Bonus_Pkt'] 
    
    Parameters
    -------------
    members: pd.DataFrame
        DataFrame of all course members incl. Name, Matrikelnr., etc.
    max_bonus: int
        maximum achievable bonus points
    """
    # iterate every member
    for p in members.index.to_list():
    # 3. Determine total bonus
        members.loc[p, 'Bonus_Pkt'] = min(max_bonus,
                                          np.nansum([members.loc[p, 'Bonus_Pra'],
                                                     members.loc[p, 'Bonus_ZT']]))
    return members


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
    all_entries = exam[0].ent.copy()
    # iterate every cohort
    for c in range(len(exam)):
        participants = exam[c].ent.any(axis=1)
        # iterate every participant
        for p in exam[c].ent.index[participants].to_list():
            members.loc[p, 'Kohorte'] = exam[c].name
            try:  # to get the exam note
                members.loc[p, 'Exam_Pkt'] = sum(exam[c].ent.loc[p, pd.IndexSlice[:, 'Pkt']])
    # 1. Determine reached Pkt by exam
    # 2. Determine reached Pkt including bonus
                members.loc[p, 'Ges_Pkt'] = np.nansum([members.loc[p, 'Exam_Pkt'], members.loc[p, 'Bonus_Pkt']])
    # 3. Evaluate course note of participant
                if members.loc[p, 'Ges_Pkt'] == np.nan: 
                    members.loc[p, 'Note'] = scheme[0][0]
                else:
                    n_sel = members.loc[p, 'Ges_Pkt'] / max_pts * 100 >= scheme
                    # print(n_sel)
                    members.loc[p, 'Note'] = n_sel.index[n_sel][-1]
                    # members.loc[p, 'Note'] = scheme[0][n_sel.index[n_sel][-1]]
            except TypeError:
                print('### Exam and note evaluation failed for Member', p,
                      members.loc[p, 'Name'], 'in test', exam[c].name, '###')
    # 4. Create pd.DataFrame containing all entries of all cohorts together
        # overwrite all_entries (usually nan-rows) by values of exam[c].ent
        all_entries.update(exam[c].ent)
    n_all = len(members['Note'])
    n_part = len(members['Note'].dropna())
    n_fail = len(members[members['Note']=='5,0'])
    print(n_part, 'of', n_all, 'registered members participated in the exam:', round(n_part/n_all*100, 2), '%')
    print('of which', n_fail, 'members have failed the exam:', round(n_fail/n_part*100, 2), '%')
    print('Accordingly', n_part-n_fail, 'members have passed the exam:', round((n_part-n_fail)/n_part*100, 2), '%')
    print(members['Note'].value_counts().sort_index())
    return members, all_entries
