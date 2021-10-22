#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Oct 13 16:17:45 2021

@author: rummeny
"""
import pandas as pd


f = pd.read_excel('ETG_2020w_exp_results_review.xlsx', sheet_name='Sheet1', index_col=0)
f = f.transpose()
s = pd.read_excel('ETG_WS2021_Exam/exp_Ergebnisse_pub_final.xlsx', skiprows=2, sheet_name='Sheet1', index_col=0)
s = s.transpose()
# do comparison with these
s_in_f = [any(s.index[i]==f.index) for i in range(len(s))]
mats_s = []
for m in s.index[s_in_f]:
    if float(f['Note'][m].replace(',','.')) < float(s['Note'][m].replace(',','.')):
        print(m, 'was better at first exam!', float(f['Note'][m].replace(',','.')), '<--', float(s['Note'][m].replace(',','.')))
        mats_s.append(m)
# check if someone passed of these
f_not_in_s = [not any(f.index[i]==s.index) for i in range(len(f))]
mats_f_be = []
mats_f_nbe = []
for m in f.index[f_not_in_s]:
    if float(f['Note'][m].replace(',','.')) < 5:
        print(m, 'did only first exam and passed it!', float(f['Note'][m].replace(',','.')))
        mats_f_be.append(m)
    else:
        # print(m, 'failed at first exam!')
        mats_f_nbe.append(m)
better_res = f.copy()
better_res['Status'] = ''
better_res.loc[:, 'Status'] = 'Klausur am Nachholtermin vom 22./23.04.2021 besser bestanden oder gleich als am 12.03.2021. Daher gilt weiter die Note vom Nachholtermin und nicht die hier genannte'
better_res.loc[mats_f_be, 'Status'] = 'Klausur am 12.03.2021 bestanden und danach nicht mehr teilgenommen'
better_res.loc[mats_f_nbe, 'Status'] = 'Klausur am 12.03.2021 nicht bestanden und danach nicht mehr teilgenommen'
better_res.loc[mats_s, 'Status'] = 'Klausur am 12.03.2021 besser bestanden als am Nachholtermin vom 22./23.04.2021'
better_res = better_res.sort_index()
better_res.columns = list(s.columns[:-3].values) + ['A41_Student', 'A41_Musterloesung', 'A41_Punkte'] + list(better_res.columns[-4:].values)
note_list = better_res['Note'][mats_f_be + mats_s].sort_index()
better_res = better_res.transpose()
better_res.to_excel('2020w_ETG_results_review.xlsx')
note_list.to_excel('2020w_ETG_results_review_psso.xlsx')
