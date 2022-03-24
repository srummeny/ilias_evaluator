#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Mar 24 09:46:10 2022

@author: rummeny
"""
import matplotlib.pyplot as plt
import pandas as pd

fig, (ax1, ax2) = plt.subplots(ncols=2)
bonus_mrg['Boni durch Zwischentests'] = bonus_mrg['Boni durch Zwischentests'].astype(int)
bonus_mrg['Boni durch Praktika'] = bonus_mrg['Boni durch Praktika'].astype(int)
ty = 90
rot = 45

bonus_mrg.boxplot(column='Exam_Pkt', by='Boni durch Zwischentests', grid=False, ax=ax1)
vc1 = bonus_mrg['Boni durch Zwischentests'].value_counts()
plt.text(-4.2, ty, 'n='+str(vc1[0]), rotation=rot)# x=-4.2, y=0.83
plt.text(-3.4, ty, 'n='+str(vc1[1]), rotation=rot)# x=-4.2, y=0.83
plt.text(-2.6, ty, 'n='+str(vc1[2]), rotation=rot)# x=-4.2, y=0.83
plt.text(-1.8, ty, 'n='+str(vc1[3]), rotation=rot)# x=-4.2, y=0.83
plt.text(-1.0, ty, 'n='+str(vc1[4]), rotation=rot)# x=-4.2, y=0.83
ax1.set_title('')
ax1.set_ylabel('Klausurpunkte [%]')
ax1.set_ylim([0, 100])

bonus_mrg.boxplot(column='Exam_Pkt', by='Boni durch Praktika', grid=False, ax=ax2)
vc2 = bonus_mrg['Boni durch Praktika'].value_counts()
plt.text(0.7, ty, 'n='+str(vc2[0]), rotation=rot)# x=-4.2, y=0.83
plt.text(1.7, ty, 'n='+str(vc2[1]), rotation=rot)# x=-4.2, y=0.83
plt.text(2.7, ty, 'n='+str(vc2[2]), rotation=rot)# x=-4.2, y=0.83
plt.text(3.7, ty, 'n='+str(vc2[3]), rotation=rot)# x=-4.2, y=0.83
ax2.set_title('')
ax2.set_ylim([0, 100])
fig.suptitle('Untersuchung WS 2021/22')