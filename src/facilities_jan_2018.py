'''
Created on 29 Jan 2018

@author: schilsm
'''

import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
import pandas as pd, numpy as np, copy as cp
from pandas import MultiIndex #.indexes.multi import MultiIndex
#from odo.convert import _ignore_index

app = QApplication(sys.argv)

file_path = QFileDialog.getOpenFileNames(
    None,
    "Open Data File",
    "",
    "Excel X files (*.xlsx);;All files (*.*)")[0]
i = 0
all_data = None
users, questions, facilities, comments = [], [], [], {}
for f in file_path:
    name = QFileInfo(f).baseName()
    users.append(name)
    raw_data = pd.read_excel(f)
#    name = raw_data.iloc[1,3] # name, but not always there
#    print(raw_data.iloc[54:55, 2:3]) # comments
#    comments[name] = raw_data.iloc[54:55, 2:3].values
    data = raw_data.iloc[0:7,14:48]
    questions = raw_data.iloc[13,4:10].values.tolist()
    facilities = raw_data.iloc[15:45, 2].drop(raw_data.index[43]).replace('[*]+','', regex=True).values.tolist()
    data = raw_data.iloc[15:45,4:10].drop(raw_data.index[43])
    data.index = facilities
    data.columns = questions
    data = data.apply(pd.to_numeric, errors='coerce')
    if i == 0:
        all_data = data.as_matrix()
    else:
        all_data = np.dstack((all_data, data.as_matrix()))
    i+= 1

    

if not all_data is None:
    x_users = cp.deepcopy(users)
    x_users.extend(['Average all users', 'StDev all users', 'N'])
    x_questions = cp.deepcopy(questions) 
    x_questions.extend(['Average all questions', 'StDev all questions', 'N'])
    x_facilities = cp.deepcopy(facilities)
    x_facilities.extend(['Average all facilities', 'StDev all facilities', 'N'])
     
    # Averages over all facilities
    avg_fac = np.nanmean(all_data, axis=0)
    m, s, n = np.nanmean(avg_fac, axis=0), np.nanstd(avg_fac, axis=0), np.sum(~np.isnan(avg_fac), axis=0)
    avg_fac = np.vstack((avg_fac, m, s, n)).T
    m, s, n = np.nanmean(avg_fac, axis=0), np.nanstd(avg_fac, axis=0), np.sum(~np.isnan(avg_fac), axis=0)
    avg_fac = np.vstack((avg_fac, m, s, n))
     
    std_fac = np.nanstd(all_data, axis=0)
    f = np.zeros((3, std_fac.shape[1]), dtype=float)/0.0
    std_fac = np.vstack((std_fac, f)).T
     
    num_fac = np.sum(~np.isnan(all_data), axis=0) #np.nanstd(all_data, axis=0)
    f = np.zeros((3, num_fac.shape[1]), dtype=float)/0.0
    num_fac = np.vstack((num_fac, f)).T
     
    avg_fac = pd.DataFrame(avg_fac, index=MultiIndex.from_arrays([['Mean']*len(x_users), x_users]), columns=x_questions)
    std_fac = pd.DataFrame(std_fac, index=MultiIndex.from_arrays([['StDev']*len(users), users]), columns=x_questions)
    num_fac = pd.DataFrame(num_fac, index=MultiIndex.from_arrays([['N']*len(users), users]), columns=x_questions)
    all_fac = avg_fac.append(std_fac)
    all_fac = all_fac.append(num_fac)
     
    # Averages over all x_questions
    avg_qns = np.nanmean(all_data, axis=1)
    m, s, n = np.nanmean(avg_qns, axis=0), np.nanstd(avg_qns, axis=0), np.sum(~np.isnan(avg_qns), axis=0)
    avg_qns = np.vstack((avg_qns, m, s, n)).T
    m, s, n = np.nanmean(avg_qns, axis=0), np.nanstd(avg_qns, axis=0), np.sum(~np.isnan(avg_qns), axis=0)
    avg_qns = np.vstack((avg_qns, m, s, n))
     
    std_qns = np.nanstd(all_data, axis=1)
    f = np.zeros((3, std_qns.shape[1]), dtype=float)/0.0
    std_qns = np.vstack((std_qns, f)).T
     
    num_qns = np.sum(~np.isnan(all_data), axis=1) #np.nanstd(all_data, axis=1)
    f = np.zeros((3, num_qns.shape[1]), dtype=float)/0.0
    num_qns = np.vstack((num_qns, f)).T
     
    avg_qns = pd.DataFrame(avg_qns, index=MultiIndex.from_arrays([['Mean']*len(x_users), x_users]), columns=x_facilities)
    std_qns = pd.DataFrame(std_qns, index=MultiIndex.from_arrays([['StDev']*len(users), users]), columns=x_facilities)
    num_qns = pd.DataFrame(num_qns, index=MultiIndex.from_arrays([['N']*len(users), users]), columns=x_facilities)
    all_qns = avg_qns.append(std_qns)
    all_qns = all_qns.append(num_qns)
     
    # Averages over all x_users
    avg_usr = np.nanmean(all_data, axis=2).T
    m, s, n = np.nanmean(avg_usr, axis=0), np.nanstd(avg_usr, axis=0), np.sum(~np.isnan(avg_usr), axis=0)
    avg_usr = np.vstack((avg_usr, m, s, n)).T
    m, s, n = np.nanmean(avg_usr, axis=0), np.nanstd(avg_usr, axis=0), np.sum(~np.isnan(avg_usr), axis=0)
    avg_usr = np.vstack((avg_usr, m, s, n))
     
    std_usr = np.nanstd(all_data, axis=2).T
    f = np.zeros((3, std_usr.shape[1]), dtype=float)/0.0
    std_usr = np.vstack((std_usr, f)).T
     
    num_usr = np.sum(~np.isnan(all_data), axis=2).T #np.nanstd(all_data, axis=2).T
    f = np.zeros((3, num_usr.shape[1]), dtype=float)/0.0
    num_usr = np.vstack((num_usr, f)).T
     
    avg_usr = pd.DataFrame(avg_usr, index=MultiIndex.from_arrays([['Mean']*len(x_facilities), x_facilities]), columns=x_questions)
    std_usr = pd.DataFrame(std_usr, index=MultiIndex.from_arrays([['StDev']*len(facilities), facilities]), columns=x_questions)
    num_usr = pd.DataFrame(num_usr, index=MultiIndex.from_arrays([['N']*len(facilities), facilities]), columns=x_questions)
    all_usr = avg_usr.append(std_usr)
    all_usr = all_usr.append(num_usr)
  
      
    file_path = QFileDialog.getSaveFileName(
        None,
        "Save results to Excel",
        "Results.xlsx",
        "Excel X files (*.xlsx);;All files (*.*)")
    if file_path[0] != "":
        writer = pd.ExcelWriter(file_path[0])
        all_fac.to_excel(writer, 'Facilities')
        all_qns.to_excel(writer, 'Questions')
        all_usr.to_excel(writer, 'Users')
        writer.save()
        writer.close()
        
