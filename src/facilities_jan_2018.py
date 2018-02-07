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
        
    

"""
# is_scored = ~np.isnan(all_data)

# avg_qns = np.nanmean(all_data, axis=1).transpose()
# avg_usr = np.nanmean(all_data, axis=2)
# 
# 
# avg_qnsilities = pd.DataFrame(np.nanmean(all_data,axis=0).transpose(), columns=questions)
# avg_questions = pd.DataFrame(np.nanmean(all_data,axis=1).transpose(), columns=facilities)
# 
# 
# x_users.extend(['Average all x_users', 'StDev all x_users', 'N'])
# questions.extend(['Average all questions', 'StDev all questions', 'N'])
# facilities.extend(['Average all facilities', 'StDev all facilities', 'N'])
# 
# h_mean_qns = pd.DataFrame(avg_qnsilities.mean(axis=1))
# h_std_qns = pd.DataFrame(avg_qnsilities.std(axis=1))
# h_count_qns = pd.DataFrame(pd.notnull(avg_qnsilities).sum(axis=1))
# v_mean_qns = avg_qnsilities.mean(axis=0)
# v_std_qns = avg_qnsilities.std(axis=0)
# v_count_qns = pd.notnull(avg_qnsilities).sum(axis=0)
# 
# avg_qnsilities = pd.concat([avg_qnsilities, h_mean_qns, h_std_qns, h_count_qns], axis=1)
# avg_qnsilities.columns = questions
# print(avg_qnsilities)
# avg_qnsilities = pd.concat([avg_qnsilities, v_mean_qns, v_std_qns, v_count_qns], axis=0)
# print(avg_qnsilities)
# # avg_qnsilities = pd.concat([avg_qnsilities, 
# #                             pd.DataFrame([avg_qnsilities.mean(axis=0)]), 
# #                             pd.DataFrame([avg_qnsilities.std(axis=0)]),
# #                             pd.notnull(avg_qnsilities).sum(axis=0)
# #                             ], 
# #                             axis=0)
# # print(avg_qnsilities)
# indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# avg_qnsilities = avg_qnsilities.set_index(indx)
# avg_qnsilities.columns = questions
# 
# print(avg_qnsilities)


# print(avg_questions.mean(axis=1))
# print(pd.notnull(avg_qnsilities).sum(axis=1)) #.sum(axis=1))


# avg_questions = pd.concat([avg_questions, 
#                            avg_questions.mean(axis=1), 
#                            avg_questions.std(axis=1),
#                            pd.notnull(avg_questions).sum(axis=1)], 
#                            axis=1)
# avg_questions = pd.concat([avg_questions, 
#                            pd.DataFrame([avg_questions.mean(axis=0)]), 
#                            pd.DataFrame([avg_questions.std(axis=0)]),
#                            pd.notnull(avg_questions).sum(axis=0)], 
#                            axis=0)
# indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# avg_questions.index = indx # = avg_questions.set_index(indx)
# avg_questions.columns = facilities
# 
# print(avg_questions)



# indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# avg_qnsilities = pd.DataFrame(np.nanmean(all_data,axis=0).transpose(), index=indx, columns=questions)
# avg_qnsilities = pd.concat([avg_qnsilities, avg_qnsilities.mean(axis=1), avg_qnsilities.std(axis=1)], axis=1)
# avg_qnsilities.columns = np.hstack((questions, ['Average all questions', 'StDev all questions']))
# cols = np.hstack((questions, ['Average all questions', 'StDev all questions']))
# x_users.extend(['Average all x_users', 'StDev all x_users'])
# indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# avg_qnsilities = pd.concat([avg_qnsilities, 
#                             pd.DataFrame([avg_qnsilities.mean(axis=0)], columns=cols), 
#                             pd.DataFrame([avg_qnsilities.std(axis=0)], columns=cols)], 
#                             axis=0)
# avg_qnsilities = avg_qnsilities.set_index(indx)
# 
# print(avg_qnsilities)

# err_qnsilities_users = pd.DataFrame(np.nanstd(avg_qnsilities, axis=1), index=['StErr']) 
# avg_qnsilities = pd.concat([avg_qnsilities, avg_qnsilities_users, err_qnsilities_users], axis=0)
# avg_qnsilities_questions = pd.DataFrame(np.nanmean(avg_qnsilities, axis=1), index=['Average all questions'])
# err_qnsilities_questions = pd.DataFrame(np.nanstd(avg_qnsilities, axis=1), index=['StErr'])
# indx = MultiIndex.from_arrays([['StDev']*len(x_users), x_users])
# std_qnsilities = pd.DataFrame(np.nanstd(all_data,axis=0).transpose(), index=indx, columns=questions)
# indx = MultiIndex.from_arrays([['N']*len(x_users), x_users])
# n_qnsilities = pd.DataFrame(np.sum(is_scored, axis=0).transpose(), index=indx, columns=questions) 
# 
# 
# indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# avg_questions = pd.DataFrame(np.nanmean(all_data, axis=1).transpose(), index=indx, columns=facilities)
# indx = MultiIndex.from_arrays([['StDev']*len(x_users), x_users])
# std_questions = pd.DataFrame(np.nanstd(all_data, axis=1).transpose(), index=indx, columns=facilities)
# indx = MultiIndex.from_arrays([['N']*len(x_users), x_users])
# n_questions = pd.DataFrame(np.sum(is_scored, axis=1).transpose(), index=indx, columns=facilities)
# 
# indx = MultiIndex.from_arrays([['Mean']*len(facilities), facilities])
# avg_users = pd.DataFrame(np.nanmean(all_data, axis=2), index=indx, columns=questions)
# indx = MultiIndex.from_arrays([['StDev']*len(facilities), facilities])
# std_users = pd.DataFrame(np.nanstd(all_data, axis=2), index=indx, columns=questions)
# indx = MultiIndex.from_arrays([['N']*len(facilities), facilities])
# n_users = pd.DataFrame(np.sum(is_scored, axis=2), index=indx, columns=questions)
# 
# all_qnsilities = pd.concat([avg_qnsilities, std_qnsilities, n_qnsilities], axis=0)
# all_questions = pd.concat([avg_questions, std_questions, n_questions], axis=0)
# all_users = pd.concat([avg_users, std_users, n_users], axis=0)
# 
# file_path = QFileDialog.getSaveFileName(
#     None,
#     "Save results to Excel",
#     "Results.xlsx",
#     "Excel X files (*.xlsx);;All files (*.*)")[0]
# if not file_path is None:
#     writer = pd.ExcelWriter(file_path)
#     all_qnsilities.to_excel(writer, 'Facilities')
#     all_questions.to_excel(writer, 'Questions')
#     all_users.to_excel(writer, 'Users')
#     writer.save()
#     writer.close()
# 

# if __name__ == '__main__':
#     pass
"""

# '''
# Created on 29 Jan 2018
# 
# @author: schilsm
# '''
# 
# import sys
# from PyQt5.QtCore import *
# from PyQt5.QtWidgets import *
# import pandas as pd, numpy as np, copy as cp
# from pandas.indexes.multi import MultiIndex
# #from odo.convert import _ignore_index
# 
# app = QApplication(sys.argv)
# 
# file_path = QFileDialog.getOpenFileNames(
#     None,
#     "Open Data File",
#     "",
#     "Excel X files (*.xlsx);;All files (*.*)")[0]
# i = 0
# all_data = None
# users, questions, facilities, comments = [], [], [], {}
# for f in file_path:
#     raw_data = pd.read_excel(f)
# #    print(raw_data.iloc[54:55, 2:3])
# #    name = raw_data.iloc[1,3]
#     name = QFileInfo(f).baseName()
#     #comments[name] = raw_data.iloc[54:55, 2:3].values
#     users.append(name)
#     data = raw_data.iloc[0:7,14:48]
#     questions = raw_data.iloc[13,4:10].values.tolist()
#     facilities = raw_data.iloc[15:45, 2].drop(raw_data.index[43]).replace('[*]+','', regex=True).values.tolist()
#     data = raw_data.iloc[15:45,4:10].drop(raw_data.index[43])
#     data.index = facilities
#     data.columns = questions
#     data = data.apply(pd.to_numeric, errors='coerce')
#     if i == 0:
#         all_data = data.as_matrix()
#     else:
#         all_data = np.dstack((all_data, data.as_matrix()))
#     i+= 1
#     
# print(comments)
# 
# if not all_data is None:
#     x_users = cp.deepcopy(users)
#     x_users.extend(['Average all users', 'StDev all users', 'N'])
#     x_questions = cp.deepcopy(questions) 
#     x_questions.extend(['Average all questions', 'StDev all questions', 'N'])
#     x_facilities = cp.deepcopy(facilities)
#     x_facilities.extend(['Average all facilities', 'StDev all facilities', 'N'])
#     
#     # Averages over all facilities
#     avg_fac = np.nanmean(all_data, axis=0)
#     m, s, n = np.nanmean(avg_fac, axis=0), np.nanstd(avg_fac, axis=0), np.sum(~np.isnan(avg_fac), axis=0)
#     avg_fac = np.vstack((avg_fac, m, s, n)).T
#     m, s, n = np.nanmean(avg_fac, axis=0), np.nanstd(avg_fac, axis=0), np.sum(~np.isnan(avg_fac), axis=0)
#     avg_fac = np.vstack((avg_fac, m, s, n))
#     
#     std_fac = np.nanstd(all_data, axis=0)
#     f = np.zeros((3, std_fac.shape[1]), dtype=float)/0.0
#     std_fac = np.vstack((std_fac, f)).T
#     
#     num_fac = np.sum(~np.isnan(all_data), axis=0) #np.nanstd(all_data, axis=0)
#     f = np.zeros((3, num_fac.shape[1]), dtype=float)/0.0
#     num_fac = np.vstack((num_fac, f)).T
#     
#     avg_fac = pd.DataFrame(avg_fac, index=MultiIndex.from_arrays([['Mean']*len(x_users), x_users]), columns=x_questions)
#     std_fac = pd.DataFrame(std_fac, index=MultiIndex.from_arrays([['StDev']*len(users), users]), columns=x_questions)
#     num_fac = pd.DataFrame(num_fac, index=MultiIndex.from_arrays([['N']*len(users), users]), columns=x_questions)
#     all_fac = avg_fac.append(std_fac)
#     all_fac = all_fac.append(num_fac)
#     
#     # Averages over all x_questions
#     avg_qns = np.nanmean(all_data, axis=1)
#     m, s, n = np.nanmean(avg_qns, axis=0), np.nanstd(avg_qns, axis=0), np.sum(~np.isnan(avg_qns), axis=0)
#     avg_qns = np.vstack((avg_qns, m, s, n)).T
#     m, s, n = np.nanmean(avg_qns, axis=0), np.nanstd(avg_qns, axis=0), np.sum(~np.isnan(avg_qns), axis=0)
#     avg_qns = np.vstack((avg_qns, m, s, n))
#     
#     std_qns = np.nanstd(all_data, axis=1)
#     f = np.zeros((3, std_qns.shape[1]), dtype=float)/0.0
#     std_qns = np.vstack((std_qns, f)).T
#     
#     num_qns = np.sum(~np.isnan(all_data), axis=1) #np.nanstd(all_data, axis=1)
#     f = np.zeros((3, num_qns.shape[1]), dtype=float)/0.0
#     num_qns = np.vstack((num_qns, f)).T
#     
#     avg_qns = pd.DataFrame(avg_qns, index=MultiIndex.from_arrays([['Mean']*len(x_users), x_users]), columns=x_facilities)
#     std_qns = pd.DataFrame(std_qns, index=MultiIndex.from_arrays([['StDev']*len(users), users]), columns=x_facilities)
#     num_qns = pd.DataFrame(num_qns, index=MultiIndex.from_arrays([['N']*len(users), users]), columns=x_facilities)
#     all_qns = avg_qns.append(std_qns)
#     all_qns = all_qns.append(num_qns)
#     
#     # Averages over all x_users
#     avg_usr = np.nanmean(all_data, axis=2).T
#     m, s, n = np.nanmean(avg_usr, axis=0), np.nanstd(avg_usr, axis=0), np.sum(~np.isnan(avg_usr), axis=0)
#     avg_usr = np.vstack((avg_usr, m, s, n)).T
#     m, s, n = np.nanmean(avg_usr, axis=0), np.nanstd(avg_usr, axis=0), np.sum(~np.isnan(avg_usr), axis=0)
#     avg_usr = np.vstack((avg_usr, m, s, n))
#     
#     std_usr = np.nanstd(all_data, axis=2).T
#     f = np.zeros((3, std_usr.shape[1]), dtype=float)/0.0
#     std_usr = np.vstack((std_usr, f)).T
#     
#     num_usr = np.sum(~np.isnan(all_data), axis=2).T #np.nanstd(all_data, axis=2).T
#     f = np.zeros((3, num_usr.shape[1]), dtype=float)/0.0
#     num_usr = np.vstack((num_usr, f)).T
#     
#     avg_usr = pd.DataFrame(avg_usr, index=MultiIndex.from_arrays([['Mean']*len(x_facilities), x_facilities]), columns=x_questions)
#     std_usr = pd.DataFrame(std_usr, index=MultiIndex.from_arrays([['StDev']*len(facilities), facilities]), columns=x_questions)
#     num_usr = pd.DataFrame(num_usr, index=MultiIndex.from_arrays([['N']*len(facilities), facilities]), columns=x_questions)
#     all_usr = avg_usr.append(std_usr)
#     all_usr = all_usr.append(num_usr)
#     
#     file_path = QFileDialog.getSaveFileName(
#         None,
#         "Save results to Excel",
#         "Results.xlsx",
#         "Excel X files (*.xlsx);;All files (*.*)")
#     if file_path[0] != "":
#         writer = pd.ExcelWriter(file_path[0])
#         all_fac.to_excel(writer, 'Facilities')
#         all_qns.to_excel(writer, 'Questions')
#         all_usr.to_excel(writer, 'Users')
#         writer.save()
#         writer.close()    
# 
# """
# # is_scored = ~np.isnan(all_data)
# 
# # avg_qns = np.nanmean(all_data, axis=1).transpose()
# # avg_usr = np.nanmean(all_data, axis=2)
# # 
# # 
# # avg_qnsilities = pd.DataFrame(np.nanmean(all_data,axis=0).transpose(), columns=questions)
# # avg_questions = pd.DataFrame(np.nanmean(all_data,axis=1).transpose(), columns=facilities)
# # 
# # 
# # x_users.extend(['Average all x_users', 'StDev all x_users', 'N'])
# # questions.extend(['Average all questions', 'StDev all questions', 'N'])
# # facilities.extend(['Average all facilities', 'StDev all facilities', 'N'])
# # 
# # h_mean_qns = pd.DataFrame(avg_qnsilities.mean(axis=1))
# # h_std_qns = pd.DataFrame(avg_qnsilities.std(axis=1))
# # h_count_qns = pd.DataFrame(pd.notnull(avg_qnsilities).sum(axis=1))
# # v_mean_qns = avg_qnsilities.mean(axis=0)
# # v_std_qns = avg_qnsilities.std(axis=0)
# # v_count_qns = pd.notnull(avg_qnsilities).sum(axis=0)
# # 
# # avg_qnsilities = pd.concat([avg_qnsilities, h_mean_qns, h_std_qns, h_count_qns], axis=1)
# # avg_qnsilities.columns = questions
# # print(avg_qnsilities)
# # avg_qnsilities = pd.concat([avg_qnsilities, v_mean_qns, v_std_qns, v_count_qns], axis=0)
# # print(avg_qnsilities)
# # # avg_qnsilities = pd.concat([avg_qnsilities, 
# # #                             pd.DataFrame([avg_qnsilities.mean(axis=0)]), 
# # #                             pd.DataFrame([avg_qnsilities.std(axis=0)]),
# # #                             pd.notnull(avg_qnsilities).sum(axis=0)
# # #                             ], 
# # #                             axis=0)
# # # print(avg_qnsilities)
# # indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# # avg_qnsilities = avg_qnsilities.set_index(indx)
# # avg_qnsilities.columns = questions
# # 
# # print(avg_qnsilities)
# 
# 
# # print(avg_questions.mean(axis=1))
# # print(pd.notnull(avg_qnsilities).sum(axis=1)) #.sum(axis=1))
# 
# 
# # avg_questions = pd.concat([avg_questions, 
# #                            avg_questions.mean(axis=1), 
# #                            avg_questions.std(axis=1),
# #                            pd.notnull(avg_questions).sum(axis=1)], 
# #                            axis=1)
# # avg_questions = pd.concat([avg_questions, 
# #                            pd.DataFrame([avg_questions.mean(axis=0)]), 
# #                            pd.DataFrame([avg_questions.std(axis=0)]),
# #                            pd.notnull(avg_questions).sum(axis=0)], 
# #                            axis=0)
# # indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# # avg_questions.index = indx # = avg_questions.set_index(indx)
# # avg_questions.columns = facilities
# # 
# # print(avg_questions)
# 
# 
# 
# # indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# # avg_qnsilities = pd.DataFrame(np.nanmean(all_data,axis=0).transpose(), index=indx, columns=questions)
# # avg_qnsilities = pd.concat([avg_qnsilities, avg_qnsilities.mean(axis=1), avg_qnsilities.std(axis=1)], axis=1)
# # avg_qnsilities.columns = np.hstack((questions, ['Average all questions', 'StDev all questions']))
# # cols = np.hstack((questions, ['Average all questions', 'StDev all questions']))
# # x_users.extend(['Average all x_users', 'StDev all x_users'])
# # indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# # avg_qnsilities = pd.concat([avg_qnsilities, 
# #                             pd.DataFrame([avg_qnsilities.mean(axis=0)], columns=cols), 
# #                             pd.DataFrame([avg_qnsilities.std(axis=0)], columns=cols)], 
# #                             axis=0)
# # avg_qnsilities = avg_qnsilities.set_index(indx)
# # 
# # print(avg_qnsilities)
# 
# # err_qnsilities_users = pd.DataFrame(np.nanstd(avg_qnsilities, axis=1), index=['StErr']) 
# # avg_qnsilities = pd.concat([avg_qnsilities, avg_qnsilities_users, err_qnsilities_users], axis=0)
# # avg_qnsilities_questions = pd.DataFrame(np.nanmean(avg_qnsilities, axis=1), index=['Average all questions'])
# # err_qnsilities_questions = pd.DataFrame(np.nanstd(avg_qnsilities, axis=1), index=['StErr'])
# # indx = MultiIndex.from_arrays([['StDev']*len(x_users), x_users])
# # std_qnsilities = pd.DataFrame(np.nanstd(all_data,axis=0).transpose(), index=indx, columns=questions)
# # indx = MultiIndex.from_arrays([['N']*len(x_users), x_users])
# # n_qnsilities = pd.DataFrame(np.sum(is_scored, axis=0).transpose(), index=indx, columns=questions) 
# # 
# # 
# # indx = MultiIndex.from_arrays([['Mean']*len(x_users), x_users])
# # avg_questions = pd.DataFrame(np.nanmean(all_data, axis=1).transpose(), index=indx, columns=facilities)
# # indx = MultiIndex.from_arrays([['StDev']*len(x_users), x_users])
# # std_questions = pd.DataFrame(np.nanstd(all_data, axis=1).transpose(), index=indx, columns=facilities)
# # indx = MultiIndex.from_arrays([['N']*len(x_users), x_users])
# # n_questions = pd.DataFrame(np.sum(is_scored, axis=1).transpose(), index=indx, columns=facilities)
# # 
# # indx = MultiIndex.from_arrays([['Mean']*len(facilities), facilities])
# # avg_users = pd.DataFrame(np.nanmean(all_data, axis=2), index=indx, columns=questions)
# # indx = MultiIndex.from_arrays([['StDev']*len(facilities), facilities])
# # std_users = pd.DataFrame(np.nanstd(all_data, axis=2), index=indx, columns=questions)
# # indx = MultiIndex.from_arrays([['N']*len(facilities), facilities])
# # n_users = pd.DataFrame(np.sum(is_scored, axis=2), index=indx, columns=questions)
# # 
# # all_qnsilities = pd.concat([avg_qnsilities, std_qnsilities, n_qnsilities], axis=0)
# # all_questions = pd.concat([avg_questions, std_questions, n_questions], axis=0)
# # all_users = pd.concat([avg_users, std_users, n_users], axis=0)
# # 
# # file_path = QFileDialog.getSaveFileName(
# #     None,
# #     "Save results to Excel",
# #     "Results.xlsx",
# #     "Excel X files (*.xlsx);;All files (*.*)")[0]
# # if not file_path is None:
# #     writer = pd.ExcelWriter(file_path)
# #     all_qnsilities.to_excel(writer, 'Facilities')
# #     all_questions.to_excel(writer, 'Questions')
# #     all_users.to_excel(writer, 'Users')
# #     writer.save()
# #     writer.close()
# # 
# 
# # if __name__ == '__main__':
# #     pass
# """