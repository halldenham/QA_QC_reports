# -*- coding: utf-8 -*-
"""
Created on Wed Jun  7 12:51:11 2017

@author: dh1023

This script selects a random number of QA and QC work orders from a given
data source, and creates an excel file for each crew that is found within
the data frame of QC and QA numbers created by this file
"""

import pandas as pd


# Location of file with WO data. Include PM and SR that were completed within
# the previous week.
wo_excel = r'Z:\All Crews - QA QC.xlsx'

# Create dataframe from previous WO data
wo_df = pd.read_excel(wo_excel)

# create a df of crews and how many random QC/QA work orders are needed
ops = ['OPS A', 'OPS B', 'OPS C', 'OPS D']
qc = [11, 20, 2, 24]
qa = [8, 16, 2, 19]
ops_qc_qa_zip = list(zip(ops, qc, qa))
ops_qc_qa = pd.DataFrame(data = ops_qc_qa_zip, columns=['Crew', 'QC', 'QA'])
print('\n', ops_qc_qa,'\n')


i = 0 # this is used to index crews and qa/qc values or data
while i < len(ops_qc_qa):
    
    # select random QC data with current crew (i value)
    qc_wo = wo_df.loc[wo_df['Crew']==ops[i]].sample(qc[i])
    
    # make sure there's enough QA wo's to sample
    if len(wo_df.loc[(wo_df['Crew']==ops[i]) & (wo_df['KPI WO Due Status']
    =='Completed Late')]) < qa[i]:
        
        # if there aren't enough QA work orders, select all late WO's
        print('Not enough late work orders to get full QA for', ops[i])
        qa_wo = wo_df.loc[(wo_df['Crew']==ops[i]) & 
                          (wo_df['KPI WO Due Status']=='Completed Late')
                          ].copy()
    else:
        # select random QA data with current crew
        qa_wo = wo_df.loc[(wo_df['Crew']==ops[i]) & 
                          (wo_df['KPI WO Due Status']=='Completed Late')
                          ].sample(qa[i])
    
    # add columns to specify if QC or QA
    qa_wo['QC or QA'] = 'QA'
    qc_wo['QC or QA'] = 'QC'
    
    # append dataframes together
    qc_wo = qc_wo.append(qa_wo, ignore_index=True)
    
    # export to an excel file
    qc_wo.to_excel(ops[i] + r' - QC and QA work orders.xlsx', 
                   index=False)
    
    print('Completed', ops[i])
    i +=1

print('\nProgram Completed!')