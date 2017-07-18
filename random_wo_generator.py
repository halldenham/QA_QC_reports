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
wo_df = pd.read_excel(wo_excel, sheetname = 'Python')

# create a df of crews and how many random QC/QA work orders are needed
ops = ['OPS A', 'OPS B', 'OPS C', 'OPS D', 'LOCK']
qc = [4, 6, 1, 8, 2]
qa = [2, 4, 1, 5, 1]

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
    
    # add columns at front of df to specify if QC or QA
    qa_wo.insert(0, 'QC/QA', 'QA')
    qc_wo.insert(0, 'QC/QA', 'QC')
    
    # append dataframes together
    qc_wo = qc_wo.append(qa_wo, ignore_index=True)
    
    
    # export to an excel file
    file_path = r'C:\Users\dh1023\Desktop\Python\qa_qc_reports\{} - QC and QA work orders.xlsx'.format(ops[i])
    
    # get the excel writer
    writer = pd.ExcelWriter(file_path)
    qc_wo.to_excel(writer, index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # column format - wrap the text and add border
    format_wrap = workbook.add_format()
    format_wrap.set_text_wrap()
    format_wrap.set_border()
    format_wrap.set_align('top')
    format_wrap.set_align('left')
    # Set the column width and format
    worksheet.set_column('A:A', 6, format_wrap)
    worksheet.set_column('B:B', 11, format_wrap)
    worksheet.set_column('C:C', 6, format_wrap)
    worksheet.set_column('D:D', 4, format_wrap)
    worksheet.set_column('E:E', 8, format_wrap)
    worksheet.set_column('F:F', 27, format_wrap)
    worksheet.set_column('G:G', 40, format_wrap)
    worksheet.set_column('H:H', 7, format_wrap)
    worksheet.set_column('I:I', 22, format_wrap)
    worksheet.set_column('J:J', 5, format_wrap)
    worksheet.set_column('K:K', 8, format_wrap)
    worksheet.set_column('L:L', 10, format_wrap)
    worksheet.set_column('M:M', 18, format_wrap)
    # write the file
    writer.save()
        
    print('Completed', ops[i])
    i +=1

print('\nProgram Completed!')
