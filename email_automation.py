# -*- coding: utf-8 -*-
"""
Created on Wed Jun 21 11:00:40 2017

@author: dh1023
"""

import win32com.client as win32
import pandas as pd

# create a dataframe of all the crews and emails that need to be sent
#emails = pd.DataFrame([{'crew': 'Ops A', 'email': 'Thomas.Kelley@unh.edu'}, 
#                       {'crew': 'Ops A', 'email': 'Tad.Thomas@unh.edu'},
#                       {'crew': 'Ops B', 'email': 'Stan.Dodier@unh.edu'},
#                       {'crew': 'Ops B', 'email': 'Carl.Whitten@unh.edu'},
#                       {'crew': 'Ops B', 'email': 'Jeffrey.Prince@unh.edu'},
#                       {'crew': 'Ops C', 'email': 'Scott.Lindquist@unh.edu'},
#                       {'crew': 'Ops D', 'email': 'Jeffrey.McGrath@unh.edu'},
#                       {'crew': 'Ops D', 'email': 'Thomas.Smith@unh.edu'}])
#
#t_emails = pd.DataFrame([{'crew': 'Ops A', 'email': 'denham.hall@unh.edu'}, 
#                         {'crew': 'Ops B', 'email': 'dghall@gcaservices.com'},
#                         {'crew': 'Ops C', 'email': 'denhamghall@gmail.com'},
#                         {'crew': 'Ops D', 'email': 'denhamghall@gmail.com'}])

# create a loop to email all the appropriate managers
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'denham.hall@unh.edu; dghall@gcaservices.com'
mail.Subject = 'Zone Manager Report'
mail.Body = 'Heres the attached report'
excl_file = r'C:\Users\dh1023\Desktop\Python\QA QC Reports\OPS A - QC and QA work orders.xlsx'
mail.Attachments.Add(excl_file)
mail.Send()