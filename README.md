# QA_QC_reports

These Files are the process of automating the QA QC reports for facilities.

"random_wo.py" pulls a file from Facilities fileshare (sent there by a scheduled WebI report) 
and makes a list of random WO's that should be QA'd and QC'd. It then writes that data 
to an excel file in this folder.

"send_email.py" takes the files and emails them to the appropriate managers/supervisors.

"Batch Python Email.bat" is the file that Windows Task Scheduler automatically runs every 
Monday morning so that the data is crunched and the emails are sent.
