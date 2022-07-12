import os
import sys
import subprocess
#import win32com.client as win32

#outlook = win32.Dispatch('outlook.application')

if len(sys.argv) == 5:
    CsvPartSuggestion = str(sys.argv[1])
    FsrFaeName = str(sys.argv[2])
    YearQx = str(sys.argv[3])
    Year20xx = str(sys.argv[4])

    ArrowSplitName = FsrFaeName.split(', ')
    ArrowLastName = ArrowSplitName[0].split(' ')
    FsrFaeName1 = ArrowSplitName[1].split(' ')
    if len(FsrFaeName1) > 1:
        ArrowEmail = str(FsrFaeName1[0]) + str(' ')
    else:
        ArrowEmail = str(FsrFaeName1[0]) + str(' ')

    if len(ArrowLastName) > 1:
        ArrowEmail = str(ArrowEmail) + str(ArrowLastName[1])
    else:
        ArrowEmail = str(ArrowEmail) + str(ArrowLastName[0])

    FaeFsrEmail = subprocess.run(["powershell.exe", "-File", "getArrowEmail.ps1",str(ArrowEmail)], shell=True ,capture_output=True,text=True)
    FaeFsrEmail = str(FaeFsrEmail.stdout)
    subprocess.call(["powershell.exe",  "-File", "sendemail.ps1",CsvPartSuggestion, FaeFsrEmail, FsrFaeName, YearQx, Year20xx])
    #mail = outlook.CreateItem(0)
    #mail.SentOnBehalfOfName = "dcdcoordinator@arrow.com"
    #mail.To = "ericmtzr89@gmail.com" #FaeFsrEmail
    #mail.Subject = 'Message subject'
    #mail.Body = 'Message body'
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    #mail.Send()
else:
    print("Error - Bad parameters")
    print('Eg: py GetSendEmail.py CsvPartSuggestion FsrFaeName Q3 2022')