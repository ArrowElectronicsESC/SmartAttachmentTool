import os
import sys
import subprocess

if len(sys.argv) == 4:
    CsvPartSuggestion = str(sys.argv[1])
    FsrFaeName = str(sys.argv[2])
    Icc3MainPart = str(sys.argv[3])

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
    subprocess.call(["powershell.exe",  "-File", "sendemail.ps1",CsvPartSuggestion, FaeFsrEmail, FsrFaeName, Icc3MainPart])
else:
    print("Error - Bad parameters")
    print('Eg: py GetSendEmail.py CsvPartSuggestion FsrFaeName Icc3MainPart')