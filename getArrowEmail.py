import os
import sys
import subprocess

arrowName = str(sys.argv[1])

ArrowSplitName = arrowName.split(', ')
ArrowLastName = ArrowSplitName[0].split(' ')
ArrowName = ArrowSplitName[1].split(' ')
if len(ArrowName) > 1:
    ArrowEmail = str(ArrowName[0]) + str(' ')
else:
    ArrowEmail = str(ArrowName[0]) + str(' ')

if len(ArrowLastName) > 1:
    ArrowEmail = str(ArrowEmail) + str(ArrowLastName[1])
else:
    ArrowEmail = str(ArrowEmail) + str(ArrowLastName[0])

subprocess.run(["powershell.exe", "-File", "getArrowEmail.ps1",ArrowEmail])