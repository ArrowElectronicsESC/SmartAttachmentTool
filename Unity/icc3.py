import pandas as pd
import datetime
import sys
import os

if len(sys.argv) == 3:
    IccMainPart = str(sys.argv[1])
    IccProperty = str(sys.argv[2])
    e = datetime.datetime.now()
    nameNewFile = IccMainPart + IccProperty + e.strftime("%Y%m%d%H%M%S") + ".xlsx"
    auxFile = "aux1.xlsx"
    pathfile = 'C:/projects/bi/testng.xlsx'

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(nameNewFile, engine='xlsxwriter')
    writer.save()
    writer1 = pd.ExcelWriter(auxFile, engine='xlsxwriter')
    writer1.save()

    columns = ['partyName','LocationID','fsrName','Region','edRegion','salesLocation','dwIndicator','projectName','boardName',
               'boardsYr1','Prototype Date','Production Date','AUN','regstatus','regNum','faeName','supplier','supplierDivision','supplierPartNumber',
               'dwMargin','submittedDate','vendorApprovedDate','vendorExpirationDate','designWinDate','regProjectedYr1Rev','regNSBtoDate','fiscalMonth',
               'fiscalQtr','fiscalYear','metric','Note','topOpportunityFlag','DWB','BSV','icc2Name','icc3Name']
    
    df = pd.read_excel(pathfile)
    xl = pd.ExcelFile(pathfile)
    df.sort_values(by='Prototype Date')
    projectName_change = df["projectName"].shift() != df["projectName"]
    filesheet = xl.sheet_names
    infofile = xl.parse(filesheet[0])
    totalRows = infofile.shape[0] - 1
    testing = df.iat[0,7]
    columnIcc = 35 #icc3Name
    index = 0
    while index <= totalRows:
        if ((totalRows == index) and (projectName_change.iloc[index] == True)):
            if(df.iat[index,columnIcc] == IccMainPart):
                dataB1=df.loc[[index]]
                iccRowFound = pd.DataFrame(dataB1)
                df1=pd.read_excel(nameNewFile)
                df3=pd.concat([df1,iccRowFound])
                df3.to_excel(nameNewFile,index=False)
            index += 1
        elif((projectName_change.iloc[index] == True) and (projectName_change.iloc[index+1] == True)):
            if(df.iat[index,columnIcc] == IccMainPart):
                dataB1=df.loc[[index]]
                iccRowFound = pd.DataFrame(dataB1)
                df1=pd.read_excel(nameNewFile)
                df3=pd.concat([df1,iccRowFound])
                df3.to_excel(nameNewFile,index=False)
            index += 1
        elif ((projectName_change.iloc[index] == True) and (projectName_change.iloc[index+1] == False)):
            flagIccMainPart = 0
            flagIccProperty = 0
            auxCounter = 1
            if(df.iat[index,columnIcc] == IccMainPart):
                dataB1=df.loc[[index]]
                iccRowFound = pd.DataFrame(dataB1)
                df1=pd.read_excel(auxFile)
                df3=pd.concat([df1,iccRowFound])
                df3.to_excel(auxFile,index=False)
                flagIccMainPart = 1
            if(df.iat[index,columnIcc] == IccProperty):
                flagIccProperty = 1
            while (projectName_change.iloc[index+auxCounter] == False):
                if(df.iat[index+auxCounter,columnIcc] == IccMainPart):
                    dataB1=df.loc[[index+auxCounter]]
                    iccRowFound = pd.DataFrame(dataB1)
                    df1=pd.read_excel(auxFile)
                    df3=pd.concat([df1,iccRowFound])
                    df3.to_excel(auxFile,index=False)
                    flagIccMainPart = 1
                if(df.iat[index,columnIcc] == IccProperty):
                    flagIccProperty = 1
                auxCounter += 1
                
            if ((flagIccMainPart == 1) and (flagIccProperty == 0)):
                df1=pd.read_excel(auxFile)
                df2=pd.read_excel(nameNewFile)
                df3=pd.concat([df1,df2])
                df3.to_excel(nameNewFile,index=False)

            if os.path.exists(auxFile):
                writer1.handles.close()
                os.remove(auxFile)
                writer1 = pd.ExcelWriter(auxFile, engine='xlsxwriter')
                writer1.save()
            index = index + auxCounter - 1
        else:
            #do nothing
            index += 1
    writer1.handles.close()
    os.remove(auxFile)
else:
    print("Error - Bad parameters")
    print('Ej: py icc3.py iccPart iccPorperty')
        
