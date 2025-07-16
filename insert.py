import openpyxl
import pandas as pd

workbook = openpyxl.load_workbook("Full Data Export for Aidan.xlsx") 
  
sheet = workbook.active

data = []
for row in sheet.iter_rows(min_row=1, min_col=1, max_row=444, max_col=21): #max row 444
    rowf = []
    for cell in row: 
        rowf.append(cell.value)
        
        #print(cell.value, end=" ") 
    #print()
    data.append(rowf)
    #print(rowf) 

###########for key/value dict just ignore the description and title and name

df = pd.DataFrame(data)

header = ()

datasettuple = []
index = True
for row in df.itertuples():
    if index:
        index = False
        header = (row[1], row[12], row[13], row[2], row[3], row[17], row[4], row[5], row[6], row[7], row[8], row[9], row[15], row[14], row[16], row[10], row[11])
    else:
        datasettuple.append((row[1], row[12], row[13], row[2], row[3], row[17], row[4], row[5], row[6], row[7], row[8], row[9], row[15], row[14], row[16], row[10], row[11].strftime("%Y-%m-%d %H:%M")))

"""
print(header)
print(datasettuple[0])
print(datasettuple[0][0])
datasettuple[0][0] = "hello"
print(datasettuple[0][0])
"""

dict = {}

with open("dictKeyValues.txt", "r") as file:
    tmp = ""
    for line in file:
        
        tmp = line.split('"')

        dict[tmp[3]] = tmp[1]

with open("injectionCode.sql", "w") as file:
    file.write("""USE [IWS_Web_Forms]
GO
 
INSERT INTO [dbo].[FirstAidForm]
    ([Title]
    ,[InjuredWorkerID]
    ,[InjuredWorkerName]
    ,[Department]
    ,[WorkArea]
    ,[OPNbr]
    ,[IncidentDetails]
    ,[InjuryStatus]
    ,[InjuryType]
    ,[POBSide]
    ,[POBInjured]
    ,[SupplyUsed]
    ,[WitnessID]
    ,[WitnessName]
    ,[FARespondentID]
    ,[FARespondentName]
    ,[Inserted_DTTM])
     VALUES\n\t""")
    
    for x in range(443):
        file.write("(")
        for i in range(17):
            item = datasettuple[x][i]
            
            if (3 <= i <=4) or (5 <= i <= 11):
                if str(item) in dict:
                    item = dict[str(item)]

            if item == None:
                item = '""'

            if item == datasettuple[-1][-1]:
                file.write(str(item) + ")\nGO")
            elif item == datasettuple[x][-1]:
                file.write(str(item) + "),\n\t")
            else:
                if i in [1, 12, 14]:
                    file.write(str(item) + ", ")
                else:
                    if item == '""':
                        file.write(str(item) + ", ")
                    else:
                        file.write('"' + str(item) + '", ')

    


        """"
        for item in datasettuple[x]:
            if item == datasettuple[x][-1]:
                file.write(str(item) + ")\n\t")
            else:
             file.write(str(item) + ",")
        """

