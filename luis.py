########### Python 3.6 #############
import requests
import openpyxl
from pandas import DataFrame
from datetime import datetime
import os, errno
import sys

headers = {
    # Request headers
    'Ocp-Apim-Subscription-Key': 'f3b40c62ca074905bafc2a8add33414d',
}
if len(sys.argv[1]) == 0:
    print ("please enter valid folder path")
    sys.exit()
else:
    foledr_path = sys.argv[1]
	
wb = openpyxl.load_workbook(foledr_path)
print("Your data is Processing ......................")
#wb.get_sheet_names()
#heet = wb.get_sheet_by_name('Sheet1')
sheet = wb.active
rowsRequired = []
rowsNotRequired = []
for cellObj in sheet["A"]:
    #print(cellObj.value)
    params ={
        # Query parameter
        'q': cellObj.value,
        # Optional request parameters, set to default values
        'timezoneOffset': '0',
        'verbose': 'false',
        'spellCheck': 'false',
        'staging': 'false',
    }

    try:
        r = requests.get('https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/9119be6a-745a-4ba5-8958-d70333316394?',headers=headers, params=params)
        print(r.json())
        #print("Query is : " + r.json()['query'])
        #print("intent is : "+ r.json()['topScoringIntent']['intent'] +" - "+ str(r.json()['topScoringIntent']['score']))
        if r.json()['topScoringIntent']['intent'] == "Website Development":
           rowsRequired.append(r.json()['query'])
        else:
           rowsNotRequired.append(r.json()['query'])	

    except Exception as e:
        pass

print("Writing data into files ...............................")
#print(rowsRequired)
#print(rowsNotRequired)
foldername = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
foldername = foldername.replace(" ", "_")
foldername = foldername.replace(":", "_")
try:
    os.makedirs(foldername)
except OSError as e:
    pass
    if e.errno != errno.EEXIST:
       raise
	   
dfRequired = DataFrame({'Website Development': rowsRequired})
dfRequired.to_excel(foldername+'/'+'Required.xlsx', sheet_name='sheet1', index=False)
dfNotRequired = DataFrame({'Not Rquired 	': rowsNotRequired})
dfNotRequired.to_excel(foldername+'/'+'Not Required.xlsx', sheet_name='sheet1', index=False)