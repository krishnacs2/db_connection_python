import cv2
import glob
import os
import requests
import numpy as np
import sys
import os.path
import json

if len(sys.argv[1]) == 0:
    print ("please enter currect input folder path")
    sys.exit()
else:
    foledr_path = sys.argv[1]
	
url = "https://api.vize.ai/v1/classify/"
headers = {"Authorization": "Token 738585a9a46881b1f54542b4a536c9960a104fb1"}
data = {"task":"161a76e2-9b74-499f-b09d-118cddb7c5bb"}
payload = {'isOverlayRequired': "false",
   'apikey': "9c80d1598188957",
   'language': "eng",
   }
path = foledr_path
folders = os.listdir(path)
print(folders)
i = 0
dataArray = []
for folder in folders:
   print(folder)
   images = os.listdir(path+"/"+folder)
   i = i+1
   print(i)
   index = str(i)
   dataArray.append("row " + index  )
   for image in images:
	   print(image)
	   files = {'image_file': open(path+"/"+folder +'/'+image, 'rb')} #use path to your image
	   response = requests.post(url, headers=headers, files=files, data=data)
	   if response.json()["best_label"]["label_name"] == "label" or response.json()["best_label"]["label_name"] == "button":
            filename = path+"/"+folder +'/'+image
            with open(filename, 'rb') as f:
               r = requests.post('https://api.ocr.space/parse/image',files={filename: f},data=payload) 
            #print(r.content)   
            result = json.loads(r.content)
            #print(result['ParsedResults'])			
            row_json = json.dumps(result['ParsedResults'][0]['ParsedText'])
            #print(row_json)	
            row_json_result = str(row_json)
            row_json_result = row_json_result.replace('"',"")
            row_json_result = row_json_result[:-4]
            #print(row_json_result)
            dataArray.append(response.json()["best_label"]["label_name"]+":"+row_json_result)			  
	   else:
		   dataArray.append(response.json()["best_label"]["label_name"])
		   #print(response.json()["best_label"]["label_name"])
	   
print(i)
print(dataArray)
