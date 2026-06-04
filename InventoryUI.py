# -*- coding: utf-8 -*-
"""
Created on Fri Sep  8 17:33:26 2023

@author: harshit.kumar01
"""
import pandas as pd
import os
import requests
import datetime
from datetime import date
today = date.today()


print("job started at : " + str(datetime.datetime.time(datetime.datetime.now())))


#Generating Token
print("Data from Salesforce Download Started at : " + str(datetime.datetime.time(datetime.datetime.now())))
def get_token():
    client_id = '3MVG9pe2TCoA1Pf4AUXJnyzPmstC0CcF5l_8P5uP77_GADjrPL2kiZCmvc2zQpxs6anlBwbQa0byW5iaf0M23'
    client_secret = '96A440D33246A148DC3FBDE4915F512FC9CC91CD08AC00DABD540B0249AECDE4'
    username='integration.user.maple@pinelabs.com.prod'
    password='Windows98'
    t_url = 'https://login.salesforce.com/services/oauth2/token?grant_type=password&client_id='+client_id+'&client_secret='+client_secret+'&username='+username+'&password='+password
    r = requests.post(t_url)
    access_token = r.json()['access_token']
    return access_token
print(get_token)
access_token = get_token()
headers = {'Authorization':'Bearer '+ access_token}  
 

os.chdir(r"D:\SF Inventory")
os.getcwd()
completed_df = pd.DataFrame()
 
#Getting Data from Salesforce
Inventory = "https://pinelabs.my.salesforce.com/services/data/v56.0/query/?q=select+ProductItemNumber,\
    Serial_Number__c,Location.Emp_Code__c,Location.Name,Location.City_Master__r.Name,\
    Account__r.City__r.Name,Location.Region__c,Id,Account__r.City__r.Region__c,Account__r.Name,ProductName,LastModifiedBy.Name,\
        CreatedDate,LastModifiedDate,SerialNumber,Product_Type__c,Account__r.City__r.State__c,Location.City_Master__r.SC_Territory__c,\
    Product_Working_Condition_Status__c,Product_Working_Condition_Sub_Status__c,Product2.Downstream_POS_ID__c,\
     Product2.Name,Location.LocationType,QuantityOnHand,EDC_Condition_Type__c,In_Transit_text__c,Product_Category__c,\
      Ageing_bucket__c+from+ProductItem+where+Product_Working_Condition_Status__c Not +IN+('Live')\
     +AND+ProductName+NOT+IN+('SIM')\
   +AND+QuantityOnHand <>0" 

r1 = requests.get(Inventory, headers=headers)
Inventory = pd.json_normalize(r1.json()['records'])
Inventory = Inventory._append(Inventory)
try:
    nextRecordsUrl = r1.json()['nextRecordsUrl']
    while True:
        #print(url)
        url = 'https://pinelabs.my.salesforce.com'+r1.json()['nextRecordsUrl']
        r1 = requests.get(url, headers=headers)
        n1 = pd.json_normalize(r1.json()['records'])
        Inventory = pd.concat([Inventory,n1])
        try:
            nextRecordsUrl = r1.json()['nextRecordsUrl']
        except:
            break
except:
    print('Q1 pull')
    
# Product_consumed.to_excel("output.xlsx",index=False)


category = pd.read_excel(r'D:\Store -Summery Folder\Model Name_Category.xlsx' , engine='openpyxl')

Inventory
Inventory['CreatedDate'] = pd.to_datetime(Inventory['CreatedDate']).dt.strftime('%d-%m-%Y')
Inventory['CreatedDate'] = Inventory['CreatedDate'].apply(lambda x: datetime.datetime.strptime(x,'%d-%m-%Y'))
Inventory['LastModifiedDate'] = pd.to_datetime(Inventory['LastModifiedDate']).dt.strftime('%d-%m-%Y')
Inventory['LastModifiedDate'] = Inventory['LastModifiedDate'].apply(lambda x: datetime.datetime.strptime(x,'%d-%m-%Y'))


def action1(cell_value):
    if pd.isna(cell_value) or cell_value == "":
          return "Non-serialized"
    else:
          response = 'Serialized'
    return response
Inventory['Serialized / Non Serialized'] = Inventory['SerialNumber'].apply(action1) 

   
Inventory['LastModifiedDate'] = pd.to_datetime(Inventory['LastModifiedDate'])
today = pd.to_datetime(today)
Inventory['Ageing Days'] = (today - Inventory['LastModifiedDate']).dt.days

def categorize_days(days):
      if days >= 0 and days <= 4:
        return "0 - 4 Days"
      elif days >= 5 and days <= 15:
        return "05 - 15 Days"
      elif days >= 16 and days <= 30:
        return "15 - 30 Days"
      elif days > 30 and days <= 60:
        return "Above 30 Days"
      elif days > 60 and days <= 90:
        return "Above 60 Days"
      elif days > 90:
        return "Above 90 Days"
      else:
        return "Invalid Input" 
Inventory['Ageing Bucket2'] = Inventory['Ageing Days'].apply(categorize_days)
    
def categorize_years(days):
    if days >=0 and days <= 365:
        return "0-1 Yr"
    elif days >= 365 and days <= 730:
        return "1-2 Yr"
    elif days >=730 and days <= 1095:
        return "2-3 Yr"
    elif days >=1095 and days<= 1460:
        return "3-4 Yr"
    elif days >=1460 and days<= 1825:
        return "4-5 Yr"
    else:
        return "5+ Yrs"
Inventory['Ageing Bucket1']= Inventory['Ageing_bucket__c'].apply(categorize_years)        
   
Model_type = category[['Product Name: Product Name','{Model Name}','{Sub-Model Name}','{Product Type}','{Category}','Device Type']]
Inventory.rename(columns={'Product2.Name': 'Product Name: Product Name'}, inplace=True)
Inventory2= pd.merge(Inventory, Model_type, on = 'Product Name: Product Name',how='left')


Location_type = category[['Location: Location Name','Location Name','Location Type','{State}','{Region}','City Name','FSR Status']]
Inventory2.rename(columns = {'Location.Name' : 'Location: Location Name'}, inplace=True)
Inventory3 = pd.merge(Inventory2 , Location_type, on = 'Location: Location Name',how='left' )

final_Inventory=Inventory3[['ProductItemNumber','Serial_Number__c','Product Name: Product Name','Location.Emp_Code__c','Location: Location Name',\
            'Location.City_Master__r.Name','Location.Region__c','Id','LastModifiedBy.Name','CreatedDate','LastModifiedDate','Serial_Number__c',\
                'Product_Type__c','Location.City_Master__r.SC_Territory__c','Product_Working_Condition_Status__c',\
'Product_Working_Condition_Sub_Status__c','Product_Category__c','Product2.Downstream_POS_ID__c','ProductName','Location.LocationType','QuantityOnHand',\
    'EDC_Condition_Type__c','In_Transit_text__c','Ageing Bucket1','Serialized / Non Serialized','{Model Name}','{Sub-Model Name}',\
        '{Product Type}','{Category}','Device Type','Location: Location Name','Location Name','Location Type','{State}','{Region}','FSR Status',\
            'Ageing Days','Ageing Bucket2']]

 
column_rename = {
    'ProductItemNumber': 'Product Item Number',
    'Serial_Number__c': 'Serial Number',
    'Product Name: Product Name': 'Product Name: HardwareModel',
    'Location.Emp_Code__c': 'Location: Emp Code',
    'Location.Name': 'Location: Location Name',
    'Location.City_Master__r.Name': 'Location: City Master: City Name',
    'Location.Region__c': 'Location: Region',
    'Id': 'Product Item ID',
    'LastModifiedBy.Name': 'Last Modified By: Full Name',
    'CreatedDate': 'Created Date',
    'Product_Category__c': 'Product Category',
    'LastModifiedDate': 'Last Modified Date',
    'Product_Type__c': 'Product Type',
    'Location.City_Master__r.SC_Territory__c': 'Location: State',
    'Location Name': 'Sub-Location Name',
    'Product_Working_Condition_Status__c': 'Product Working Condition Status',
    'Product_Working_Condition_Sub_Status__c': 'Product Working Condition Sub Status',
    'Product2.Downstream_POS_ID__c': 'Product Name: Downstream POS ID',
    'Product2.Name': 'Product Name: Product Name',
    'Location.LocationType': 'Location: Location Type',
    'QuantityOnHand': 'Quantity On Hand',
    'EDC_Condition_Type__c': 'EDC Condition Type',
    'In_Transit_text__c': 'In Transit?',    
}


final_Inventory.rename(columns=column_rename, inplace=True)
final_Inventory = final_Inventory[final_Inventory['Quantity On Hand'] != 0]

final_Inventory.to_excel('D:\SF Inventory\Input\SF_Inventory.xlsx', index=False)





