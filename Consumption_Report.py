# -*- coding: utf-8 -*-
"""
Created on Thu Sep  7 15:46:52 2023

@author: harshit.kumar01
"""

import pandas as pd
import datetime
import os
import requests

 

os.chdir(r"D:\Consumption")
os.getcwd()
 

print("Job Started at : "  + str(datetime.datetime.time(datetime.datetime.now())))

 

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
 

Consumption1 = pd.DataFrame()

 
#Getting Data from Salesforce
Consumption1 = "https://pinelabs.my.salesforce.com/services/data/v56.0/query/?q=select+ProductConsumedNumber,Id,\
                     ProductItem.Serial_Number__c,ProductItem.ProductItemNumber,ProductItem.Location.Region__c,\
                         ProductItem.ProductName,ProductItem.Location.Name,ProductItem.QuantityOnHand,WorkOrder.Case.CaseNumber,\
                             WorkOrder.RootWorkOrder.WorkOrderNumber,WorkOrder.Status,WorkOrder.City,\
                    WorkOrder.ServiceTerritory.Name,WorkOrder.StatusCategory,WorkOrder.Subject,workorder.Owner.Name,\
                   WorkOrder.Account.Name,WorkOrderId,WorkOrder.IsClosed,QuantityConsumed,WorkOrderLineItem.LineItemNumber,\
                       WorkOrderLineItem.Location.Emp_Code__c,POS_ID__c,WorkOrder.Account.Parent.Name,WorkOrderLineItemId,WorkOrderLineItem.Type_of_Work__c,\
                        LastModifiedBy.Name,CreatedDate,WorkOrder.Account.Parent.Account_Type_Text__c,WorkOrderLineItem.External_Id__c,LastModifiedDate,\
                            WorkOrder.Location.Name,CreatedById,CreatedBy.Emp_ID__c,ProductItem.EDC_Condition_Type__c,ProductItem.Product_Category__c,Product_Type__c,Product2.Name,CreatedBy.Name\
                          +from+ProductConsumed+where+LastModifiedDate+>+2025-08-01T00:00:00.000Z"
 

r1 = requests.get(Consumption1, headers=headers)
Consumption1 = pd.json_normalize(r1.json()['records'])
Consumption1 = Consumption1._append(Consumption1)
try:
    nextRecordsUrl = r1.json()['nextRecordsUrl']
    while True:
        #print(url)
        url = 'https://pinelabs.my.salesforce.com'+r1.json()['nextRecordsUrl']
        r1 = requests.get(url, headers=headers)
        n1 = pd.json_normalize(r1.json()['records'])
        Consumption1 = pd.concat([Consumption1,n1])
        try:
            nextRecordsUrl = r1.json()['nextRecordsUrl']
        except:
            break
except:
    print('Q1 pull')
    
# Consumption1.to_excel("Consumption1.xlsx",index=False)

#2nd Query
Consumption2 = pd.DataFrame()
Consumption2 = "https://pinelabs.my.salesforce.com/services/data/v56.0/query/?q=select+ServiceAppointment.AppointmentNumber,ServiceAppointment.Work_Order_Line_Item__r.External_Id__c,\
    CreatedDate,LastModifiedBy.Emp_ID__c,CreatedBy.Emp_ID__c,Assigned_Resource__r.ServiceResource.Location.Emp_Code__c,\
        Assigned_Resource__r.ServiceResource.Location.Name,Device_Sell__c+from+ServiceAppointment+WHERE+LastModifiedDate+>+2025-04-01T00:00:00.000Z"

r2 = requests.get(Consumption2, headers=headers)
Consumption2 = pd.json_normalize(r2.json()['records'])
Consumption2 = Consumption2._append(Consumption2)
try:
    nextRecordsUrl = r2.json()['nextRecordsUrl']
    while True:
        #print(url)
        url = 'https://pinelabs.my.salesforce.com'+r2.json()['nextRecordsUrl']
        r2 = requests.get(url, headers=headers)
        n2 = pd.json_normalize(r2.json()['records'])
        Consumption2 = pd.concat([Consumption2,n2])
        try:
            nextRecordsUrl = r2.json()['nextRecordsUrl']
        except:
            break
except:
    print('Q2 pull')
    
# Consumption2.to_excel("Consumption2.xlsx",index=False)



Consumption1['CreatedDate'] = pd.to_datetime(Consumption1['CreatedDate']).dt.strftime('%d-%m-%Y')
Consumption1['CreatedDate'] = Consumption1['CreatedDate'].apply(lambda x: datetime.datetime.strptime(x,'%d-%m-%Y'))
Consumption1['LastModifiedDate'] = pd.to_datetime(Consumption1['LastModifiedDate']).dt.strftime('%d-%m-%Y')
Consumption1['LastModifiedDate'] = Consumption1['LastModifiedDate'].apply(lambda x: datetime.datetime.strptime(x,'%d-%m-%Y'))


Consumption2 = Consumption2.rename(columns={'Work_Order_Line_Item__r.External_Id__c': 'WorkOrderLineItem.External_Id__c'})
Consumption3 = pd.merge(Consumption1,Consumption2[['WorkOrderLineItem.External_Id__c','AppointmentNumber','LastModifiedBy.Emp_ID__c',\
                        'Assigned_Resource__r.ServiceResource.Location.Emp_Code__c','Assigned_Resource__r.ServiceResource.Location.Name','Device_Sell__c']], on ='WorkOrderLineItem.External_Id__c',how = 'left')

Consumption3['duplicated_flag'] = Consumption3['ProductConsumedNumber'].duplicated()
Consumption3 = Consumption3[Consumption3.duplicated_flag == False]
Consumption3 = Consumption3.drop(['duplicated_flag'],axis=1)
# Consumption3.to_csv('Consumption3.csv',index = False)

 

def action1(cell_value):
    if pd.isna(cell_value) or cell_value == "":
          return "Non-serialized"
    else:
          response = 'Serialized'
    return response
Consumption3['Serialized / Non Serialized'] = Consumption3['ProductItem.Serial_Number__c'].apply(action1)

 

category = pd.read_excel(r'D:\Store -Summery Folder\Model Name_Category.xlsx' , engine='openpyxl')

 

Consumption3 = Consumption3.rename(columns = {'ProductItem.ProductName':'Product Name: Product Name'})
Consumption4 = pd.merge(Consumption3, category[['Product Name: Product Name','{Sub-Model Name}','{Product Type}']],on = 'Product Name: Product Name', how ='left')

 

    

 

final_comsumption = Consumption4[['ProductConsumedNumber','ProductItem.Serial_Number__c','WorkOrder.RootWorkOrder.WorkOrderNumber','WorkOrderLineItem.LineItemNumber','WorkOrder.Case.CaseNumber',\
                          'WorkOrder.Status','WorkOrder.City','WorkOrder.ServiceTerritory.Name',\
                'ProductItem.Location.Region__c','WorkOrder.StatusCategory','WorkOrder.Subject','WorkOrder.Owner.Name',\
            'Id','Product Name: Product Name','WorkOrder.Account.Parent.Name','WorkOrder.Account.Name','WorkOrder.IsClosed','ProductItem.QuantityOnHand',\
        'QuantityConsumed','CreatedDate','LastModifiedDate','LastModifiedBy.Name','WorkOrderLineItem.External_Id__c','WorkOrderId',\
            'ProductItem.ProductItemNumber','WorkOrderLineItemId','POS_ID__c',\
        'ProductItem.Serial_Number__c','Assigned_Resource__r.ServiceResource.Location.Emp_Code__c','Assigned_Resource__r.ServiceResource.Location.Name','AppointmentNumber','WorkOrder.Account.Parent.Account_Type_Text__c','Product Name: Product Name',\
            '{Sub-Model Name}','{Product Type}','Serialized / Non Serialized','WorkOrderLineItem.Type_of_Work__c','ProductItem.EDC_Condition_Type__c','ProductItem.Product_Category__c','Device_Sell__c']]

 


column_rename = {
    'ProductConsumedNumber': 'Product Consumed Number',
    'ProductItem.Serial_Number__c': 'Product Item: Serial Number',
    'WorkOrder.Case.CaseNumber': 'Case: Case Number',
    'WorkOrder.Status': 'Status',
    'WorkOrder.City': 'City',
    'POS_ID__c': ' POS ID',
    'WorkOrder.ServiceTerritory.Name': 'Service Territory: Name',
    'ProductItem.Location.Region__c': 'Region',
    'WorkOrder.StatusCategory': 'Status Category',
    'WorkOrder.Subject': 'Subject',
    'WorkOrder.Owner.Name': 'Owner: Full Name',
    'Id': 'Product Consumed ID',
    'WorkOrder.Account.Name': 'Account: Account Name',
    'WorkOrder.IsClosed': 'Is Closed',
    'WorkOrder.Account.Parent.Name': 'Parent Account Name',
    'ProductItem.QuantityOnHand': 'Product Item: Quantity On Hand',
    'QuantityConsumed': 'Quantity Consumed',
    'CreatedDate_x': 'Created Date',
    'Device_Sell__c': 'Device Sell',
    'LastModifiedDate': 'Date-Wise',
    'LastModifiedBy.Name': 'Last Modified By: Full Name',
    'WorkOrderLineItem.External_Id__c': 'Work Order Line Item: External Id',
    'WorkOrderId': 'Work Order ID',
    'ProductItem.EDC_Condition_Type__c': 'EDC Condition',
    'ProductItem.Product_Category__c': 'Product Category',
    'WorkOrderLineItem.LineItemNumber': 'Line Items',
    'WorkOrderLineItemId': 'Work Order Line Item ID',
    'WorkOrder.Account.Parent.Account_Type_Text__c': 'Parent.Account-Type',
    'Assigned_Resource__r.ServiceResource.Location.Emp_Code__c':'Emp ID',
    'Assigned_Resource__r.ServiceResource.Location.Name': 'Location Name',    
    'ServiceAppointment.AppointmentNumber': 'Service Appointment',     
    'Product_Type__c': 'Type',    
    'WorkOrderLineItem.Type_of_Work__c': 'Type Of Work',
}

 

final_comsumption.rename(columns=column_rename, inplace=True)

 

final_comsumption['duplicated_flag'] = final_comsumption['Product Consumed Number'].duplicated()
final_comsumption = final_comsumption[final_comsumption.duplicated_flag == False]
final_comsumption = final_comsumption.drop(['duplicated_flag'],axis=1)

 

 

final_comsumption.to_excel('SF_Consumption.xlsx', index=False)