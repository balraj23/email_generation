from configparser import ConfigParser
import configparser
import psycopg2
import csv
import json
import pandas as pd
import xlwings as xw
import openpyxl
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import re

global email_from, email_pwd

def email():
    config = ConfigParser()
    config.read("app-email.properties")
    email_from =config.get("email","email_from")
    email_pwd = config.get("email","email_pwd")

def load_connection():
    data = None
    connection = {}
    with open('app_mutual_funds.json') as json_file:
        data = json.load(json_file)
    for database in data["databases"]:
        # print(database)
        key = list(database.values())[0]
        connection[key] = database
    return connection

def get_config_object():
    config = ConfigParser()
    config.read("app-service.properties")
    return config
    
def get_config_attributes(config):
    if config:
        host = config.get("database","host")
        port = config.get("database","port")
        database = config.get("database","database")
        username = config.get("database","user")
        password = config.get("database","password")         
    else:
        print("invalid config object.")
        return None            
    return host, port, database, username, password

def get_connection(properties):
    if properties:
        connection = psycopg2.connect(host = properties[0],
            port = properties[1],
            database = properties[2],
            user = properties[3],
            password = properties[4])
        if connection:
            return connection
        else:
            return None    
    else:
        print("plese create a property file correctly")
     
def get_query(connection):
    cursor = connection.cursor()
    select_query = "select query, database, group_name, subject, email_to, email_cc from public.admin " 
    cursor.execute(select_query)
    fetch_data = cursor.fetchall()    
    querys  = {}
    for row in fetch_data:
        query = row[0]
        database = str(row[1])
        group_name = row[2]
        subject = row[3] 
        email_to = row[4] 
        email_cc = row[5]      
        if database not in querys:
            querys[database] = []
        querys[database].append({"group_name": group_name, "query": query, "subject":subject,"email_to":email_to, "email_cc":email_cc})         
    cursor.close()
    if connection:
        connection.close()
    return querys

def get_to_connection_attributes(database,connection):
    db = database.keys()
    print(db)
    connection_attributes = None
    databases = list(connection.keys())
    for database in db:
        if database in databases:
            connection_attributes = connection[database]
        else:
            print("incorrect database")
        
    return connection_attributes,db

def get_to_connection(to_config_properties):
    print(to_config_properties)
    to_config_properties = to_config_properties[0]
    if to_config_properties:
        connection = psycopg2.connect(host = to_config_properties["host"],
            port = to_config_properties["port"],
            database = to_config_properties["database"],
            user = to_config_properties["user"],
            password = to_config_properties["password"])
        if connection:
            print("success")
            return connection            
        else:
            return None    
    else:
        print("plese create a property file correctly")
        
def execute_query(connection,database_query,to_config):
    db = str(next(iter(to_config[1])))
    querys = database_query[db]           
    for item in querys:         
        data_name_query = item['query']        
        data_name_query = eval(data_name_query)
        
        group_name = item['group_name'] 
        email_to = item['email_to']
        email_cc = item['email_cc']
        subject = item['subject']
        workbook = openpyxl.Workbook()       
        for data in data_name_query:            
            query_file_name = data["queryFileName"]
            t_query = data['query']
            print(t_query)
            cursor = connection.cursor()           
            create_worksheet(cursor, t_query,query_file_name,workbook)
        workbook.remove(workbook['Sheet'])        
        workbook.save(f'{group_name}.xlsx') 
        file_name = f'{group_name}.xlsx'
        # send_email(email_to,email_cc,subject,file_name)       

def create_worksheet(cursor,t_query,query_file_name,workbook):    
    cursor.execute(t_query)
    fetch_data = cursor.fetchall()
    df = pd.DataFrame(fetch_data,columns=[desc[0] for desc in cursor.description],index=None)        
    data = df.values.tolist()
    sheet_name = f"{query_file_name}"
    worksheet = workbook.create_sheet(title = sheet_name)
    for row in dataframe_to_rows(df, index=False, header=True):
        worksheet.append(row)
        
def send_email(email_to,email_cc,subject,file_name):
    msg = attach_files(email_to,email_cc,subject)
    with open(file_name, "rb") as file:
        send = MIMEBase("application", "octet-stream")
        send.set_payload(file.read())
        encoders.encode_base64(send)
        send.add_header("Content-Disposition", f"attachment; filename={file_name}")
        msg.attach(send)        
    with smtplib.SMTP("smtp.zoho.com", 587) as server:
        server.starttls()
        server.login(email_from, email_pwd)
        # server.send_message(msg)
       
def attach_files(email_too,email_ccc,subject):
    email_to = re.findall(r'"([^"]+)"', email_too)
    email_to = ';'.join(email for email in email_to)
    email_cc = re.findall(r'"([^"]+)"', email_ccc)
    email_cc =';'.join(email for email in email_cc)
    msg = MIMEMultipart()
    msg["From"] = email_from
    msg["To"] = email_to
    msg["Cc"] = email_cc
    msg["Subject"] = subject
    msg.attach(MIMEText("Plese refer to the attachments", "plain"))
    return msg
        
if __name__ == "__main__":
    email() 
    connection_object = load_connection() 
    config = get_config_object()
    properties = get_config_attributes(config)
    connection = get_connection(properties)
    database_query = get_query(connection)
    to_config_properties = get_to_connection_attributes(database_query,connection_object)
    to_connection = get_to_connection(to_config_properties) 
    excel_report = execute_query(to_connection,database_query,to_config_properties)
    if connection:
        connection.close()