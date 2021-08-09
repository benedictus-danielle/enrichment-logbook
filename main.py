import requests
from bs4 import BeautifulSoup as bs
import pandas
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()
session = requests.session()

ENRICHMENT_BASE_URL = "https://enrichment.apps.binus.ac.id"
ACTIVITY_ENRICHMENT_BASE_URL = "https://activity-enrichment.apps.binus.ac.id"
ENRICHMENT_LOGIN_URL = f"{ENRICHMENT_BASE_URL}/Login/Student/Login"

def getRequestVerificationToken():
    token = ""
    r = session.get(ENRICHMENT_LOGIN_URL)
    soup = bs(r.text, 'html.parser')
    token = soup.find('input', {"name":"__RequestVerificationToken"}).get('value')
    return token

def login():
    username = os.getenv("EMAIL")
    password = os.getenv("PASSWORD")
    token = getRequestVerificationToken()
    data = {
        "login.Username":username,
        "login.Password":password,
        "__RequestVerificationToken": token,
        "btnLogin":"Login"
    }
    r = session.post(f"{ENRICHMENT_BASE_URL}/Login/Student/DoLogin", data)

def goToActivityEnrichment():
    login()
    r = session.get(f"{ENRICHMENT_BASE_URL}/Login/Student/SSOToActivity")
    pass

def getLogbook():
    goToActivityEnrichment()
    r = session.get(f"{ACTIVITY_ENRICHMENT_BASE_URL}/LogBook/GetMonths")
    return r.json()

logbooks = getLogbook()["data"]

data = pandas.read_excel("template.xlsx", engine='openpyxl')
data['Month'] = data['Date'].dt.month_name()
data['InsertDate'] = data['Date'].dt.strftime("%Y-%m-%dT00:00:00")

for index, row in data.iterrows():
    logbook = next(logbook for logbook in logbooks if logbook["month"] == row["Month"])
    logbookHeaderID = logbook["logBookHeaderID"]
    r = session.post(f"{ACTIVITY_ENRICHMENT_BASE_URL}/LogBook/GetLogBook",{
        "logBookHeaderID": logbookHeaderID
    })
    logbookMonth = r.json()["data"]
    logBookRow = next(x for x in logbookMonth if x["date"] == row["InsertDate"])
    data = {
        "ID": logBookRow["id"],
        "LogBookHeaderID": logbookHeaderID,
        "Date":row["InsertDate"],
        "Activity":row["Activity"],
        "ClockIn":row["ClockIn"],
        "ClockOut":row["ClockOut"],
        "Description": row["Description"]
    }
    r = session.post(f"{ACTIVITY_ENRICHMENT_BASE_URL}/LogBook/StudentSave",data)
    if r.json()["json"]:
        print(f"Logbook for {row['Date']} success to insert or update")
    else:
        print(f"Logbook for {row['Date']} fail to insert or update")

