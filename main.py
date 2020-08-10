import ezsheets
import datetime
import threading
import time
import requests
import os
from bs4 import BeautifulSoup

os.chdir("C:\\Users\\Gabriel\\Desktop\\programming\\Python")
pricesRetrieved = False 
symbolDict = {"BRD": 0, "TLV": 0, "SNG": 0, "SNN": 0, "FP": 0}
link = "http://www.bvb.ro/FinancialInstruments/Details/FinancialInstrumentsDetails.aspx?s="

# retrieve the prices from bvb's website
def stock():
    print("Accessing BVB's website...")
    for i in symbolDict:
        url = link + i
        user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
        data = requests.get(url=url, headers={'user-agent': user_agent})
        soup = BeautifulSoup(data.text, "html.parser")
        x = soup.find_all("strong", class_="value")  
        x = str(x).split(">")
        x = x[1].split("<")
        price = x[0]
        print("The price for " + i + " is: " + str(price))              
        price = price.replace(".", ",")                             #replace the dot (.) with comma (,) for Google Sheets' format to recognize the price correctly
        symbolDict[i] = price

# save the results to Google Sheets
def googleSheets():
    sheets = ["BRD", "BANCA TRANSILVANIA", "ROMGAZ", "NUCLEARELECTRICA", "FONDUL PROPRIETATEA"]
    ss = ezsheets.Spreadsheet("1zQJxdZZNFXsmOFWYEe3NdMaS2scpMr_BBE2SOJbecdI")
    print("Accessing the Google Sheet '" + str(ss.title) + "'...")

    j = 0
    for i in symbolDict:
        sheet = ss[sheets[j]]
        print("Accessing sheet '" + str(sheet.title) + "'...")
        sheet["P2"] = symbolDict[i]
        print("Adding price " + symbolDict[i] + " for symbol " + i + "...")
        print(i + ": " + symbolDict[i])
        j += 1
    
    now = datetime.datetime.now()      
    sheet = ss["Dashboard"]         

    # use zfill to add leading "0" to one digit values
    sheet["G5"] = "LAST KNOWN MARKET VALUE (" + str(now.day).zfill(2) + "/" + str(now.month).zfill(2) + "," + str(now.hour).zfill(2) + ":" + str(now.minute).zfill(2) + ")"
    sheet = ss["historicData"]         
    print("Accessing sheet '" + str(sheet.title) + "'...")
    # find the last row
    i = 1
    while sheet.get(1,i) != '':
        i+=1
    lastRow = i                        
    print("Adding date and time in the row: " + str(lastRow))
    sheet[1, lastRow] = str(now.day).zfill(2) + "/" + str(now.month).zfill(2) + "," + str(now.hour).zfill(2) + ":" + str(now.minute).zfill(2)
    # add the prices for each symbol on the last row
    j = 2
    for i in symbolDict:                
        print("Adding price " + symbolDict[i] + " for symbol " + i + " on column " + str(j) + ", row " + str(lastRow) + "...")
        sheet[j, lastRow] = symbolDict[i]
        j += 1 

# call the stock() function in advance
threadObj = threading.Thread(target=stock)
threadObj.start()
googleSheets()
