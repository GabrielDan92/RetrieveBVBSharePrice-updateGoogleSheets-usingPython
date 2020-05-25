import ezsheets, os, datetime, threading, time
from selenium import webdriver

os.chdir("C:\\Users\\user\\Desktop\\programming\\Python")
pricesRetrieved = False 
symbolDict = {"BRD": 0, "TLV": 0, "SNG": 0, "SNN": 0, "FP": 0}

# retrieve the prices from bvb's website
def stock():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--incognito")
    #options.add_argument("--headless")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    browser = webdriver.Chrome(executable_path=r"C:\Users\user\Desktop\programming\Python\chromedriver.exe", options=options)
    browser.delete_all_cookies()
    browser.implicitly_wait(10)         # wait 10 seconds for each element to be available
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {
        get: () => undefined
        })
    """
    })
    browser.execute_cdp_cmd("Network.enable", {})
    browser.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"User-Agent": "browser1"}})
    link = "http://www.bvb.ro/FinancialInstruments/Details/FinancialInstrumentsDetails.aspx?s="
    print("Accessing BVB's website...")

    for i in symbolDict:
        browser.get(link + i)                                       #build the link
        price = browser.find_element_by_class_name("value").text
        print("The price for " + i + " is: " + price)               #get the current price
        price = price.replace(".", ",")                             #replace the dot (.) with comma (,) for Google Sheets' format to recognize the price correctly
        symbolDict[i] = price
        
    print(symbolDict)
    browser.quit()

    global pricesRetrieved    
    pricesRetrieved = True
    return pricesRetrieved

# save the results to Google Sheets
def googleSheets():

    sheets = ["BRD", "BANCA TRANSILVANIA", "ROMGAZ", "NUCLEARELECTRICA", "FONDUL PROPRIETATEA"]
    # assign the Google Sheets workbook id
    ss = ezsheets.Spreadsheet("")
    print("Accessing the Google Sheet '" + str(ss.title) + "'...")

    # keep looping until the prices have been retrieved
    # while True:
    #     print("Waiting for prices to be retrieved...")
    #     if pricesRetrieved == True:
    #         break
    #     time.sleep(.5)
    #     continue

    j = 0
    for i in symbolDict:
        sheet = ss[sheets[j]]
        print("Accessing sheet '" + str(sheet.title) + "'...")
        sheet["P2"] = symbolDict[i]
        print("Adding price " + symbolDict[i] + " for symbol " + i + "...")
        print(i + ": " + symbolDict[i])
        j += 1
    
    now = datetime.datetime.now()       #retrieve the current time
    sheet = ss["Dashboard"]             #change the active sheet to the 'Dashboard' sheet

    # use zfill to add leading "0" to one digit values
    sheet["G5"] = "LAST KNOWN MARKET VALUE (" + str(now.day).zfill(2) + "/" + str(now.month).zfill(2) + "," + str(now.hour).zfill(2) + ":" + str(now.minute).zfill(2) + ")"

    sheet = ss["historicData"]          #change the active sheet to the 'historicData' sheet
    print("Accessing sheet '" + str(sheet.title) + "'...")

    i = 1
    while sheet.get(1,i) != '':
        i+=1
    lastRow = i                         #find the last row
    # add the date on the last row
    print("Adding date and time in the row: " + str(lastRow))
    sheet[1, lastRow] = str(now.day).zfill(2) + "/" + str(now.month).zfill(2) + "," + str(now.hour).zfill(2) + ":" + str(now.minute).zfill(2)
    
    j = 2
    for i in symbolDict:                #add the prices for each symbol on the last row
        print("Adding price " + symbolDict[i] + " for symbol " + i + " on column " + str(j) + ", row " + str(lastRow) + "...")
        sheet[j, lastRow] = symbolDict[i]
        j += 1 

    ss.refresh()

# call the stock() function in advance
threadObj = threading.Thread(target=stock)
threadObj.start()

# call the google sheets() function
googleSheets()
