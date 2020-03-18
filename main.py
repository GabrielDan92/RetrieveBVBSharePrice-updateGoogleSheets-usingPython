import ezsheets, os
from selenium import webdriver

os.chdir("C:\\Users\\user\\Desktop\\programming\\Python")


######## RETRIEVE THE PRICES FROM BVB'S WEBSITE AND SAVE THEM IN A DICTIONARY #######

browser = webdriver.Chrome(executable_path=r"C:\Users\user\Desktop\programming\Python\chromedriver.exe")

# wait 10 seconds for each element to be available
browser.implicitly_wait(10)

# create the emitents list
symbolDict = {"BRD": 0, "TLV": 0, "SNG": 0, "SNN": 0, "FP": 0}
# assign the main link
link = "http://www.bvb.ro/FinancialInstruments/Details/FinancialInstrumentsDetails.aspx?s="

for i in symbolDict:
    # build the link
    browser.get(link + i)

    # get the current price
    price = browser.find_element_by_class_name("value").text
    print("The price for " + i + " is: " + price)

    #replace the dot (.) with comma (,) for Google Sheets' format to recognize the price correctly
    price = price.replace(".", ",")
    symbolDict[i] = price

print(symbolDict)
browser.quit()


####### PUSH THE RESULTS IN GOOGLE SHEETS #######

# create the spreadsheeets name list
sheets = ["BRD", "BANCA TRANSILVANIA", "ROMGAZ", "NUCLEARELECTRICA", "FONDUL PROPRIETATEA"]

# assign the Google Sheets workbook id
ss = ezsheets.Spreadsheet("1zQJxdZZNFXsmOFWYEe3NdMaS2scpMr_BBE2SOJbecdI")

print("Accessing workbook " + str(ss.title) + "...")

j = 0
for i in symbolDict:
    sheet = ss[sheets[j]]
    print("Accessing sheet " + str(sheet.title) + "...")
    sheet["P2"] = symbolDict[i]
    print("Adding price " + symbolDict[i] + " for symbol " + i + "...")
    print(i + ": " + symbolDict[i])
    j += 1

# refresh the Google Sheets workbook
ss.refresh()
