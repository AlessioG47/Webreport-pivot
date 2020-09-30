import time, datetime
import traceback
import pandas as pd
import numpy as np
import xlsxwriter
import os
from IPython.display import display_html
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

i = 1 # Respect the Almighty Loop Tracker! Praise be to i

logName = 'Pivot report run'

sleepTimer = 2  # - Increase wait timer depending on stability of reporting website, 1,2 or 3+

# ARRAYS
print('Reading country list...')
dfem = pd.read_excel(open('C://Users//alessio//Desktop//Python Scripts//Country Link list.xlsx','rb'), sheet_name='Sheet1')

countries = dfem['Country'].tolist()
links = dfem['Weblink'].tolist()

#CREATE NEW FOLDER
str1 = "/"
date = datetime.date.today()
path = str('C:/Users/alessio/Desktop/Python Scripts/Pivot Reports/2020')+str1+str(date)
os.mkdir(path)

# BEGIN LOOP --------------------------
for country, link in zip(countries, links) :
    try:
        #OPEN CHROME
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-ssl-errors=yes')
        options.add_argument('--ignore-certificate-errors')
        chrome = webdriver.Chrome(options=options)
        actions = ActionChains(chrome)
        wait = WebDriverWait(chrome, 10)

        #GO TO WEBREPORT PAGE
        chrome.get('https://webreportlink.com/Portal.ASPX') #goes to the home page
        time.sleep(sleepTimer)
        chrome.get(link)
        #OPEN THE PAYMENTS REQUIRING ACTION WEB REPORT
        chrome.get('https://webreportlink.com/Portal/menuReports.aspx')
        chrome.get('https://webreportlink.com/Portal/ReportSample.aspx')
        time.sleep(sleepTimer)
        #GET THE PAYMENTS REQUIRING ACTION EXCEL REPORT
        element = wait.until(EC.element_to_be_clickable((By.NAME, 'ExpToExcel')))
        chrome.find_element_by_name('ExpToExcel').click() #presses the export to excel button
        time.sleep(sleepTimer+8)
        chrome.quit()
        time.sleep(sleepTimer+3)
        #GENERATE THE PIVOT TABLE REPORT
        df = pd.read_html(r'C:\Users\alessio\Downloads\Report.xls', header=0)
        df = df[0]
        df.drop(df[df['Group Type'] == 'Finance'].index, inplace=True)
        df.drop(df[df['Group Type'] == 'Final Total'].index, inplace=True)
        df.drop_duplicates(subset ="Payment Id", inplace = True)
        
        def label_paydesc (row):
           if row['Pay Sts'] == 'A':
              return 'Status 1'
           if row['Pay Sts'] == 'B':
              return 'Status 2'
           if row['Pay Sts'] == 'C':
              return 'Status 3'
           if row['Pay Sts'] == 'D':
              return 'Status 4'
           if row['Pay Sts'] == 'E':
              return 'Status 5'
           if row['Pay Sts'] == 'F':
              return 'Status 6'
           if row['Pay Sts'] == 'G':
              return 'Status 7'
           if row['Pay Sts'] == 'H':
              return 'Status 8'
           if row['Pay Sts'] == 'I':
              return 'Status 9'
           if row['Pay Sts'] == 'J':
              return 'Status 10'
           return 'Other'
        #ADD NEW COLUMNS TO REPORT
        df['Pay Sts Desc'] = df.apply(lambda row: label_paydesc(row), axis=1)
        df['Year'] = pd.DatetimeIndex(df['Raised Date']).year #add year column to data
        df['Quarter'] = pd.DatetimeIndex(df['Raised Date']).quarter #add quarter column to data
        #CREATE THE PIVOT TABLE
        pivot = df.groupby(['LOB Description']).apply(lambda sub: sub.pivot_table(index = ['LOB Description', 'Year', 'Quarter'], columns=["Pay Sts Desc"], values="Claim Number", aggfunc='count', fill_value="", margins=True, margins_name='Subtotal'))
        pivot.index = pivot.index.droplevel(0)
        pivot.fillna('', inplace=True)
        pivot = pivot.filter(items=['Status 2', 'Status 3', 'Status 6', 'Status 9'])

        #SAVE TO EXCEL
        str2 = str(country)
        str3 = ".xlsx"
        path = str('C:/Users/alessio/Documents/2020')+str1+str(date)+str1+str2+str3
        writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
        df.to_excel(writer, sheet_name = 'Report', index=False)
        pivot.to_excel(writer, sheet_name = 'Pivot')
        writer.save()
        writer.close()
        
        os.remove('C:/Users/alessio/Downloads/Report.xls')
        loopCount = str(i)
        f = open(logName, 'a') # Open Log File
        f.write('\n' + loopCount + ' ' + country + ' ' + ' Report Run Successfully') #Log Success in File
        f.close() # Save and Close Log File
        print (loopCount + ' ' + country + ' Report Run Successfully')
        i = i + 1 # Increment i to count and track loops
        time.sleep(sleepTimer+1)
    except Exception as errException:
        os.remove('C:/Users/alessio/Downloads/Report.xls')
        loopCount = str(i) # Get Loops and Convert to string
        timeStamp = str(time.ctime()) # Get Current Time and Convert to String
        errMsg = str(errException) # Get Error Message and Convert into String
        f = open(logName, 'a') # Open Log File
        f.write('\n' + loopCount + ' ' + country + ' Errored! -' + errMsg + timeStamp) #Log Error in File
        f.close() # Save and Close Log File
        print(errException)
        print(i) # Print Loop Number 
        i = i + 1 # Increment i
        time.sleep(5)
    continue

# PRINT AND LOG COMPLETION NOTES -------
i = i - 1 #Because the loop always increments by 1, Last loop will increase increment by 1 more than there are actual items
loopCount = str(i)
timeStamp = str(time.ctime()) # Get Current Time
f = open(logName, 'a') # Open Log File
f.write('\n' + '0----------Script finished! at ' + timeStamp + '. Python tried to run ' + loopCount + ' reports') #Log Completion in File
f.close() # Save and Close Log File
print('0----------Script finished! at ' + timeStamp + '. Python tried to run ' + loopCount + ' reports. Please see '+ logName +' for any errors!')