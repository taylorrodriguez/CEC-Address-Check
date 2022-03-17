from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
from openpyxl import Workbook, load_workbook

#Chrome Driver Path
PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

#Site used to get address data
driver.get("https://neighborhoodinfo.lacity.org/")

#Get addresses from Excel File using Pandas
file_loc = "Addresses Mejia 5Q-0 - cleaned.xlsx"
df = pd.read_excel(file_loc, index_col=None, na_values=['NA'], usecols="A")

#Create blank arrays (lists) for City and District
cityList = []
districtList = []

#Iterate through addresses, search, and retrieve data
for i, row in df.iterrows():
         time.sleep(2)
         search = driver.find_element(By.ID, 'ita-acsf-neighborhoodinfo-modal-field-address').clear()
         search = driver.find_element(By.ID, 'ita-acsf-neighborhoodinfo-modal-field-address')
         search.send_keys(row)
         search.send_keys(Keys.RETURN)
         time.sleep(2)
       


       #Decide if address is In City and In District according to whether Mayor and CD Official(ex. CD 13 Mitch O'Farrell) names are on the site. 

         if "Eric Garcetti" in driver.page_source:
            time.sleep(1)    
            cityList.append("Y")
            if "Mitch O'Farrell" in driver.page_source:
               time.sleep(1) 
               districtList.append("Y")
            else:
               time.sleep(1) 
               districtList.append("N")
         else:
            time.sleep(1)  
            cityList.append("N")
            time.sleep(1)
            districtList.append("N")
           
#Can check result in terminal 
print(cityList,districtList)
       

#Post to Excel

wb = load_workbook("Addresses Mejia 5Q-0 - cleaned.xlsx")
ws = wb.create_sheet("output")

for row in zip(cityList, districtList):
   ws.append(row)

wb.save("Addresses Mejia 5Q-0 - cleaned.xlsx")