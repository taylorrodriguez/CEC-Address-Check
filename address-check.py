from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
from openpyxl import load_workbook

#Chrome Driver Path
PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

#Site used to get address data
driver.get("https://neighborhoodinfo.lacity.org/")

#Get addresses from Excel File using Pandas
file_loc = "Pynoos2 - cleaned.xlsx"
df = pd.read_excel(file_loc, index_col=None, na_values=['NA'], usecols="A")

#Create blank arrays (lists) for City and District
cityList = []
districtList = []
districtNumberList = []

#Iterate through addresses, search, and retrieve data
for i, row in df.iterrows():
         print(cityList,districtList)
         time.sleep(2)
         search = driver.find_element(By.ID, 'ita-acsf-neighborhoodinfo-modal-field-address').clear()
         search = driver.find_element(By.ID, 'ita-acsf-neighborhoodinfo-modal-field-address')
         search.send_keys(row)
         search.send_keys(Keys.RETURN)
         time.sleep(2)
       
       #Decide if address is In City and In District according to whether Mayor and CD Official(ex. CD 13 Mitch O'Farrell) names are on the site. 

         if "Eric Garcetti" in driver.page_source:
            time.sleep(1)

            #Find HTML element where the District information is located, by XPATH
            districtNumber = driver.find_element(By.XPATH,'/html/body/div[2]/div/div/div[3]/div[2]/div/main/section/div[2]/div/div/div[3]/table[1]/tr[5]/td[2]')
            
            #Get Text from element, example: "District 13: Mitch O'Farrell", and convert to "CD 13"
            districtNumberList.append("CD " + ((districtNumber.text).split(" ")[1])[:-1])
            
            print(districtNumberList)
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
            districtNumberList.append("None")
           
#Can check result in terminal 
print(cityList,districtList, districtNumberList)
       

#Post to Excel

wb = load_workbook(file_loc)
ws = wb.create_sheet("output")

for row in zip(cityList, districtList, districtNumberList):
   ws.append(row)

wb.save(file_loc)