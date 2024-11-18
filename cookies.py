from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
import time
import json
import pickle

service = Service(executable_path=r'c:\webdriver\msedgedriver.exe')
driver = webdriver.Edge(service=service)

driver.get("https://www.mercadolibre.com.co")

time.sleep(2)  

cookies = driver.get_cookies()

pickle.dump(cookies, open("cookies.pkl", "wb"))
driver.quit()

with open('cookies.json', 'w') as file:
    json.dump(cookies, file)

print(cookies)
