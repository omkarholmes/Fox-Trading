#!/usr/bin/env python
# coding: utf-8

# In[14]:


#pip install requests


# In[11]:


#pip install pyexcel-xlsx


# In[16]:


#Imports required for the task
import requests
import json
import pyexcel_xlsx as pe
import threading
from datetime import datetime


# In[17]:


# Name of File we need to work on
fileName = "Task1.xlsx"


# In[18]:


# Gets us the OrderedDict object from the file 
def getFile(fileName):
    return pe.get_data(afile=fileName)


# In[19]:


# API Key needed for to access the Weather API and retrive the response object
apiKey = "3242278dab236a765d0bf76f2349c875"
def fetchWeatherForCity(cityName):
    url ="http://api.openweathermap.org/data/2.5/weather?q="+cityName+"&appid="+apiKey
    response = requests.get(url)
    response.json()
    return response.json()["main"]


# In[20]:


# Simple function for conversion from Kelvin to Celsius
def convertToCelsius(kel):
    return (kel-273.15)


# In[21]:


# Simple function for conversion from Kelvin to Fahrenheit
def convertToFahrenheit(kel):
    cel = convertToCelsius(kel)
    return cel*9/5 + 32


# In[22]:


#Method to get the City name from the adjacent sheet mapped with City Token
def getCityName(cityToken,sheet):
    for city in sheet["City Tokens"][1:]:
        if city[1] == cityToken:
            return city[0]
    print("Error : No City Found with given token")


# In[23]:


# A Function to regularly update the excel sheet with the updated weather details of the specified cities in required units
def updateSheet(sheet):
    for city in sheet["Weather"][1:]:
        if(city[4]==1):
            cityName = getCityName(city[0],sheet)
            temp = fetchWeatherForCity(cityName)
            if(city[3]=="F"):
                city[1]=round(convertToFahrenheit(temp["temp"]),1)
            elif (city[3]=="C"):
                city[1]=round(convertToCelsius(temp["temp"]),1)
            city[2]=temp["humidity"]       
    saveSheet(sheet)


# In[24]:


#A function to save the excel sheet after the regular update to the columns
def saveSheet(excel):
    excel={
        "Weather":excel["Weather"],
        "City Tokens":excel["City Tokens"]
    }
    pe.save_data(fileName, excel)


# In[ ]:


def printAck():
    now = datetime.now()
    current_time = now.strftime("%d/%b/%Y %I:%M:%S %p")
    print(fileName+" updated on "+current_time) 
    # x=input("To stop the script enter 'x'")
    # if(x=="x"):
    #     exit()


# In[26]:


# A Function that calls the updateSheet at an interval of 30 seconds
t=None
def intervalScript():
    global t
    t= threading.Timer(30.0, intervalScript)
    t.start()
    updateSheet(getFile(fileName))
    printAck()
# A accessory method to exit the continuous cycle to update the excel sheet   
def exit():
    t.cancel()


# In[27]:


intervalScript()





