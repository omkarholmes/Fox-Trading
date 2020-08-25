#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Imports required for the task
import pyexcel_xlsx as pe
import pandas as pd
import datetime as dt


# In[2]:


# Name of File we need to work on
fileName = "NIFTY25JUN2010000PE.xlsx"
# Name of the Sheet from which we generate report
sheet="NIFTY25JUN2010000PE"


# In[3]:


# Gets us the OrderedDict object from the file 
def getFile(fileName):
    return pe.get_data(afile=fileName)


# In[4]:


# Returns the Report entries for the exit criteria
def createReport(df):
    report=[list(df.columns)]
    for i in range(1,(df.shape[0])):
        c_row = df.iloc[i]
        p_row = df.iloc[i-1]
        if(c_row[6] < p_row[6]):
            if(p_row[5] < c_row[5] or (c_row[2]-p_row[2]).days > 0):
                report.append(list(c_row))
    return report


# In[5]:


# Processing the sheet for the resampling of the entries from a granularity of 1 minute to 15 minutes.
def preProcess(excel):
    df = pd.DataFrame(excel[1:])
    df.columns=excel[0]
    df.insert(1,
          "Timestamp",
          df["Date"].map(lambda x : x.strftime("%Y-%m-%d"))+" "+df["Time"].map(lambda x : x.strftime("%H:%M:%S")))
    df.drop(["Date","Time"],axis=1)
    df["Timestamp"]=df["Timestamp"].map(lambda x : dt.datetime.strptime(x, "%Y-%m-%d %H:%M:%S"))
    df=(df.set_index('Timestamp')
        .resample('15T').first()
        .reset_index()
        .reindex(columns=df.columns)
       )
    df.dropna(inplace=True)
    df.index=range(0,df.shape[0])
    return df


# In[12]:


# To start the script for generating report based on the startegy mentioned and computing profit/loss
def run():
    excel = getFile(fileName)
    report=createReport(preProcess(excel[sheet]))
    dr=pd.DataFrame(report[1:])
    dr.columns=report[0]
    dr.insert(dr.columns.shape[0],"Profit/Loss",(dr[dr.columns[4]]-dr[dr.columns[7]]))
    dr.insert(dr.columns.shape[0],"Profit/Loss Volume", dr[dr.columns[9]]*dr[dr.columns[8]])
    saveSheet(excel,convertDFtoOrdDict(dr))
    print("Report Generated for the file "+fileName)


# In[7]:


#A method convert the generated report to a sheet
def convertDFtoOrdDict(df):
    tep=[list(df.columns)]
    for row in df.iterrows():
        tep.append(list(row[1]))
    return tep


# In[10]:


#A function to save sheet after creating/updating the sheets.
def saveSheet(excel,report):
    excel = {
        sheet:excel[sheet],
        "Report":report
    }
    pe.save_data(fileName,excel)


# In[11]:


run()


# In[ ]:





# In[ ]:




