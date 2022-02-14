#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:


rw = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Inbound\Inbound\export1.xlsx")

rw = rw.loc[rw['Direction'] == "inbound"]

rw.head()


# In[3]:


remark = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Inbound\Remarks.xlsx")
remark.head()


# In[4]:


from datetime import date
from datetime import timedelta
  


# In[5]:


rw['datehour'] = rw['StartTime'].dt.hour
rw.head()


# In[6]:


def test1(x):
    if x == 8:
        return '8 AM to 9 AM'
    if x == 9:
        return '9 AM to 10 AM'
    if x == 10:
        return '10 AM to 11 AM'
    if x == 11:
        return '11 AM to 12 PM'
    if x == 12:
        return '12 PM to 1 PM'
    if x == 13:
        return '1 PM to 2 PM'
    if x == 14:
        return '2 PM to 3 PM'
    if x == 15:
        return '3 PM to 4 PM'
    if x == 16:
        return '4 PM to 5 PM'
    if x == 17:
        return '5 PM to 6 PM'
    if x == 18:
        return '6 PM to 7 PM'
    if x == 19:
        return '7 PM to 8 PM'
    if x == 20:
        return '8 PM to 9 PM'
    if x == 21:
        return '9 PM to 10 PM'
    if x == 22:
        return '10 PM to 11 PM'
    else:
        return '11 PM to 12 AM'

     


    










# In[7]:


rw['datehour'] = rw['datehour'].apply(lambda x: test1(x))


# In[8]:


rw.head()


# In[9]:


rw["ToName"].fillna("Drop", inplace = True)


# In[10]:


def flag_df(df):

    if (df['ToName'] == "Drop") and (df['Status'] == "missed-call"):
        return 'Drop Calls'
    
    else:
        return 'Answered Calls'

rw['Flag'] = rw.apply(flag_df, axis = 1)


# In[11]:


rw['Flag'].value_counts()


# In[12]:


rw.head()


# In[13]:


AHT = pd.pivot_table(data = rw, index = ['datehour'], columns=['Flag'],values = 'StartTime', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
#AHT = AHT.rename(columns = {'StartTime' : 'Count'})

AHT = AHT.rename(columns = {'Total' : 'Offered Calls'})

AHT


# In[14]:



AHT = AHT[['Offered Calls', 'Answered Calls', 'Drop Calls']]
AHT


# In[15]:


Result =  pd.merge(AHT ,remark, how = 'left', on = "datehour")
Result


# In[16]:


Result["Remark"].fillna("-", inplace = True)
Result


# In[17]:


def test1(x):
    if x == '8 AM to 9 AM':
        return 1
    if x == '9 AM to 10 AM':
        return 2
    if x == '10 AM to 11 AM':
        return 3
    if x == '11 AM to 12 PM':
        return 4
    if x == '12 PM to 1 PM':
        return 5
    if x == '1 PM to 2 PM':
        return 6
    if x == '2 PM to 3 PM':
        return 7
    if x == '3 PM to 4 PM':
        return 8
    if x == '4 PM to 5 PM':
        return 9
    if x == '5 PM to 6 PM':
        return 10
    if x == '6 PM to 7 PM':
        return 11
    if x == '7 PM to 8 PM':
        return 12
    if x == '8 PM to 9 PM':
        return 13
    if x == '9 PM to 10 PM':
        return 14
    if x == '10 PM to 11 PM':
        return 15
    else:
        return 16

     


    










# In[18]:


Result['Numm'] = Result['datehour'].apply(lambda x: test1(x))


# In[19]:


Result


# In[20]:


Result = Result.sort_values("Numm", ascending=True)
Result


# In[21]:


Result = Result[['datehour', 'Offered Calls', 'Answered Calls', 'Drop Calls', 'Remark']]
Result = Result.rename(columns = {'datehour' : 'Time'})

Result


# In[ ]:





# In[22]:


rw = rw.rename(columns = {'ToName' : 'Agent Name'})


agent = pd.pivot_table(data = rw, index = ['Agent Name'], values = 'Status', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

agent = agent.rename(columns = {'Direction' : 'Count'})
#agent = agent[['Answered Calls']]
agent = agent.rename(columns = {'Answered Calls' : 'Total'})

agent


# In[23]:


service =  pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Inbound\Inbound\service.xlsx" , parse_dates=['Timestamp'])


service.head()


# In[24]:


import datetime


service['Timestamp'] = service['Timestamp'].dt.strftime('%d %b, %Y')
service.head()


# In[25]:


service['Timestamp'].value_counts()


# In[26]:


service = service.loc[(service['Timestamp'] == "28 Aug, 2021")]
service.head()


# In[ ]:





# In[27]:


status = pd.pivot_table(data = service, index = ['Status'], values = 'Timestamp', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
status = status.rename(columns = {'Timestamp' : 'Count'})

status


# In[28]:




writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\ADANI Inboundreport.xlsx", engine='xlsxwriter')


# In[29]:





agent.to_excel(writer, sheet_name='agentcount')
status.to_excel(writer, sheet_name='statuscount')
Result.to_excel(writer, sheet_name='Report')


writer.save()

