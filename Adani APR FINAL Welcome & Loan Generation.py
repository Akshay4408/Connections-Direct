#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:


apr1 = pd.read_excel(r"C:\Users\data anylitics\Desktop\APR\one.xlsx")
apr1 = apr1.iloc[3:]
#apr1.apr1[:-1,:]
apr1 = apr1.rename(columns=apr1.iloc[0]).drop(apr1.index[0])
apr1 = apr1.iloc[:-1 , :]

apr1.head()


# In[3]:


# APR
apr2 = pd.read_excel(r"C:\Users\data anylitics\Desktop\APR\two.xlsx")

apr2 = apr2.iloc[3:]
#apr2.iloc[:-1,:]

apr2 = apr2.rename(columns=apr2.iloc[0]).drop(apr2.index[0])

apr2 = apr2.iloc[:-1 , :]

apr2.head()


# In[4]:


#CALL EXPORT
rw = pd.read_excel(r"C:\Users\data anylitics\Desktop\APR\call.xlsx")


# In[5]:


#Agent Name
agent = pd.read_excel(r"C:\Users\data anylitics\Desktop\APR\adani agent.xlsx")


# In[6]:


apr1 = apr1[['USER NAME', 'TALK']]
#apr1 = apr1.rename(columns = {'TOTAL' : 'LOGIN'})
apr1.head()


# In[7]:


apr2 = apr2[['USER NAME', 'TOTAL', 'LNCH', 'TEA']]
apr2 = apr2.rename(columns = {'TOTAL' : 'LOGIN'})
apr2.head()


# In[8]:


def test(x):
    if x == 'Agent Not Available':
        return 'Not Connect'
    if x == 'Busy Auto':
        return 'Not Connect'
    if x == 'No Answer AutoDial':
        return 'Not Connect'
    if x == 'Out Of Network':
        return 'Not Connect'
    if x == 'Ringing':
        return 'Not Connect'
    if x == 'Short Hang UP':
        return 'Not Connect'
    if x == 'Lead Being Called':
        return 'Not Connect'
    if x == 'Switch Off':
        return 'Not Connect'
    if x == 'Busy':
        return 'Not Connect'
    if x == 'Promise To Pay NC':
        return 'Not Connect'
    if x == 'Promise To Pay NC':
        return 'Not Connect'
    if x == 'Refer To Field':
        return 'Connect'


    else:
        return 'Connect'




# In[9]:


rw['Dispo'] = rw['status_name'].apply(lambda x: test(x))

rw = rw.rename(columns = {'full_name' : 'USER NAME'})
rw.head()


# In[10]:


rw =pd.pivot_table(data= rw, index = ['USER NAME'], columns = ['Dispo'],values = 'phone_number_dialed', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')

rw.head()


# In[11]:


Result = pd.merge(apr1, apr2, how = 'inner', on = 'USER NAME')
Result = pd.merge(Result, rw, how = 'inner', on = 'USER NAME')

Result


# In[12]:


from datetime import date
from datetime import timedelta
  


# In[13]:


Result['DATE'] =  date.today()
#Result['DATE']  =Result['DATE'].dt.strftime("%m/%d/%y")
Result['DATE']  =  Result['DATE'] - timedelta(days = 1)


# In[14]:


Result = Result[['DATE', 'USER NAME', 'TALK', 'LNCH', 'TEA', 'LOGIN', 'Connect', 'Not Connect', 'Total']]


# In[15]:


Result['Lhour'] = pd.to_datetime(Result['LNCH'], format='%H:%M:%S').dt.hour
Result['Lminutes'] = pd.to_datetime(Result['LNCH'], format='%H:%M:%S').dt.minute
Result['Lsecond'] = pd.to_datetime(Result['LNCH'], format='%H:%M:%S').dt.second


Result['Lhour'] = Result['Lhour'] * 3600
Result['Lminutes'] = Result['Lminutes'] * 60

Result['Llunch']  = Result['Lhour'] + Result['Lminutes'] + Result['Lsecond']


Result['Thour'] = pd.to_datetime(Result['TEA'], format='%H:%M:%S').dt.hour
Result['Tminutes'] = pd.to_datetime(Result['TEA'], format='%H:%M:%S').dt.minute
Result['Tsecond'] = pd.to_datetime(Result['TEA'], format='%H:%M:%S').dt.second



Result['Thour']= Result['Thour'] * 3600
Result['Tminutes'] = Result['Tminutes'] * 60

Result['Ttea']  = Result['Thour'] +Result['Tminutes'] + Result['Tsecond']


# In[16]:


#Result['BMinutes'] = Result['Lminutes'] + Result['Tminutes'] 

Result['Break'] = Result['Llunch'] + Result['Ttea'] 



#operator for time

import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')
Result['BREAK'] = pd.to_datetime(Result['Break'], unit='s').map(fmt)


# In[17]:


Result = Result[['DATE', 'USER NAME', 'TALK' , "BREAK", 'LOGIN',"Connect", "Not Connect", "Total"]]


# In[18]:


agent = agent.rename(columns = {'Agent Name' : 'USER NAME'})


# In[19]:


Result = pd.merge(Result, agent, how = 'inner', on = 'USER NAME')
Result


# In[20]:


Result.to_excel(r'C:\Users\data anylitics\Desktop\Akshay\adaniwelcome.xlsx')


# In[ ]:




