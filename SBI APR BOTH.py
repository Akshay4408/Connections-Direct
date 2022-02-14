#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:


# APR one
apr1 = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\APR\\one.xlsx")


apr1 = apr1.iloc[3:]
#apr1.iloc[:-1,:]
apr1 = apr1.iloc[:-1 , :]

apr1 = apr1.rename(columns=apr1.iloc[0]).drop(apr1.index[0])



# In[3]:


# APR two
apr2 = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\APR\\two.xlsx")

apr2 = apr2.iloc[3:]
#apr2.iloc[:-1,:]

apr2 = apr2.rename(columns=apr2.iloc[0]).drop(apr2.index[0])
apr2 = apr2.iloc[:-1 , :]


# In[4]:


apr = pd.merge(apr1, apr2, how = 'inner', on = 'USER NAME')
apr = apr.rename(columns = {'PAUSE_x' : 'PAUSE'})
apr.head()


# In[5]:


#CALL EXPORT
rw = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\APR\\call.xlsx")
rw = rw.rename(columns = {'full_name' : 'USER NAME'})
rw.head()


# In[6]:


#TL FILE EXPORT
TL = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\APR\\TL FILE.xlsx")

TL = TL.rename(columns = {'Agent' : 'USER NAME'})

TL


# In[7]:


apr.head()


# In[8]:


apr = apr[['USER NAME', 'CALLS','TALK', 'TOTAL', 'PAUSE', 'NONPAUSE','WAIT', 'TEA',  'LUNCH','BIO', 'TM' ]]
apr = apr.rename(columns = {'TOTAL' : 'LOGIN'})


# In[9]:


apr.head()


# In[10]:


from datetime import date
from datetime import timedelta
  


# In[11]:


apr['DATE'] =  date.today()
apr['DATE']  =  apr['DATE'] - timedelta(days = 1)
apr.head()


# In[12]:


apr = apr.rename(columns = {'LOGIN' : 'Total Login'})
apr.head()


# In[13]:


apr = apr[['DATE', 'USER NAME', 'CALLS','TALK', 'TEA', 'BIO','LUNCH', 'PAUSE','NONPAUSE','WAIT','TM','Total Login']]
apr.head()


# In[14]:


rw['Dates'] = pd.to_datetime(rw['call_date']).dt.date
rw['Time'] = pd.to_datetime(rw['call_date']).dt.time


# In[15]:


Login = rw.groupby(['USER NAME'])['Time'].min().reset_index(name='Login')
Logout = rw.groupby(['USER NAME'])['Time'].max().reset_index(name='Logout')


# In[16]:


APR = pd.merge(Login, Logout, how = 'inner', on= 'USER NAME')
APR


# In[17]:


APR = pd.merge(APR, apr, how = 'right', on= 'USER NAME')
APR.head()


# In[18]:


APR = pd.merge(APR, TL, how = 'inner', on= 'USER NAME')

#APR = APR.drop('Logout', axis = 1)
APR.head()

#APR = APR[['DATE', 'USER NAME','Team', 'TL','Login','Logout','TALK','TEA', 'BIO' , 'LUNCH','PAUSE','WAIT','Total Login']


# In[19]:


#First Login
#APR['Lhours'] = pd.to_datetime(APR['Login'], format='%H:%M:%S').dt.hour

#APR['Lminutes'] = pd.to_datetime(APR['Login'], format='%H:%M:%S').dt.minute




#APR['Lhours'] = APR['Lhours'] * 60



#APR['Ltime'] = APR['Lhours'] + APR['Lminutes'] 

#Total Login
#APR['Thours'] = pd.to_datetime(APR['Total Login'], format='%H:%M:%S').dt.hour

#APR['Tminutes'] = pd.to_datetime(APR['Total Login'], format='%H:%M:%S').dt.minute




#APR['Thours'] = APR['Thours'] * 60



#APR['Ttime'] = APR['Thours'] + APR['Tminutes'] 

#Logout

#APR['Logout'] = APR['Ltime'] + APR['Ttime']


#operator for time

#import operator
#fmt = operator.methodcaller('strftime', '%H:%M:%S')
#APR['Logout Time'] = pd.to_datetime(APR['Logout'], unit='m').map(fmt)



#Result['Ptime'] = Result['Ptime'] - Result['BMinutes'] - Result['TMT']
#Result['TLtime'] = Result['TLtime'].astype(float) 

#Result['TLtime'] = Result['TLtime'] 


# In[20]:


APR.head()


# In[21]:


APR = APR[['DATE', 'USER NAME', 'TL/ATL','Login','Logout','CALLS','TALK','TEA', 'BIO' , 'LUNCH','PAUSE','NONPAUSE','WAIT','TM','Total Login']]
APR.head()


# In[22]:


APR.to_excel(r'C:\Users\data anylitics\Desktop\Akshay\\aprReport.xlsx')


# In[ ]:




