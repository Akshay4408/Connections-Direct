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


apr = apr[['USER NAME', 'CALLS','TALK', 'TOTAL', 'PAUSE', 'NONPAUSE','WAIT', 'TEA',  'LUNCH','BIO', 'TM' ]]
apr = apr.rename(columns = {'TOTAL' : 'LOGIN'})


# In[6]:


apr.head()


# In[7]:


from datetime import date
from datetime import timedelta
  


# In[8]:


apr['DATE'] =  date.today()
apr['DATE']  =  apr['DATE'] - timedelta(days = 1)
apr.head()


# In[9]:


apr = apr.rename(columns = {'LOGIN' : 'Total Login'})
apr.head()


# In[10]:


apr = apr[['DATE','USER NAME', 'CALLS','TALK', 'PAUSE', 'NONPAUSE','WAIT', 'TEA',  'LUNCH','BIO', 'TM', 'Total Login' ]]


# In[11]:


#CALL EXPORT
rw = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\APR\\call.xlsx")
rw = rw.rename(columns = {'full_name' : 'USER NAME'})
rw.head()


# In[12]:


agent = pd.read_excel(r"C:\Users\data anylitics\Desktop\APR\Agent List.xlsx")


# In[13]:


rw['Dates'] = pd.to_datetime(rw['call_date']).dt.date
rw['Time'] = pd.to_datetime(rw['call_date']).dt.time


# In[14]:


Login = rw.groupby(['USER NAME'])['Time'].min().reset_index(name='Login')
Logout = rw.groupby(['USER NAME'])['Time'].max().reset_index(name='Logout')


# In[15]:


APR = pd.merge(Login, Logout, how = 'inner', on= 'USER NAME')
APR


# In[16]:


APR = pd.merge(APR, apr, how = 'right', on= 'USER NAME')

APR.head()


# In[17]:


APR.info()


# In[18]:


#TL FILE EXPORT
TL = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\APR\\TL FILE.xlsx")

TL = TL.rename(columns = {'Name' : 'USER NAME'})

TL


# In[19]:


#APR = pd.merge(APR, TL, how = 'inner', on= 'USER NAME')

#APR = APR[['DATE', 'USER NAME','Team', 'TL','Login','Logout','TALK','TEA', 'BIO' , 'LUNCH','PAUSE','WAIT','Total Login']


# In[20]:


APR = APR[['DATE','USER NAME', 'CALLS','Login','Logout','TALK', 'PAUSE', 'NONPAUSE','WAIT', 'TEA',  'LUNCH','BIO', 'TM', 'Total Login' ]]


# In[21]:


#APR = pd.merge(APR, agent, how = 'inner', on= 'USER NAME')


# In[22]:


APR.info()


# In[23]:


APR.to_excel(r'C:\Users\data anylitics\Desktop\Akshay\\Report1.xlsx')


# In[ ]:




