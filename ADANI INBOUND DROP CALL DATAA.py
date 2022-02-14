#!/usr/bin/env python
# coding: utf-8

# In[7]:


import pandas as pd
import numpy as np


# In[8]:


rw = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Inbound\Inbound\export1.xlsx")
rw = rw.loc[rw['Direction'] == "inbound"]

rw.head()


# In[9]:


rw = rw.loc[(rw['Status'] == "missed-call") | (rw['Status'] == "call-attempt")]


# In[10]:


rw = rw.drop_duplicates(
  subset = ['From'],
  keep = 'first').reset_index(drop = True)


# In[11]:


rw


# In[12]:


rw.to_excel(r'C:\Users\data anylitics\Desktop\Akshay\Dropcall Data.xlsx')


# In[ ]:




