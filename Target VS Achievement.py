#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:


# Export Data
rw = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\BAJAJ\\TARGET VS ACHIEVEMENT\\Report\\Dialy Sales_May'21.xlsx")


# In[3]:


# Client Data

Clinet = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\BAJAJ\\TARGET VS ACHIEVEMENT\\Report\\Client File.xlsx")


# In[4]:


rw.head()


# In[5]:


sales = rw.groupby(['Agent Name']).size().reset_index(name='Sales')


# In[6]:


premium = rw.groupby(['Agent Name'])['Premium Amount'].sum().reset_index(name='Net Premium')


# In[7]:


final = rw.groupby(['Agent Name'])['FINAL'].sum().reset_index(name='Final Premium')


# In[8]:


Result = pd.merge(sales, premium, how = 'inner', on = 'Agent Name')


# In[9]:


Result = pd.merge(Result, final, how = 'inner', on = 'Agent Name')


# In[10]:


Presentcount = rw.groupby(['Agent Name']).size().reset_index(name='Present Count')


# In[11]:


Presentcount


# In[12]:


Result = pd.merge(Result, Presentcount, how = 'inner', on = 'Agent Name')


# In[13]:


Result['Target'] = Result['Present Count'] * 0.65

Result['Target'] = Result['Target'].round()


# In[14]:


#Result['Achievement'] = Result['Sales'].divide(Result['Target'], fill_value=0)
Result['Achievement%'] = (Result['Sales'] / Result['Target']) * 100


Result['Achievement%'] = Result['Achievement%'].astype(int)

Result['Achievement%'] = Result['Achievement%'].astype(str) + '%'

#Result['Achievement'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Result['Achievement']], index = Result.index)



# In[15]:


Result = Result[['Agent Name', 'Present Count', 'Target','Sales', 'Net Premium', 'Final Premium', 'Achievement%']]


# In[16]:


Result


# In[17]:


Report = pd.merge(Clinet, Result, how = 'inner', on = 'Agent Name')


# In[18]:


Report = Report[['Emp Code', 'Agent Name', 'Designation', 'Language', 'LOB', 'TL Name', 'Present Count', 'Target', 'Sales', 'Net Premium', 'Final Premium', 'Achievement%', 'Tenure']]


# In[19]:


Report.to_excel(r'C:\\Users\\data anylitics\\Desktop\\BAJAJ\\TARGET VS ACHIEVEMENT\\Report\\Report.xlsx')


# In[ ]:




