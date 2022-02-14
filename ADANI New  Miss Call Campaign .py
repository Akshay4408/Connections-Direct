#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:



#Export Data
export = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Miss Call\Overall Export Aug.xlsx")
export = export[['call_date', 'phone_number_dialed', 'full_name', 'status_name', 'Disposition', 'BDD']]
export.head()


# In[3]:



#Client Data
client = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Miss Call\Miss Call Data august.xlsx")
client.head()


# In[4]:


export["status_name"].fillna("Ringing", inplace = True)


# In[5]:


client = client.rename(columns = {'number' : 'phone_number_dialed'})
client.head()


# In[6]:




export = export.sort_values("BDD", ascending=True)

export.head()



# In[7]:





export = export.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


export.head()


# In[8]:




Result =  pd.merge(client ,export, how = 'left', on = "phone_number_dialed")
Result.head()


# In[9]:


Result = Result[['datetime', 'Month','phone_number_dialed', 'full_name','status_name', 'Disposition']]
Result.head()


# In[10]:


Result["status_name"].fillna("No Dial", inplace = True)
Result["Disposition"].fillna("No Dial", inplace = True)


# In[11]:



StatusCount= pd.pivot_table(data = Result, index = ['Disposition'], values = 'status_name', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

StatusCount['Avg%'] = StatusCount['status_name'] / 196
StatusCount['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in StatusCount['Avg%']], index = StatusCount.index)

StatusCount = StatusCount.rename(columns = {'status_name' : 'Count'})



StatusCount


# In[12]:


# Connect



connect = Result.loc[Result['Disposition'] == "Connect"]


Notconnect = Result.loc[Result['Disposition'] == "Not Connected"]


# In[13]:


connect1= pd.pivot_table(data = connect, index = ['Disposition', 'status_name'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
connect1 = connect1.rename(columns = {'phone_number_dialed': 'Count'})


connect1['Avg%'] = connect1['Count'] / 137
connect1['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in connect1['Avg%']], index = connect1.index)


connect1


# In[14]:




Notconnect2= pd.pivot_table(data = Notconnect, index = ['Disposition', 'status_name'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
Notconnect2 = Notconnect2.rename(columns = {'phone_number_dialed': 'Count'})


Notconnect2['Avg%'] = Notconnect2['Count'] / 46
Notconnect2['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Notconnect2['Avg%']], index = Notconnect2.index)


Notconnect2


# In[15]:




writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\New MissCall Report.xlsx", engine='xlsxwriter')


# In[16]:




connect1.to_excel(writer, sheet_name='connect Status')
Notconnect2.to_excel(writer, sheet_name='Notconnect status')
StatusCount.to_excel(writer, sheet_name='StatusCount')
export.to_excel(writer, sheet_name='export')
Result.to_excel(writer, sheet_name='Dump')


writer.save()


# In[ ]:




