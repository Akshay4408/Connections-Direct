#!/usr/bin/env python
# coding: utf-8

# In[21]:


import pandas as pd
import numpy as np


# In[22]:



#Client Data
client = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Loan Hindi\Sept\client.xlsx", parse_dates=['entry_date'])
client.head()


# In[23]:



#Export Data
export = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Loan Hindi\Sept\Overall Export of Sept.xlsx")
#export = export[['call_date', 'phone_number_dialed', 'full_name', 'status_name', 'Disposition', 'BDD']]
export.head()


# In[24]:


import datetime


client['entry_date'] = client['entry_date'].dt.strftime('%d %b, %Y')
client.head()


# In[25]:


client['entry_date'].value_counts()


# In[26]:


client = client.loc[(client['entry_date'] == '01 Sep, 2021') ]
client.head()
                    


# In[27]:


client = client.rename(columns = {'entry_date': 'Entry_Date'})


# In[28]:


export["status_name"].fillna("Ringing", inplace = True)
export.head()


# In[29]:




export = export.sort_values("BDD", ascending=True)

export


# In[30]:





export = export.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


export.head()


# In[31]:


client = client.rename(columns = {'phone_number': 'phone_number_dialed'})
export = export[['phone_number_dialed', 'full_name', 'status_name', 'Disposition']]
export.head()


# In[32]:


client1 = client[['lead_id', 'Entry_Date', 'user', 'phone_number_dialed', 'Lead_ID_1', 'Cust_Name', 'Postal_Code_1', 'Lead_Source', 'Product', 'State_1', 'Agent_Remark']]
client1.head()


# In[33]:


client1 = client1.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


# In[34]:




Result =  pd.merge(client1 ,export, how = 'inner', on = "phone_number_dialed")
Result.head()


# In[35]:



StatusCount= pd.pivot_table(data = Result, index = ['Entry_Date'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#StatusCount['Avg%'] = StatusCount['(Count, Total)'] / 155
#StatusCount['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in StatusCount['Avg%']], index = StatusCount.index)

#StatusCount = StatusCount.rename(columns = {'status_name' : 'Count'})



StatusCount


# In[36]:



Conn = pd.pivot_table(data = Result, index = ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

Conn['Avg%'] = Conn['status_name'] / 3
Conn['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Conn['Avg%']], index = Conn.index)

Conn = Conn.rename(columns = {'status_name' : 'Count'})



Conn


# In[37]:


# Connect



connect = Result.loc[Result['Disposition'] == "Connect"]


Notconnect = Result.loc[Result['Disposition'] == "Not Connected"]


# In[38]:



ConnStatus = pd.pivot_table(data = connect, index = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

ConnStatus['Avg%'] = ConnStatus['Disposition'] / 334
ConnStatus['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in ConnStatus['Avg%']], index = ConnStatus.index)

ConnStatus = ConnStatus.rename(columns = {'Disposition' : 'Count'})



ConnStatus


# In[39]:



NonConnStatus = pd.pivot_table(data = Notconnect, index = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

NonConnStatus['Avg%'] = NonConnStatus['Disposition'] / 3
NonConnStatus['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in NonConnStatus['Avg%']], index = NonConnStatus.index)

NonConnStatus = NonConnStatus.rename(columns = {'Disposition' : 'Count'})



NonConnStatus


# In[40]:




writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\Loan Lead Conectivity.xlsx", engine='xlsxwriter')


# In[41]:




NonConnStatus.to_excel(writer, sheet_name='NonConn')
ConnStatus.to_excel(writer, sheet_name='Connect')
Conn.to_excel(writer, sheet_name='Connctivity')
StatusCount.to_excel(writer, sheet_name='Datecount')
Result.to_excel(writer, sheet_name='Result')


writer.save()

