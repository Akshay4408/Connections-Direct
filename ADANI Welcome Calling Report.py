#!/usr/bin/env python
# coding: utf-8

# In[21]:


import pandas as pd
import numpy as np


# In[22]:



#Export Data
export = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Welcome Calling\Client\August\Overall August Export.xlsx")
export = export[['call_date', 'phone_number_dialed', 'full_name', 'status_name', 'Disposition', 'BDD']]
export.head()


# In[23]:



#Client Data
client = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Welcome Calling\Client\August\Overall Aug Dump.xlsx")
client.head()


# In[24]:


export["status_name"].fillna("Ringing", inplace = True)


# In[25]:




export = export.sort_values("BDD", ascending=True)

export.head()



# In[26]:





export1 = export.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


export1.head()


# In[27]:


export1 = export1.rename(columns = {'phone_number_dialed' : 'Phone'})


# In[28]:




Result =  pd.merge(client ,export1, how = 'left', on = "Phone")
Result.head()


# In[29]:




Result["Disposition"].fillna("No Dial", inplace = True)
Result["status_name"].fillna("No Dial", inplace = True)



# In[30]:



StatusCount= pd.pivot_table(data = Result, index = ['Disposition'], values = 'status_name', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

StatusCount['Avg%'] = StatusCount['status_name'] / 1355
StatusCount['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in StatusCount['Avg%']], index = StatusCount.index)

StatusCount = StatusCount.rename(columns = {'status_name' : 'Count'})



StatusCount


# In[31]:


# Connect



connect = Result.loc[Result['Disposition'] == "Connect"]


Notconnect = Result.loc[Result['Disposition'] == "Not Connected"]


# In[32]:


connect1= pd.pivot_table(data = connect, index = ['Disposition', 'status_name'], values = 'Phone', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
connect1 = connect1.rename(columns = {'Phone': 'Count'})


connect1['Avg%'] = connect1['Count'] / 1321
connect1['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in connect1['Avg%']], index = connect1.index)


connect1


# In[33]:




Notconnect2= pd.pivot_table(data = Notconnect, index = ['Disposition', 'status_name'], values = 'Phone', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
Notconnect2 = Notconnect2.rename(columns = {'Phone': 'Count'})


Notconnect2['Avg%'] = Notconnect2['Count'] / 34



Notconnect2['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Notconnect2['Avg%']], index = Notconnect2.index)


Notconnect2


# In[34]:


connect.head()


# In[35]:


# Connect



Success = connect.loc[(connect['status_name'] == "Information Shared") ]


Success


# In[36]:



SuccessLead = pd.pivot_table(data = Success, index = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0)

SuccessLead['Avg%'] = SuccessLead['Disposition'] / 654
SuccessLead['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in SuccessLead['Avg%']], index = SuccessLead.index)

SuccessLead = SuccessLead.rename(columns = {'Disposition' : 'Total'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
SuccessLead


# In[37]:



States = pd.pivot_table(data = Result, index = ['State'],values = ['Disposition'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

States['Avg%'] = States['Disposition'] / 1355
States['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in States['Avg%']], index = States.index)

States = States.rename(columns = {'Disposition' : 'Count'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
States


# In[38]:



product = pd.pivot_table(data = Result, index = ['Product Name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

product['Avg%'] = product['Disposition'] / 1355
product['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in product['Avg%']], index = product.index)

product = product.rename(columns = {'Disposition' : 'Count'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
product


# In[39]:




writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\Welcome callingreport.xlsx", engine='xlsxwriter')


# In[40]:




connect1.to_excel(writer, sheet_name='connect Status')
Notconnect2.to_excel(writer, sheet_name='Notconnect status')
StatusCount.to_excel(writer, sheet_name='StatusCount')
export.to_excel(writer, sheet_name='export')
Result.to_excel(writer, sheet_name='Dump')
SuccessLead.to_excel(writer, sheet_name='SuccessLead')
product.to_excel(writer, sheet_name='product')
States.to_excel(writer, sheet_name='States')


writer.save()


# In[ ]:




