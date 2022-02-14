#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:



rw = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\BAJAJ\\Penetration Report\\New Data\\Export.xlsx")


# In[3]:


client = pd.read_excel(r"C:\\Users\\data anylitics\\Desktop\\BAJAJ\\Penetration Report\\New Data\\Row Outbound.xlsx")


# In[4]:


rw.info()


# In[5]:


a = rw.value_counts(['phone_number']) 

x = pd.DataFrame(a)

x = x.rename(columns = {0 : 'Attempt'})


# In[6]:


rw = rw[['call_date', 'phone_number_dialed', 'full_name', 'status_name', 'phone_number']]


# In[7]:


rw.head()


# In[8]:


def test(x):
    
    if x == "Activated another policy":
        return 5
    if x == "DO NOT CALL":
        return 6
    if x == "No EMI Option":
        return 13
    if x == "Wrong Number":
        return 5
    if x == "Financial Problem":
        return 12
    if x == "ALREADY BOUGHT":
        return 3
    if x == "No Sales Opportunity":
        return 13
    if x == "Inquiry for personal loan":
        return 12
    if x == "Test Call":
        return 18
    if x == "pre-existing":
        return 11
    if x == "Duplicate Data":
        return 7
    if x == "Test Data":
        return 17
    if x == "Renewal":
        return 10
    if x == "Expired":
        return 8
    if x == "Not Interested":
        return 6
    if x == "Done With Competition":
        return 9
    if x == "Follow Up":
        return 4
    if x == "Voice Issue":
        return 20
    if x == "Call Back":
        return 4
    if x == "RPC Not Available":
        return 10
    if x == "Payment Done":
        return 1
    if x == "Renewal Done":
        return 2
    if x == "Renewal Done":
        return 2
    if x == "Language Barrier":
        return 7
    if x == "High Premium":
        return 19
    else:
        return 21



  


# In[9]:


rw['ST'] = rw['status_name']


# In[10]:


rw['ST'] = rw['ST'].apply(lambda x: test(x))


# In[11]:


rw = rw.sort_values("ST", ascending=True)

rw.head()


# In[12]:


rw = rw.drop_duplicates(
  subset = ['phone_number'],
  keep = 'first').reset_index(drop = True)


# In[13]:


rw.head()


# In[14]:


def test1(x):
    if x == 'Follow Up':
        return 'Connect'
    if x == 'Voice Issue':
        return 'Connect'
    if x == 'Call Back':
        return 'Connect'
    if x == 'RPC Not Available':
        return 'Connect'
    if x == 'Language Barrier':
        return 'Connect'
    if x == 'Payment Done':
        return 'Connect'
    if x == 'Renewal Done':
        return 'Connect'
    if x == 'High Premium':
        return 'Connect'
    if x == 'Not Interested':
        return 'Connect'
    if x == 'Done With Competition':
        return 'Connect'
    if x == 'Wrong Number':
        return 'Connect'
    if x == 'Renewal':
        return 'Connect'
    if x == 'pre-existing':
        return 'Connect'
    if x == 'No Sales Opportunity':
        return 'Connect'
    if x == 'No EMI Option':
        return 'Connect'
    if x == 'Inquiry for personal loan':
        return 'Connect'
    if x == 'Financial Problem':
        return 'Connect'
    if x == 'Expired':
        return 'Connect'
    if x == 'Duplicate Data':
        return 'Connect'
    if x == 'DO NOT CALL':
        return 'Connect'
    if x == 'ALREADY BOUGHT':
        return 'Connect'
    if x == 'Activated another policy':
        return 'Connect'
  
  
    else:
        return 'Not Connect'
  
  
  


  
  





  


# In[15]:


rw['Dispo'] = rw['status_name'].apply(lambda x: test1(x))
client = client.rename(columns = {'mobilenumber' : 'phone_number'})


# In[16]:


rw.head()


# In[17]:


# Connect
Connect = rw.loc[rw['Dispo'] == "Connect"]


# In[18]:


# Not Connect

NotConnected = rw.loc[rw['Dispo'] == "Not Connect"]


# In[19]:


# Not Lead
NonLead = rw.loc[rw['Dispo'] == "Non Lead"]


# In[20]:


Result =  pd.merge(client, NonLead, on = 'phone_number', how = 'inner')


# In[21]:


Result.head()


# In[22]:


Result1 =  pd.merge(client, Connect, on = 'phone_number', how = 'inner')


# In[23]:


Result2 =  pd.merge(client, NotConnected, on = 'phone_number', how = 'inner')


# In[24]:


output = Result.append(Result1)


# In[25]:


output = output.append(Result2)


# In[26]:


output.dropna(subset = ["status_name"], inplace=True)


# In[27]:


output = pd.merge(output, x , on = 'phone_number', how = 'inner')


# In[28]:


output.head()


# In[29]:


def test2(x):
    
    
    if x == "Connect":
        return 1
    
  
    else:
        return 2


# In[30]:


output['Countt'] = output['Dispo']


# In[31]:


output['Countt'] = output['Countt'].apply(lambda x: test2(x))


# In[32]:


output = output.sort_values("Countt", ascending=True)

output.head()


# In[33]:


output = output.drop_duplicates(
  subset = ['phone_number'],
  keep = 'first').reset_index(drop = True)


# In[34]:


output.head()


# In[35]:


Connectvity = pd.pivot_table(data = output, index = ['Dispo'], values = 'status_name', aggfunc = np.size, margins = True, margins_name = 'Total' )


# In[36]:


Connectvity


# In[37]:


Connectvity['Avg%'] = (Connectvity['status_name'] / 
                  49073) * 100

Connectvity['Avg%'] = Connectvity['Avg%'].astype(int)

Connectvity['Avg%'] = Connectvity['Avg%'].astype(str) + '%'

Connectvity


# In[38]:


#NOT CONNECT

NotConnect = output.loc[output['Dispo'] == "Not Connect"]
NotConnect.head()


NotConnect = pd.pivot_table(data = NotConnect, index = ['status_name'], values = 'Dispo', aggfunc = np.size, margins = True, margins_name = 'Total' )
NotConnect


NotConnect['Avg%'] = (NotConnect['Dispo'] / 
                  18625) * 100

NotConnect['Avg%'] = NotConnect['Avg%'].astype(int)

NotConnect['Avg%'] = NotConnect['Avg%'].astype(str) + '%'


NotConnect


# In[39]:


#CONNECT

connect = output.loc[output['Dispo'] == "Connect"]
connect = pd.pivot_table(data = connect, index = ['status_name'], values = 'Dispo', aggfunc = np.size, margins = True, margins_name = 'Total' )

connect

connect['Avg%'] = (connect['Dispo'] / 
                  30448) * 100

connect['Avg%'] = connect['Avg%'].astype(int)

connect['Avg%'] = connect['Avg%'].astype(str) + '%'

connect


# In[40]:


agent = pd.pivot_table(data = output, index = ['full_name'], columns = ['Dispo'], values = 'status_name', aggfunc = np.size, fill_value= 0,margins = True, margins_name= 'Total')

agent


# In[41]:


get_ipython().system(' pip install xlsxwriter')


# In[42]:


import xlsxwriter


# In[43]:


overall = pd.pivot_table(data = output, index = ['status_name'], values = 'Dispo', aggfunc = np.size, fill_value= 0,margins = True, margins_name= 'Total')

overall1 = pd.pivot_table(data = output, index = ['status_name'], values = 'Attempt', aggfunc = np.sum, fill_value= 0,margins = True, margins_name= 'Total')

overall = pd.merge(overall, overall1 , on = 'status_name', how = 'inner')

overall = overall.rename(columns = {'Dispo' : 'Count'})


# In[44]:


overall


# In[45]:


best = overall

best = best.rename(columns = {'Count' : 'Unique', 'Attempt' : 'Duplicate'})

best['Total Dial'] = best['Unique'] + best['Duplicate']


# In[46]:


# Unique and Duplicate
best


# In[47]:


Report1 = best

Report1['Attempt'] = Report1['Total Dial'] / Report1['Unique'] 

Report1['Pentration'] = Report1['Attempt'] / Report1['Unique'] 

Report1['Contribution'] = Report1['Unique'] / 49073 


# In[48]:


Report1


# In[49]:


Report1.info()


# In[50]:


Report1['Contribution'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Report1['Contribution']], index = Report1.index)

Report1['Pentration'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Report1['Pentration']], index = Report1.index)


# In[51]:


# Pentration
Report1


# In[52]:


# Connectivity
Report2 = pd.pivot_table(data = output, index = ['Dispo'], values = 'status_name', aggfunc = np.size, fill_value= 0,margins = True, margins_name= 'Total')

Report22 = pd.pivot_table(data = output, index = ['Dispo'], values = 'Attempt', aggfunc = np.sum, fill_value= 0,margins = True, margins_name= 'Total')

Report2 = pd.merge(Report2, Report22 , on = 'Dispo', how = 'inner')

Report2 = Report2.rename(columns = {'status_name' : 'Unique', 'Attempt' : 'Duplicate'})


# In[53]:


Report2


# In[54]:


# Connectivity
Report3 = pd.pivot_table(data = output, index = ['Dispo', 'Language'], values = 'status_name', aggfunc = np.size, fill_value= 0,margins = True, margins_name= 'Total')

Report33 = pd.pivot_table(data = output, index = ['Dispo', 'Language'], values = 'Attempt', aggfunc = np.sum, fill_value= 0,margins = True, margins_name= 'Total')

Report3 = pd.merge(Report3, Report33 , on = ['Dispo', 'Language'], how = 'inner')

Report3= Report3.rename(columns = {'status_name' : 'Unique', 'Attempt' : 'Duplicate'})


# In[55]:


writer = pd.ExcelWriter(r"C:\\Users\\data anylitics\\Desktop\\Akshay\\pentration.xlsx", engine='xlsxwriter')


# In[56]:


#Connectvity.to_excel(writer, sheet_name='Connectvity')
#best.to_excel(writer, sheet_name='best')
Report1.to_excel(writer, sheet_name='Report1')
Report2.to_excel(writer, sheet_name='Report2')
Report3.to_excel(writer, sheet_name='Report3')
output.to_excel(writer, sheet_name='output')
writer.save()


# In[ ]:




