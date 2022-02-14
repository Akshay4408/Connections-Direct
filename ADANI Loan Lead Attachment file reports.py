#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:



#Client Data
client = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Loan Hindi\client\client.xlsx", parse_dates=['entry_date'])
client.head()


# In[3]:



#Export Data
export = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Loan Hindi\Overall Export of Loan Lead.xlsx")
#export = export[['call_date', 'phone_number_dialed', 'full_name', 'status_name', 'Disposition', 'BDD']]
export.head()


# In[4]:


client['entry_date'].value_counts()


# In[5]:


import datetime


client['entry_date'] = client['entry_date'].dt.strftime('%d %b, %Y')
client.head()


# In[6]:


client['entry_date'].value_counts()


# In[7]:


client = client.loc[(client['entry_date'] == '06 Aug, 2021') | (client['entry_date'] == '05 Aug, 2021') | (client['entry_date'] == '04 Aug, 2021') | (client['entry_date'] == '07 Aug, 2021') | (client['entry_date'] == '09 Aug, 2021')| (client['entry_date'] == '10 Aug, 2021') | (client['entry_date'] == '11 Aug, 2021') | (client['entry_date'] == '12 Aug, 2021') | (client['entry_date'] == '13 Aug, 2021') | (client['entry_date'] == '14 Aug, 2021') | (client['entry_date'] == '15 Aug, 2021') | (client['entry_date'] == '17 Aug, 2021') | (client['entry_date'] == '18 Aug, 2021') | (client['entry_date'] == '19 Aug, 2021')  | (client['entry_date'] == '20 Aug, 2021') | (client['entry_date'] == '21 Aug, 2021') | (client['entry_date'] == '22 Aug, 2021') | (client['entry_date'] == '23 Aug, 2021') | (client['entry_date'] == '24 Aug, 2021') | (client['entry_date'] == '25 Aug, 2021') | (client['entry_date'] == '26 Aug, 2021') | (client['entry_date'] == '27 Aug, 2021') | (client['entry_date'] == '28 Aug, 2021') | (client['entry_date'] == '30 Aug, 2021') ]
client.head()
                    


# In[8]:


client = client.rename(columns = {'entry_date': 'Entry_Date'})


# In[9]:


export["status_name"].fillna("Ringing", inplace = True)
export.head()


# In[10]:




export = export.sort_values("BDD", ascending=True)

export


# In[11]:





export = export.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


export.head()


# In[12]:


client = client.rename(columns = {'phone_number': 'phone_number_dialed'})
export = export[['phone_number_dialed', 'full_name', 'status_name', 'Disposition']]
export.head()


# In[13]:


client.info()


# In[14]:


client1 = client[['lead_id', 'Entry_Date', 'user', 'phone_number_dialed', 'Lead_ID_1', 'Cust_Name', 'Postal_Code_1', 'Lead_Source', 'Product', 'State_1', 'Agent_Remark']]
client1.head()


# In[15]:


client1 = client1.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


# In[16]:




Result =  pd.merge(client1 ,export, how = 'inner', on = "phone_number_dialed")
Result.head()


# In[17]:



StatusCount= pd.pivot_table(data = Result, index = ['Entry_Date'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#StatusCount['Avg%'] = StatusCount['(Count, Total)'] / 155
#StatusCount['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in StatusCount['Avg%']], index = StatusCount.index)

#StatusCount = StatusCount.rename(columns = {'status_name' : 'Count'})



StatusCount


# In[18]:


StatusCount.info()


# In[19]:


# Connect



connect = Result.loc[Result['Disposition'] == "Connect"]


Notconnect = Result.loc[Result['Disposition'] == "Not Connected"]


# In[20]:


connect1= pd.pivot_table(data = connect, index = ['Disposition','Entry_Date'],columns= ['status_name'] ,values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
connect1 = connect1.rename(columns = {'phone_number_dialed': 'Count'})


#connect1['Avg%'] = connect1['Count'] / 7
#connect1['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in connect1['Avg%']], index = connect1.index)


connect1


# In[21]:


Notconnect


# In[ ]:





# In[22]:




Notconnect2= pd.pivot_table(data = Notconnect, index = [ 'Disposition', 'Entry_Date'], columns= ['status_name'],values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
Notconnect2 = Notconnect2.rename(columns = {'phone_number_dialed': 'Count'})


#Notconnect2['Avg%'] = Notconnect2['Count'] / 8

#Notconnect2['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Notconnect2['Avg%']], index = Notconnect2.index)


Notconnect2


# In[23]:




data =  pd.merge(client ,export, how = 'inner', on = "phone_number_dialed")
data.head()


# In[24]:



datecount = pd.pivot_table(data = Result, index = ['Entry_Date'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#StatusCount['Avg%'] = StatusCount['(Count, Total)'] / 155
#StatusCount['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in StatusCount['Avg%']], index = StatusCount.index)

datecount = datecount.rename(columns = {'Total' : 'Total Data Received'})



datecount


# In[25]:


Success = connect.loc[(connect['status_name'] == "Confirmed Lead CC") |  (connect['status_name'] == "Confirmed_Lead _ QEC_Done")]
Success.head()


# In[26]:



SuccessLead = pd.pivot_table(data = Success, index = ['status_name'],values = ['Entry_Date'],aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')

SuccessLead['Avg%'] = SuccessLead['Entry_Date'] / 79
SuccessLead['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in SuccessLead['Avg%']], index = SuccessLead.index)

SuccessLead = SuccessLead.rename(columns = {'Entry_Date' : 'Total'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
SuccessLead


# In[27]:



Product = pd.pivot_table(data = Success, index = ['Product'],columns = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')

#Product['Avg%'] = Product['Disposition'] / 76
#Product['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Product['Avg%']], index = Product.index)

Product = Product.rename(columns = {'Disposition' : 'Total'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
Product


# In[28]:



Source = pd.pivot_table(data = Success, index = ['Lead_Source'],columns = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')

#Product['Avg%'] = Product['Disposition'] / 76
#Product['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Product['Avg%']], index = Product.index)

Source = Source.rename(columns = {'Disposition' : 'Total'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
Source


# In[29]:



State_1 = pd.pivot_table(data = Success, index = ['State_1'],columns = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')

#Product['Avg%'] = Product['Disposition'] / 76
#Product['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Product['Avg%']], index = Product.index)

State_1 = State_1.rename(columns = {'Disposition' : 'Total', 'State_1': 'State'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
State_1


# In[30]:


Result.head()


# In[31]:



Conn = pd.pivot_table(data = Result, index = ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

Conn['Avg%'] = Conn['status_name'] / 416
Conn['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Conn['Avg%']], index = Conn.index)

Conn = Conn.rename(columns = {'status_name' : 'Count'})



Conn


# In[32]:



ConnStatus = pd.pivot_table(data = connect, index = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

ConnStatus['Avg%'] = ConnStatus['Disposition'] / 331
ConnStatus['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in ConnStatus['Avg%']], index = ConnStatus.index)

ConnStatus = ConnStatus.rename(columns = {'Disposition' : 'Count'})



ConnStatus


# In[33]:



NonConnStatus = pd.pivot_table(data = Notconnect, index = ['status_name'],values = ['Disposition'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

NonConnStatus['Avg%'] = NonConnStatus['Disposition'] / 85
NonConnStatus['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in NonConnStatus['Avg%']], index = NonConnStatus.index)

NonConnStatus = NonConnStatus.rename(columns = {'Disposition' : 'Count'})



NonConnStatus


# In[34]:


Result.head()


# In[35]:



Statewise = pd.pivot_table(data = Result, index = ['Product'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#Statewise['Avg%'] = Statewise['status_name'] / 403
#Statewise['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Statewise['Avg%']], index = Statewise.index)

Statewise = Statewise.rename(columns = {'Total' : 'Total Data Received'})



Statewise


# In[36]:



Productwise = pd.pivot_table(data = Result, index = ['Product'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#Statewise['Avg%'] = Statewise['status_name'] / 403
#Statewise['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Statewise['Avg%']], index = Statewise.index)

Productwise = Productwise.rename(columns = {'Total' : 'Total Data Received'})



Productwise


# In[37]:



Sourcewise = pd.pivot_table(data = Result, index = ['Lead_Source'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#Statewise['Avg%'] = Statewise['status_name'] / 403
#Statewise['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Statewise['Avg%']], index = Statewise.index)

Sourcewise = Sourcewise.rename(columns = {'Total' : 'Total Data Received'})



Sourcewise


# In[38]:



Statewise = pd.pivot_table(data = Result, index = ['State_1'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#Statewise['Avg%'] = Statewise['status_name'] / 403
#Statewise['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Statewise['Avg%']], index = Statewise.index)

Statewise = Statewise.rename(columns = {'Total' : 'Total Data Received'})



Statewise


# In[39]:




writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\Loan Attachment report.xlsx", engine='xlsxwriter')


# In[40]:




State_1.to_excel(writer, sheet_name='State')
Notconnect2.to_excel(writer, sheet_name='Notconnect')
connect1.to_excel(writer, sheet_name='connect')
Product.to_excel(writer, sheet_name='Product')
Result.to_excel(writer, sheet_name='Result')

Source.to_excel(writer, sheet_name='Source')

writer.save()


# In[ ]:




