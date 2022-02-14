#!/usr/bin/env python
# coding: utf-8

# In[85]:


import pandas as pd
import numpy as np


# In[86]:



#Client Data
client = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Loan Hindi\client\client.xlsx", parse_dates=['entry_date'])
client.head()


# In[87]:



#Export Data
export = pd.read_excel(r"C:\Users\data anylitics\Desktop\ADANI\Loan Hindi\Overall Export of Loan Lead.xlsx")
#export = export[['call_date', 'phone_number_dialed', 'full_name', 'status_name', 'Disposition', 'BDD']]
export.head()


# In[88]:


import datetime


client['entry_date'] = client['entry_date'].dt.strftime('%d %b, %Y')
client.head()


# In[89]:


client['entry_date'].value_counts()


# In[90]:


client = client.loc[(client['entry_date'] == '06 Aug, 2021') | (client['entry_date'] == '05 Aug, 2021') | (client['entry_date'] == '04 Aug, 2021') | (client['entry_date'] == '07 Aug, 2021') | (client['entry_date'] == '09 Aug, 2021')| (client['entry_date'] == '10 Aug, 2021') | (client['entry_date'] == '11 Aug, 2021') | (client['entry_date'] == '12 Aug, 2021') | (client['entry_date'] == '13 Aug, 2021') | (client['entry_date'] == '14 Aug, 2021') | (client['entry_date'] == '15 Aug, 2021') | (client['entry_date'] == '17 Aug, 2021') | (client['entry_date'] == '18 Aug, 2021') | (client['entry_date'] == '19 Aug, 2021')  | (client['entry_date'] == '20 Aug, 2021') | (client['entry_date'] == '21 Aug, 2021') | (client['entry_date'] == '22 Aug, 2021') | (client['entry_date'] == '23 Aug, 2021') | (client['entry_date'] == '24 Aug, 2021') | (client['entry_date'] == '25 Aug, 2021') | (client['entry_date'] == '26 Aug, 2021') | (client['entry_date'] == '27 Aug, 2021') | (client['entry_date'] == '28 Aug, 2021')  | (client['entry_date'] == '31 Aug, 2021')]
client.head()
                    


# In[91]:


client = client.rename(columns = {'entry_date': 'Entry_Date'})


# In[92]:


export["status_name"].fillna("Ringing", inplace = True)
export.head()


# In[93]:




export = export.sort_values("BDD", ascending=True)

export


# In[94]:





export = export.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


export.head()


# In[95]:


client = client.rename(columns = {'phone_number': 'phone_number_dialed'})
export = export[['phone_number_dialed', 'full_name', 'status_name', 'Disposition']]
export.head()


# In[96]:


client1 = client[['lead_id', 'Entry_Date', 'user', 'phone_number_dialed', 'Lead_ID_1', 'Cust_Name', 'Postal_Code_1', 'Lead_Source', 'Product', 'State_1', 'Agent_Remark']]
client1.head()


# In[97]:


client1 = client1.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


# In[98]:




Result =  pd.merge(client1 ,export, how = 'inner', on = "phone_number_dialed")
Result.head()


# In[99]:



Productwise = pd.pivot_table(data = Result, index = ['Product'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#Statewise['Avg%'] = Statewise['status_name'] / 403
#Statewise['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Statewise['Avg%']], index = Statewise.index)

Productwise = Productwise.rename(columns = {'Total' : 'Total Data Received'})



Productwise


# In[100]:



Sourcewise = pd.pivot_table(data = Result, index = ['Lead_Source'],columns= ['Disposition'],values = ['status_name'],aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

#Statewise['Avg%'] = Statewise['status_name'] / 403
#Statewise['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Statewise['Avg%']], index = Statewise.index)

Sourcewise = Sourcewise.rename(columns = {'Total' : 'Total Data Received'})



Sourcewise


# In[101]:


# Connect



connect = Result.loc[Result['Disposition'] == "Connect"]


Notconnect = Result.loc[Result['Disposition'] == "Not Connected"]


# In[102]:


Success = connect.loc[(connect['status_name'] == "Confirmed Lead CC") |  (connect['status_name'] == "Confirmed_Lead _ QEC_Done")]
Success.head()


# In[103]:



SuccessLead = pd.pivot_table(data = Success, index = ['status_name'],values = ['Entry_Date'],aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')

SuccessLead['Avg%'] = SuccessLead['Entry_Date'] / 80
SuccessLead['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in SuccessLead['Avg%']], index = SuccessLead.index)

SuccessLead = SuccessLead.rename(columns = {'Entry_Date' : 'Total'})


#SuccessLead.sort_values(by='Disposition', ascending=False)
SuccessLead


# In[104]:




writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\Loan Lead SPS.xlsx", engine='xlsxwriter')


# In[105]:




SuccessLead.to_excel(writer, sheet_name='SuccessLead')
Sourcewise.to_excel(writer, sheet_name='Sourcewise')
Productwise.to_excel(writer, sheet_name='Productwise')
#StatusCount.to_excel(writer, sheet_name='Datecount')
Result.to_excel(writer, sheet_name='Result')


writer.save()


# In[ ]:




