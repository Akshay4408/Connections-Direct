#!/usr/bin/env python
# coding: utf-8

# In[463]:



import pandas as pd
import numpy as np


# In[464]:


#Export Data
rw = pd.read_excel(r"C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Export Data\JULEX.xlsx")


# In[465]:


#Client Data
client = pd.read_excel(r'C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Client Data\JULY.xlsx')


# In[466]:



rw.head()


# In[467]:


rw.head()


# In[468]:


a = rw.value_counts(['phone_number_dialed']) 

x = pd.DataFrame(a)

x = x.rename(columns = {0 : 'Attempts'})


# In[469]:


rw.info()


# In[470]:


rw['status_name'].value_counts()


# In[471]:


def test(x):
    if x == "Busy Auto":
        return "Not Connect"
    if x == "Call Back":
        return "Connect"
    if x == "Call Back After Presentation":
        return "Connect"
    if x == "Not Interested":
        return "Connect"
    if x == "Language Barrier":
        return "Non Lead"
    if x == "DO NOT CALL":
        return "Non Lead"
    if x == "Not Eligible":
        return "Non Lead"
    if x == "Previous Sale":
        return "Non Lead"
    if x == "Motor_Interested_Callback":
        return "Connect"
    if x == "ALREADY BOUGHT":
        return "Non Lead"
    if x == "Sale Made":
        return "Connect"
    if x == "Done With Competition":
        return "Non Lead"
    if x == "Promise To Pay":
        return "Connect"
    if x == "NEVER_SHOW_INTEREST":
        return "Non Lead"
    if x == "Invalid Number":
        return "Non Lead"
  
    if x == "INTERESTED":
        return "Connect"
    if x == "Interested  Customer":
        return "Connect"
    if x == "Health Check Up":
        return "Connect"
    if x == "Renewal Case":
        return "Non Lead"
    if x == "Sale from Branch":
        return "Non Lead"
    if x == "Test Call":
        return "Non Lead"
    if x == "Agent Not Available":
        return "Not Connect"
    if x == "Blank Call":
        return "Not Connect"
    if x == "Call Disconnected by Customer":
        return "Not Connect"
    if x == "No Answer AutoDial":
        return "Not Connect"
    if x == "Ringing":
        return "Not Connect"
    if x == "Not Reachable":
        return "Not Connect"
    if x == "OUT OF SERVICE":
        return "Not Connect"
    if x == "Busy":
        return "Not Connect"
    if x == "Service Related":
        return "Not Connect"
    if x == "Outbound Pre-Routing Drop":
        return "Not Connect"
    if x == "Switched Off":
        return "Not Connect"
    if x == "Disconnected Number Auto":
        return "Not Connect"
    if x == "Lead Being Called":
        return "Not Connect"
    if x == "Disconnected Number Auto":
        return "Not Connect"
    if x == "Disconnected Number Auto":
        return "Not Connect"
    if x == "Disconnected Number Auto":
        return "Not Connect"
    if x == "Already Bought":
        return "Non Lead"
    if x == "DNC":
        return "Non Lead"
    if x == "enquiry":
        return "Non Lead"
    if x == "Other Quote":
        return "Non Lead"
    if x == "Renewal with Competition":
        return "Non Lead"
    if x == "Sale from Broker":
        return "Non Lead"
    if x == "Sale from Branch":
        return "Non Lead"
    if x == "Service Related":
        return "Non Lead"
    if x == "Last Month Sale":
        return "Non Lead"
    if x == "Previous Sale":
        return "Non Lead"
    if x == "Wrong Number":
        return "Non Lead"
    if x == "Customer_Never_Shown _Int":
        
        return "Non Lead"
    if x == "Wrong Number":
        return "Non Lead"
    if x == "Kavach_Rakshak":
        return "Non Lead"
    if x == "Dead Case Vehicle Sold":
        return "Non Lead"
    if x == "Language Barrier Telgu":
        return "Non Lead"
    if x == "Language_Barrier":
        return "Non Lead"
    if x == "Only Maternity":
        return "Non Lead"
    if x == "Policy_issued _n_System":
        return "Non Lead"
    if x == "Renewed with Competition":
        return "Non Lead"
    if x == "Engaged":
        return "Not Connect"
    if x == "Not reachable":
        return "Not Connect"
    if x == "Out of Service":
        return "Not Connect"
    
    else:
        return "Connect"
    
    
    
    

    
    


# In[472]:


def test3(x):
    if x == "Busy Auto":
        return "Valid"
    if x == "Call Back":
        return "Valid"
    if x == "Call Back After Presentation":
        return "Valid"
    if x == "Not Interested":
        return "Valid"
    if x == "Language Barrier":
        return "Invalid"
    if x == "DO NOT CALL":
        return "Invalid"
    if x == "Not Eligible":
        return "Invalid"
    if x == "not eligible":
        return "Invalid"
    if x == "Previous Sale":
        return "Invalid"
    if x == "Motor_Interested_Callback":
        return "Valid"
    if x == "ALREADY BOUGHT":
        return "Invalid"
    if x == "Sale Made":
        return "Valid"
    if x == "Done With Competition":
        return "Invalid"
    if x == "Promise To Pay":
        return "Valid"
    if x == "NEVER_SHOW_INTEREST":
        return "Invalid"
    if x == "Invalid Number":
        return "Invalid"
  
    if x == "INTERESTED":
        return "Valid"
    if x == "Interested  Customer":
        return "Valid"
    if x == "Health Check Up":
        return "Valid"
    if x == "Renewal Case":
        return "Invalid"
    if x == "Sale from Branch":
        return "Invalid"
    if x == "Test Call":
        return "Invalid"
    if x == "Agent Not Available":
        return "Valid"
    if x == "Blank Call":
        return "Valid"
    if x == "No Answer AutoDial":
        return "Valid"
    if x == "Ringing":
        return "Valid"
    if x == "Not Reachable":
        return "Valid"
    if x == "OUT OF SERVICE":
        return "Valid"
    if x == "Busy":
        return "Valid"
    if x == "Service Related":
        return "Valid"
    if x == "Outbound Pre-Routing Drop":
        return "Valid"
    if x == "Switched Off":
        return "Valid"
    if x == "Disconnected Number Auto":
        return "Valid"
    if x == "Lead Being Called":
        return "Valid"
    if x == "Disconnected Number Auto":
        return "Valid"
  
    else:
        return "Valid"


     


# In[473]:


rw['Dispo'] = rw['status_name'].apply(lambda x: test(x))


# In[474]:


rw['Validity'] = rw['status_name'].apply(lambda x: test3(x))


# In[475]:


rw.head()


# In[476]:


def test1(x):
   
    if x == "Call Back":
        return 7
    if x == "Call Back After Presentation":
        return 6
    if x == "Not Interested":
        return 9
    if x == "Language Barrier":
        return 3
    if x == "DO NOT CALL":
        return 3
    if x == "Not Eligible":
        return 3
    if x == "Previous Sale":
        return 3
    if x == "Motor_Interested_Callback":
        return 7
    if x == "ALREADY BOUGHT":
        return 3
    if x == "Sale Made":
        return 1
    if x == "Done With Competition":
        return 3
    if x == "Promise To Pay":
        return 2
    if x == "NEVER_SHOW_INTEREST":
        return 3
    if x == "Invalid Number":
        return 3
  
    if x == "INTERESTED":
        return 4
    if x == "Interested  Customer":
        return 4
    if x == "Health Check Up":
        return 7
    if x == "Renewal Case":
        return 3
    if x == "Sale from Branch":
        return 3
    if x == "Test Call":
        return 3
    if x == "Agent Not Available":
        return 15
    if x == "Blank Call":
        return 15
    if x == "No Answer AutoDial":
        return 15                        
    
    if x == "Ringing":
        return 15
    if x == "Not Reachable":
        return 15
    if x == "OUT OF SERVICE":
        return 15
    if x == "Busy":
        return 15
    if x == "Service Related":
        return 15
    if x == "Call Disconnected by Customer":
        return 15
    if x == "Outbound Pre-Routing Drop":
        return 15
    if x == "Switched Off":
        return 15
    if x == "Disconnected Number Auto":
        return 15
    if x == "Lead Being Called":
        return 15
    if x == "Disconnected Number Auto":
        return 15
    if x == "Disconnected Number Auto":
        return 15
    if x == "Disconnected Number Auto":
        return 15
    if x == "Already Bought":
        return 3
    if x == "DNC":
        return 3
    if x == "enquiry":
        return 3
    if x == "Other Quote":
        return 3
    if x == "Renewal with Competition":
        return 3
    if x == "Sale from Broker":
        return 3
    if x == "Sale from Branch":
        return 3
    if x == "Service Related":
        return 3
    if x == "Last Month Sale":
        return 3
    if x == "Previous Sale":
        return 3
    if x == "Wrong Number":
        return 3
    if x == "Busy Auto":
        return 15
    if x == "Call Disconnected by Customer":
        return 15
    if x == "Outbound Pre-Routing Drop":
        return 15
    if x == "Customer_Never_Shown _Int":
        
        return 3
    if x == "Kavach_Rakshak":
        return 3
    if x == "Dead Case Vehicle Sold":
        return 3
    if x == "Language Barrier Telgu":
        return 3
    if x == "Language_Barrier":
        return 3
    if x == "Only Maternity":
        return 3
    if x == "Policy_issued _n_System":
        return 3
    if x == "Renewed with Competition":
        return 3
    if x == "Engaged":
        return 15
    if x == "Not reachable":
        return 15
    if x == "Out of Service":
        return 15
    

    else:
        return 4
    
    
    
    


# In[477]:


rw['BD'] = rw['status_name'].apply(lambda x: test1(x))


# In[478]:


rw = rw.sort_values("BD", ascending=True)

rw.head()


# In[479]:


rw = rw.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


# In[480]:


#rw.to_excel(r'C:\Users\data anylitics\Desktop\Akshay\\Report.xlsx')


# In[481]:


rw.head()


# In[482]:


rw['status_name'].value_counts()


# In[483]:


rw = rw[['call_date','phone_number_dialed', 'full_name', 'status_name', 'Dispo', 'Validity' ]]
rw.head()


# In[484]:


sale = rw.loc[rw['status_name'] == "Sale Made"]
sale.head()


# In[485]:


sale['status_name'].value_counts()


# In[486]:


rw = pd.merge(rw, x, how = "inner", on = "phone_number_dialed")

rw.head()


# In[487]:


client = client.rename(columns = {'Mobile Number' : 'phone_number_dialed'})


# In[488]:


Result = pd.merge(client ,rw, how = "left", on = "phone_number_dialed")


# In[489]:


Result = Result.rename(columns = {'Date' : 'Created', 'LOB Campaign' : 'Lead Source'})

Result['Campaign Name'] = 'Response'


# In[490]:


Result["Dispo"].fillna("No Dial", inplace = True)
Result["status_name"].fillna("No Dial", inplace = True)
Result["Validity"].fillna("No Dial", inplace = True)
   


# In[491]:


Result.head()


# In[492]:


Connectvity = pd.pivot_table(data = Result, index = ['Dispo'], values = 'status_name', aggfunc = np.size, margins = True, margins_name = 'Total' )

Connectvity['Avg%'] = (Connectvity['status_name'] / 
                  55227) * 100

Connectvity['Avg%'] = Connectvity['Avg%'].astype(int)

Connectvity['Avg%'] = Connectvity['Avg%'].astype(str) + '%'

Connectvity = Connectvity.rename(columns = {'status_name' : 'Count'})
Connectvity


# In[493]:


NE = Result.loc[Result['status_name'] == "Not Eligible"]
NE.head()
#NE = NE['status_name'].value_counts()
#NE


# In[494]:


#CONNECT



connect = pd.pivot_table(data = Result, index = ['Dispo','status_name'], values = 'phone_number_dialed', aggfunc = np.size, margins = True, margins_name = 'Total' )

connect

connect['Avg%'] = (connect['phone_number_dialed'] / 
                  55227) * 100

connect['Avg%'] = connect['Avg%'].astype(int)

connect['Avg%'] = connect['Avg%'].astype(str) + '%'


connect = connect.rename(columns = {'phone_number_dialed' : 'Count'})

connect


# In[495]:


Connectvity1 = pd.pivot_table(data = Result, index = ['Validity'], values = 'status_name', aggfunc = np.size, margins = True, margins_name = 'Total' )

Connectvity1['Avg%'] = (Connectvity1['status_name'] / 
                  55227) * 100

Connectvity1['Avg%'] = Connectvity1['Avg%'].astype(int)

Connectvity1['Avg%'] = Connectvity1['Avg%'].astype(str) + '%'

Connectvity1 = Connectvity1.rename(columns = {'status_name' : 'Count'})

Connectvity1


# In[496]:


#CONNECT



valid = pd.pivot_table(data = Result, index = ['Validity', 'status_name'], values = 'phone_number_dialed', aggfunc = np.size, margins = True, margins_name = 'Total' )

valid

valid['Avg%'] = (valid['phone_number_dialed'] / 
                  55227) * 100

valid['Avg%'] = valid['Avg%'].astype(int)

valid['Avg%'] = valid['Avg%'].astype(str) + '%'

valid = valid.rename(columns = {'phone_number_dialed' : 'Count'})


valid


# In[497]:


Result.head()


# In[498]:


#CONNECT



campaigncon= pd.pivot_table(data = Result, index = ['Lead Source'], columns= ['Dispo'],values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

campaigncon = campaigncon.rename(columns = {'Total': 'Total Data Received'})
campaigncon['%Connect'] = campaigncon['Connect'] / campaigncon['Total Data Received']
campaigncon['%Non Lead'] = campaigncon['Non Lead'] / campaigncon['Total Data Received']
campaigncon['%No Dial'] = campaigncon['No Dial'] / campaigncon['Total Data Received']
campaigncon['Dialed Data'] = campaigncon['Total Data Received'] - campaigncon['No Dial']
campaigncon['%Dial'] = campaigncon['Dialed Data'] / campaigncon['Total Data Received']

#campaigncon['%Connect'] = pd.Series(["{0:.2f}%".format(val * 100) for val in campaigncon['%Connect']], index = campaigncon.index)
#campaigncon['%Non Lead'] = pd.Series(["{0:.2f}%".format(val * 100) for val in campaigncon['%Non Lead']], index = campaigncon.index)
#campaigncon['%No Dial'] = pd.Series(["{0:.2f}%".format(val * 100) for val in campaigncon['%No Dial']], index = campaigncon.index)
#campaigncon['%Dial'] = pd.Series(["{0:.2f}%".format(val * 100) for val in campaigncon['%Dial']], index = campaigncon.index)

campaigncon = campaigncon[['Total Data Received', 'Dialed Data', 'No Dial', 'Connect', 'Not Connect', 'Non Lead', '%Dial','%No Dial','%Connect', '%Non Lead']]
campaigncon.head()


# In[499]:


#Validity



#campaignval= pd.pivot_table(data = Result, index = ['Lead Source'], columns= ['Validity'],values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )
#campaignval = campaignval.rename(columns = {'Total': 'Total Data Received'})
#campaignval['%Connect'] = campaignval['Connect'] / campaignval['Total Data Received']
#amit['%Valid'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit['%Valid']], index = amit.index)
#campaignval = campaignval.drop(['Total'])
#campaignval.head()


# In[500]:


Result.head()


# In[501]:


sales = Result.loc[Result['status_name'] == "Sale Made"]
sales


# In[502]:


sales= pd.pivot_table(data = sales, index = ['Lead Source'], columns= ['status_name'],values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0 )



#sales = sales.append(sales.sum().rename('Total'))
sales


# In[503]:


amit = pd.concat([campaigncon, sales], axis=1)
amit.fillna(0, inplace = True)
amit.drop(amit.tail(1).index,inplace=True) # drop last n rows
#amit = amit.append(amit.sum().rename('Total'))

amit['%Connect Conversion'] = amit['Sale Made'] / amit['Connect']


#amit['%Connect'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit['%Connect']], index = amit.index)
#amit['%Non Lead'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit['%Non Lead']], index = amit.index)
#amit['%No Dial'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit['%No Dial']], index = amit.index)
#amit['%Dial'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit['%Dial']], index = amit.index)
#amit['%Connect Conversion'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit['%Connect Conversion']], index = amit.index)

amit


# In[504]:


amit1 = amit[['Total Data Received', 'Dialed Data', 'No Dial', 'Connect', 'Not Connect', 'Non Lead', 'Sale Made']]
amit1 = amit1.append(amit1.sum().rename('Total'))
amit1


# In[505]:


amit2 = amit[['%Dial', '%No Dial', '%Connect', '%Non Lead', '%Connect Conversion']]
amit2 = amit2.append(amit2.mean().rename('Total'))


amit2['%Connect'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit2['%Connect']], index = amit2.index)
amit2['%Non Lead'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit2['%Non Lead']], index = amit2.index)
amit2['%No Dial'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit2['%No Dial']], index = amit2.index)
amit2['%Dial'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit2['%Dial']], index = amit2.index)
amit2['%Connect Conversion'] = pd.Series(["{0:.2f}%".format(val * 100) for val in amit2['%Connect Conversion']], index = amit2.index)


amit2


# In[506]:


amitreport = pd.merge(amit1 ,amit2, how = "left", on = "Lead Source")


# In[507]:


amitreport.index.name


# In[508]:


amitreport =  amitreport.rename_axis('Campaigns')

amitreport = amitreport[['Total Data Received', 'Dialed Data', 'No Dial', 'Connect', 'Not Connect', 'Non Lead', '%Dial', '%No Dial', '%Connect', '%Non Lead', 'Sale Made', '%Connect Conversion']]


amitreport


# In[509]:


#Validity



campaignstat= pd.pivot_table(data = Result, index= ['Dispo','status_name'],columns = ['Lead Source'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

campaignstat.head()


# In[510]:


Result.head()


# In[511]:


#Result.drop(['call_date', 'Mobile Number.1'], axis=1)
Result.drop(['call_date'], axis=1)


# In[512]:


get_ipython().system(' pip install xlsxwriter')


# In[513]:


writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\amitmis.xlsx", engine='xlsxwriter')


# In[514]:


#Connectvity.to_excel(writer, sheet_name='Connectvity')
#best.to_excel(writer, sheet_name='best')
#Connectvity.to_excel(writer, sheet_name='Connectvity')
#campaignstat.to_excel(writer, sheet_name='campaignstat')
#Connectvity1.to_excel(writer, sheet_name='Connectvity1')
campaignstat.to_excel(writer, sheet_name='campaign')
Connectvity.to_excel(writer, sheet_name='Connectvity')
amitreport.to_excel(writer, sheet_name='amitreport')
Result.to_excel(writer, sheet_name='Result')
writer.save()


# In[ ]:





# In[ ]:





# In[ ]:




