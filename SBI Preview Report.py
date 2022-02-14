#!/usr/bin/env python
# coding: utf-8

# In[49]:




import pandas as pd
import numpy as np


# In[50]:




#Export Data
rw = pd.read_excel(r"C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Export Data\export.xlsx")
rw.head()


# In[51]:





#Client Data
client = pd.read_excel(r'C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Client Data\nclient.xlsx')


# In[52]:



#Sales Data
sale = pd.read_excel(r"C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Sales file\sales file.xlsx")
sale = sale.iloc[:-1 , :]


# In[53]:



sale = sale[['Lead Source','Month', 'GWP Amount']]
sale.head()


# In[54]:



table = pd.pivot_table(sale, index=['Month'], values = 'GWP Amount', aggfunc=np.size, fill_value=0)


# In[55]:


table1 = pd.pivot_table(sale, index=['Month'], values = 'GWP Amount', aggfunc=np.sum, fill_value=0)


# In[56]:


table = pd.merge(table ,table1, how = "inner", on = "Month")


# In[57]:


table = table.rename(columns = {'GWP Amount_x': 'Nop', 'GWP Amount_y': 'GWP'})


# In[58]:


table


# In[59]:




a = rw.value_counts(['phone_number_dialed']) 

x = pd.DataFrame(a)

x = x.rename(columns = {0 : 'Attempts'})


# In[60]:


rw["status_name"].fillna("Ringing", inplace = True)


# In[61]:




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
    
    if x == "Outbound Pre-Routing Drop":
        return "Not Connect"
    if x == "Switched Off":
        return "Not Connect"
    if x == "Switched_Off":
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
    if x == "Language Barrier":
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
    if x == "Call_Disconnected_by_Customer":
        return "Not Connect"
    if x == "Not_reachable":
        return "Not Connect"
    if x == "Not Connected":
        return "Not Connect"
    if x == "Test_Call":
        return "Non Lead"
    if x == "Non Lead":
        return "Non Lead"
    if x == "Lead To Be Called":
        return "Not Connect"
    
    
    
    else:
        return "Connect"
    
    
    


# In[62]:


rw['Dispo'] = rw['status_name'].apply(lambda x: test(x))


# In[63]:



def test1(x):
    if x == "Non Lead":
        return 1
    if x == "Connect":
        return 2
         
    if x == "Not Connect":
        return 3
    else:
        return 4


# In[64]:


rw['BD'] = rw['Dispo'].apply(lambda x: test1(x))


# In[65]:


rw = rw.sort_values("BD", ascending=True)

rw



# In[66]:


rw =  pd.merge(rw ,x, how = "inner", on = "phone_number_dialed")
rw.head()


# In[67]:





rw1 = rw.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


# In[68]:


rw1.head()


# In[69]:




rw1 = pd.merge(rw1, x, how = "inner", on = "phone_number_dialed")

rw1.head()


# In[70]:




client = client.rename(columns = {'Mobile Number' : 'phone_number_dialed'})
client.head()


# In[71]:


Result = pd.merge(client ,rw1, how = "left", on = "phone_number_dialed")


# In[72]:




Result = Result.rename(columns = {'Date' : 'Created', 'LOB Campaign' : 'Lead Source'})

Result['Campaign Name'] = 'Response'

Result.drop(['Attempts_y'], axis = 1)
Result.head()


# In[73]:




Result["Dispo"].fillna("No Dial", inplace = True)
Result["status_name"].fillna("No Dial", inplace = True)


# In[74]:



Result = Result.rename(columns = {'Attempts_x': 'Attempts'})
Result = Result.drop(['Attempts_y'], axis = 1)

Result.head()


# In[75]:




LeadSource= pd.pivot_table(data = Result, index = ['Month'], columns= ['Dispo'],values = 'Attempts', aggfunc = np.mean, fill_value= 0,margins = True, margins_name = 'Total' )



# In[76]:




LeadSource = LeadSource.rename(columns = {'Total': 'Total Attempts','Connect': 'Connect Attempts', 'Non Lead' : 'Non Lead Attempts', 'Not Connect' : 'NotConnect Attempts'})


LeadSource


# In[77]:




LeadSource1= pd.pivot_table(data = Result, index = ['Month'], columns= ['Dispo'],values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')


# In[78]:



LeadSource1 = LeadSource1.rename(columns = {'Total': 'Total Data Received'})

LeadSource1 = LeadSource1.iloc[:-1 , :]
LeadSource1 = pd.merge(LeadSource1 ,table, how = "left", on = "Month")

LeadSource1["Nop"].fillna(0, inplace = True)
LeadSource1["GWP"].fillna(0, inplace = True)

LeadSource1 = LeadSource1.append(LeadSource1.sum().rename('Total'))
LeadSource1


# In[79]:



LeadSource1['Total Valid Data'] = LeadSource1['Connect'] + LeadSource1['Not Connect'] 
LeadSource1


# In[80]:



LeadSourcedata = LeadSource1
LeadSourcedata = LeadSourcedata.iloc[:-1 , :]
LeadSourcedata['Conversion On Connect%'] = LeadSourcedata['Nop'] /  LeadSourcedata['Connect'] 
LeadSourcedata['Conversion On Over All%'] = LeadSourcedata['Nop'] /  LeadSourcedata['Total Valid Data'] 
LeadSourcedata = LeadSourcedata[['Conversion On Connect%', 'Conversion On Over All%']]
LeadSourcedata = LeadSourcedata.append(LeadSourcedata.mean().rename('Total'))

LeadSourcedata["Conversion On Connect%"].fillna(0, inplace = True)
LeadSourcedata["Conversion On Over All%"].fillna(0, inplace = True)

LeadSourcedata['Conversion On Connect%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in LeadSourcedata['Conversion On Connect%']], index = LeadSourcedata.index)
LeadSourcedata['Conversion On Over All%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in LeadSourcedata['Conversion On Over All%']], index = LeadSourcedata.index)


LeadSourcedata


# In[81]:



LeadSource12 = pd.merge(LeadSource1 ,LeadSourcedata, how = "left", on = "Month")
LeadSource13 = LeadSource12
LeadSource13


# In[82]:




LeadSource13['Valid Leads %'] = LeadSource13['Total Valid Data'] / LeadSource13['Total Data Received'] 
LeadSource13['Connectivity %'] = LeadSource13['Connect'] / LeadSource13['Total Valid Data'] 
LeadSource13['Invalid Leads %'] = LeadSource13['Non Lead'] / LeadSource13['Total Data Received'] 

LeadSource13['Connectivity %'] = pd.Series(["{0:.2f}%".format(val * 100) for val in LeadSource13['Connectivity %']], index = LeadSource13.index)
LeadSource13['Valid Leads %'] = pd.Series(["{0:.2f}%".format(val * 100) for val in LeadSource13['Valid Leads %']], index = LeadSource13.index)
LeadSource13['Invalid Leads %'] = pd.Series(["{0:.2f}%".format(val * 100) for val in LeadSource13['Invalid Leads %']], index = LeadSource13.index)


# In[83]:




LeadSource14 = LeadSource13

LeadSource14 = LeadSource14.rename(columns = {'Connect': 'Total Connected Data', 'Not Connect': 'Total Not Connected Leads', 'Non Lead': 'Invalid Data Count', 'Nop': 'Total Nop', 'GWP': 'Total Gwp'})
 
LeadSource14


# In[84]:




LeadSource15= pd.pivot_table(data = Result, index = ['Month'],values = 'Attempts', aggfunc = np.mean, fill_value= 0, margins = True, margins_name = 'Total')

LeadSource15 = LeadSource15.round()

LeadSource15 = pd.merge(LeadSource14 ,LeadSource15, how = "inner", on = "Month")

LeadSource15 = LeadSource15.rename(columns = {'Attempts': 'Avg. Attempt', 'No Dial' : 'Pending For Dial'})
LeadSource15 = LeadSource15[['Total Data Received', 'Total Valid Data', 'Total Connected Data', 'Total Not Connected Leads', 'Invalid Data Count','Pending For Dial', 'Valid Leads %', 'Connectivity %', 'Invalid Leads %', 'Avg. Attempt', 'Total Nop', 'Total Gwp', 'Conversion On Over All%', 'Conversion On Connect%']]
LeadSource15



# In[85]:




Main1 = LeadSource15.transpose()
Main1


# In[86]:


Result['status_name'] = Result['status_name'].replace(['Already Bought'],'ALREADY BOUGHT')
Result['status_name'] = Result['status_name'].replace(['Lead Being Called'],'Ringing')
Result['status_name'] = Result['status_name'].replace(['Out of Service'],'OUT OF SERVICE')
Result['status_name'] = Result['status_name'].replace(['Not reachable'],'Not Reachable')


# In[87]:




CN1 = Result.loc[(Result['Dispo'] == "Connect")]
CN1.head()


# In[88]:



CN11 = pd.pivot_table(data = CN1, index = ['Dispo', 'status_name'],values = 'Month', aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')

CN12 = pd.pivot_table(data = CN1, index = ['Dispo', 'status_name'],values = 'Attempts', aggfunc = np.mean, fill_value= 0, margins = True, margins_name = 'Total')
CN12 = CN12.round()

CN1 = pd.merge(CN11 ,CN12, how = "inner", on = ["Dispo", "status_name"])

CN1  


# In[89]:



CN2 = Result.loc[(Result['Dispo'] == "Not Connect")]
CN2.head()


# In[90]:




CN22 = pd.pivot_table(data = CN2, index = ['Dispo', 'status_name'],values = 'Month', aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')
CN23 = pd.pivot_table(data = CN2, index = ['Dispo', 'status_name'],values = 'Attempts', aggfunc = np.mean, fill_value= 0, margins = True, margins_name = 'Total')

CN23 = CN23.round()


# In[91]:


CN2 = pd.merge(CN22 ,CN23, how = "inner", on = ["Dispo", "status_name"])

CN2


# In[92]:



Con = CN1.append(CN2)
Con = Con.rename(columns = {'Month': 'Data Count', 'Attempts': 'Average of Attempts Count'})

Con



# In[93]:



NL = Result.loc[(Result['Dispo'] == "Non Lead")]
NL['status_name'] = NL['status_name'].replace(['Non Lead'],'Not Eligible')

NL.head()


# In[94]:



NL1 = pd.pivot_table(data = NL, index = ['Dispo', 'status_name'],values = 'Month', aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')
NL2 = pd.pivot_table(data = NL, index = ['Dispo', 'status_name'],values = 'Attempts', aggfunc = np.mean, fill_value= 0, margins = True, margins_name = 'Total')

NL2 = NL2.round()
NL = pd.merge(NL1 ,NL2, how = "inner", on = ["Dispo", "status_name"])
NL = NL.rename(columns = {'Month': 'Data Count', 'Attempts': 'Average of Attempts Count'})


NL


# In[95]:


writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\Preview report.xlsx", engine='xlsxwriter')


# In[96]:




Con.to_excel(writer, sheet_name='Con')
NL.to_excel(writer, sheet_name='NL')
Result.to_excel(writer, sheet_name='Result')
Main1.to_excel(writer, sheet_name='Main1')


writer.save()


# In[ ]:





# In[ ]:





# In[ ]:




