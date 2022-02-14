#!/usr/bin/env python
# coding: utf-8

# In[55]:




import pandas as pd
import numpy as np


# In[56]:



#Export Data
rw = pd.read_excel(r"C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Export Data\export.xlsx")
rw.head()


# In[57]:




#rw['status_name'] = rw['status_name'].replace(['Not Eligible'],'Non Lead')


# In[58]:





#Client Data
client = pd.read_excel(r'C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Client Data\nclient.xlsx')



# In[59]:



#Sales Data
sale = pd.read_excel(r"C:\Users\data anylitics\Desktop\SBI\MIS\HEALTH\Sales file\sales file.xlsx")
sale = sale.iloc[:-1 , :]


# In[60]:




sale = sale[['Lead Source', 'GWP Amount']]
sale.head()


# In[61]:



table = pd.pivot_table(sale, index=['Lead Source'], values = 'GWP Amount', aggfunc=np.size, fill_value=0)


# In[62]:


table1 = pd.pivot_table(sale, index=['Lead Source'], values = 'GWP Amount', aggfunc=np.sum, fill_value=0)


# In[63]:


table = pd.merge(table ,table1, how = "inner", on = "Lead Source")


# In[64]:


table = table.rename(columns = {'GWP Amount_x': 'Nop', 'GWP Amount_y': 'GWP'})
table


# In[65]:


rw.head()


# In[66]:




a = rw.value_counts(['phone_number_dialed']) 

x = pd.DataFrame(a)

x = x.rename(columns = {0 : 'Attempts'})



# In[67]:


rw["status_name"].fillna("Ringing", inplace = True)


# In[68]:





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
    
    
    


# In[69]:


rw['Dispo'] = rw['status_name'].apply(lambda x: test(x))


# In[70]:



def test1(x):
    if x == "Non Lead":
        return 1
    if x == "Connect":
        return 2
         
    if x == "Not Connect":
        return 3
    else:
        return 4


# In[71]:


rw['BD'] = rw['Dispo'].apply(lambda x: test1(x))


# In[72]:




rw = rw.sort_values("BD", ascending=True)

rw


# In[73]:




rw =  pd.merge(rw ,x, how = "inner", on = "phone_number_dialed")
rw.head()


# In[74]:





rw1 = rw.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


rw1.head()


# In[75]:



rw1 = pd.merge(rw1, x, how = "inner", on = "phone_number_dialed")

rw1.head()


# In[76]:


client = client.rename(columns = {'Mobile Number' : 'phone_number_dialed'})
client.head()


# In[77]:


Result = pd.merge(client ,rw1, how = "left", on = "phone_number_dialed")


# In[78]:




Result = Result.rename(columns = {'Date' : 'Created', 'LOB Campaign' : 'Lead Source'})


# In[79]:


Result['Campaign Name'] = 'Response'


# In[80]:



Result.drop(['Attempts_y'], axis = 1)
Result.head()


# In[81]:




Result["Dispo"].fillna("No Dial", inplace = True)
Result["status_name"].fillna("No Dial", inplace = True)



# In[82]:



Result = Result.rename(columns = {'Attempts_x': 'Attempts'})

Result.head()


# In[83]:




LeadSource= pd.pivot_table(data = Result, index = ['Lead Source'], columns= ['Dispo'],values = 'Attempts', aggfunc = np.mean, fill_value= 0,margins = True, margins_name = 'Total' )



# In[84]:




LeadSource = LeadSource.rename(columns = {'Total': 'Total Attempts','Connect': 'Connect Attempts', 'Non Lead' : 'Non Lead Attempts', 'Not Connect' : 'NotConnect Attempts'})


LeadSource


# In[85]:




LeadSource1= pd.pivot_table(data = Result, index = ['Lead Source'], columns= ['Dispo'],values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0, margins = True, margins_name = 'Total')


# In[86]:


LeadSource1 = LeadSource1.rename(columns = {'Total': 'Total Data Received'})


# In[87]:


LeadSource1 = LeadSource1.iloc[:-1 , :]


# In[88]:


LeadSource1 = pd.merge(LeadSource1 ,table, how = "left", on = "Lead Source")


# In[89]:


LeadSource1["Nop"].fillna(0, inplace = True)
LeadSource1["GWP"].fillna(0, inplace = True)


# In[90]:


LeadSource1 = LeadSource1.append(LeadSource1.sum().rename('Total'))
LeadSource1


# In[91]:



LeadSourcedata = LeadSource1
LeadSourcedata = LeadSourcedata.iloc[:-1 , :]
LeadSourcedata['Conversion On Connect'] = LeadSourcedata['Nop'] /  LeadSourcedata['Connect'] 
LeadSourcedata['Conversion On Over All'] = LeadSourcedata['Nop'] /  LeadSourcedata['Total Data Received'] 
LeadSourcedata = LeadSourcedata[['Conversion On Connect', 'Conversion On Over All']]
LeadSourcedata = LeadSourcedata.append(LeadSourcedata.mean().rename('Total'))


# In[92]:



LeadSourcedata["Conversion On Connect"].fillna(0, inplace = True)
LeadSourcedata["Conversion On Over All"].fillna(0, inplace = True)

LeadSourcedata['Conversion On Connect'] = pd.Series(["{0:.2f}%".format(val * 100) for val in LeadSourcedata['Conversion On Connect']], index = LeadSourcedata.index)
LeadSourcedata['Conversion On Over All'] = pd.Series(["{0:.2f}%".format(val * 100) for val in LeadSourcedata['Conversion On Over All']], index = LeadSourcedata.index)


LeadSourcedata


# In[93]:



LeadSource12 = pd.merge(LeadSource1 ,LeadSourcedata, how = "left", on = "Lead Source")
LeadSource13 = LeadSource12
LeadSource13


# In[94]:



Result1 = pd.merge(LeadSource ,LeadSource13, how = "inner", on = "Lead Source")


# In[95]:


Result1 = Result1.rename(columns = {'No Dial_y': 'No Dial'})


# In[96]:



Result1 = Result1[['Total Data Received', 'Total Attempts','Connect', 'Connect Attempts', 'Not Connect', 'NotConnect Attempts', 'Non Lead', 'Non Lead Attempts', 'No Dial']]

Result1 = Result1.round()


# In[97]:



#Result1 = Result1.data

Result1 = Result1.rename(columns = {'Total Attempts': 'Attempts', 'Connect Attempts': 'Attempts', 'NotConnect Attempts': 'Attempts', 'Non Lead Attempts': 'Attempts'})

Result1


# In[98]:



Result2 = LeadSource12


# In[99]:


Result2['Total Valid Data'] = Result2['Total Data Received'] - Result2['Non Lead'] 
Result2 = Result2[['Total Data Received', 'Connect', 'Not Connect', 'Non Lead', 'Total Valid Data','Nop','GWP','Conversion On Over All','Conversion On Connect','No Dial']]

Result2


# In[100]:



Result3 = Result2

Result3 = Result3.drop(['Total'])

Result3 = Result3.append(Result3.mean().rename('Total'))

Result3


# In[101]:



#Result3['%Connect'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Result3['%Connect']], index = Result3.index)


# In[102]:



Result3['Connect%'] = Result3['Connect'] / Result3['Total Valid Data']

Result3['Connect%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Result3['Connect%']], index = Result3.index)

Result3['Valid%'] = Result3['Total Valid Data'] / Result3['Total Data Received']

Result3['Valid%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Result3['Valid%']], index = Result3.index)

Result3 = Result3[['Connect%', 'Valid%']]


Result3


# In[103]:



AMIT = pd.merge(Result2 ,Result3, how = "left", on = "Lead Source")

#AMIT = pd.merge(AMIT ,table, how = "left", on = "Lead Source")

AMIT = AMIT[['Total Data Received', 'Connect', 'Not Connect', 'Non Lead', 'Total Valid Data','Nop','GWP','Valid%','Connect%','Conversion On Over All','Conversion On Connect','No Dial']]


AMIT


# In[104]:


Result.head()


# In[105]:



Result['Created'] = pd.to_datetime(Result['Created']).dt.date

Result['call_date'] = pd.to_datetime(Result['call_date']).dt.date

rw['call_date'] = pd.to_datetime(Result['call_date']).dt.date

Result5 = Result


# In[106]:


rw5 = rw
Result5['status_name'] = Result5['status_name'].replace(['Non Lead'],'Not Eligible')
Result5['status_name'] = Result5['status_name'].replace(['Already Bought'],'ALREADY BOUGHT')
Result5['status_name'] = Result5['status_name'].replace(['Lead Being Called'],'Ringing')
Result5['status_name'] = Result5['status_name'].replace(['Out of Service'],'OUT OF SERVICE')
Result5['status_name'] = Result5['status_name'].replace(['Not reachable'],'Not Reachable')




rw5['status_name'] = rw5['status_name'].replace(['Non Lead'],'Not Eligible')
rw5['status_name'] = rw5['status_name'].replace(['Already Bought'],'ALREADY BOUGHT')
rw5['status_name'] = rw5['status_name'].replace(['Lead Being Called'],'Ringing')
rw5['status_name'] = rw5['status_name'].replace(['Out of Service'],'OUT OF SERVICE')
rw5['status_name'] = rw5['status_name'].replace(['Not reachable'],'Not Reachable')

Result.head()


# In[107]:


DispoST= pd.pivot_table(data = Result, index = ['Dispo', 'status_name'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

DispoST = DispoST.rename(columns = {'phone_number_dialed': 'Count'})


# In[108]:




writer = pd.ExcelWriter(r"C:\Users\data anylitics\Desktop\Akshay\call report.xlsx", engine='xlsxwriter')


# In[109]:




Result1.to_excel(writer, sheet_name='Result1')
AMIT.to_excel(writer, sheet_name='AMIT')
rw5.to_excel(writer, sheet_name='rw5')
DispoST.to_excel(writer, sheet_name='DispoST')
Result5.to_excel(writer, sheet_name='Result5')


writer.save()


# In[ ]:





# In[ ]:





# In[ ]:




