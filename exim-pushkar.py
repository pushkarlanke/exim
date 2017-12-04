
# coding: utf-8

# In[1]:

import pandas as pd
import difflib , datetime # difflib package is required the calc the % matching words


# In[2]:

start = datetime.datetime.now() # Code start time

print('Code initialization stated at: ' + str(start))


# In[3]:

# Creating Dataframe for input excel
df = pd.read_excel("EXIM_Data/EXIM_Data.xlsx")



# In[4]:

# Creating dataframe for all the masters
MakeModel = pd.read_excel("masters/Make-Model.xlsx")
Transmission = pd.read_excel("masters/Exim_Transmission.xlsx")
FuelType = pd.read_excel("masters/Fuel_Type.xlsx")
CountryPort = pd.read_excel("masters/Country_port.xlsx")


# In[5]:

start2 = datetime.datetime.now() # Code start time

print('Code execution stated at: ' + str(start2))


# In[6]:

def format(InputString1):
    return InputString1.upper()


# In[7]:

df['ITEMTRIM'] =df['ITEM'].map(format)
MakeModel['MakeT'] = MakeModel['MAKE'].map(format)
MakeModel['ModelT'] = MakeModel['MODEL'].map(format)
Transmission['TCode'] = Transmission['Transmission type'].map(format)


# In[8]:

# not needed if make model are presorted in reuqired order.
MakeModel.sort_values(by=['MakeT','ModelT'],ascending=[True,True],inplace =True)


# In[9]:

df['ITCCODE']=df['ITC-CODE'].astype(str).str[0:6]
df['Engine cc']= 'NaN'
df['Fuel']= 'NaN'
df['Tranmission code']= 'NaN'


# In[10]:

for idx, row in MakeModel.iterrows():
    df.loc[df[df.ITEMTRIM.str.contains(row['MakeT'],regex =False)].index,'make']=row['MAKE']
    df.loc[df[df.ITEMTRIM.str.contains(row['ModelT'],regex =False)].index,'model']=row['MODEL']


# In[11]:

for idx, row in Transmission.iterrows():
    df.loc[df[df.ITEMTRIM.str.contains(row['TCode'],regex =False)].index,'Tranmission code']=row['Transmission code']


# In[12]:

for idx, row in FuelType.iterrows():
    #df.loc[df[df.ITCCODE==row['ITC Code']].index,'ITC Code' ]=row['ITC Code']
    df.loc[df[df.ITCCODE==row['ITC Code']].index,'Fuel' ]=row['Fuel'] 
    df.loc[df[df.ITCCODE==row['ITC Code']].index,'Engine cc' ]=row['Engine cc'] 


# In[13]:

end = datetime.datetime.now() # Code end time

print('Code execution Completed at: ' + str(end))

print('Total execution time taken: ' + str(end - start2)) # Total execution time


# In[14]:

output = df[['ITEM','ITC-CODE','make','model','Tranmission code','Tranmission code','ITCCODE','Fuel','Engine cc']]


# In[15]:

output.to_csv('output.csv')


# In[16]:

end2 = datetime.datetime.now() # Code end time

print('Code cleanup Completed at: ' + str(end2))

print('Total execution time taken: ' + str(end2 - start)) # Total execution time


# In[ ]:




# In[ ]:



