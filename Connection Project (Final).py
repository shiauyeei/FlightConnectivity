# -*- coding: utf-8 -*-
"""
Created on Fri Jun 19 15:23:53 2020

@author: Jimosy
"""

# Import Library (OK)
import pandas as pd
import numpy as np
import datetime as dt
import math


# Read Excel File (OK)
InHK = pd.read_excel(r'C:\Users\Jimosy\Desktop\Python Project\Connection Project\Jun_Jul20 Hubchart 24 hrs Original.xlsm', sheet_name = 'InHK')
OutHK = pd.read_excel(r'C:\Users\Jimosy\Desktop\Python Project\Connection Project\Jun_Jul20 Hubchart 24 hrs Original.xlsm', sheet_name = 'OutHK')
InHK['DayofArrival'] = ""
OutHK['DayofArrival'] = ""

################ Clean Time Data ################
# Clean Data - Convert OutHK['STD'] Str to Datetime format (OK)
OutHK['STD'] = pd.to_datetime(OutHK['STD'], format = '%H:%M')
InHK['STA'] = pd.to_datetime(InHK['STA'], format = '%H:%M', errors = 'ignore')
InHK_row = InHK.shape[0]
OutHK_row = OutHK.shape[0]


# Clean Data - Convert some InHK['STA'] str to Datetime format (OK)
for i in range(0, InHK_row):
    if type(InHK.at[i,'STA']) == str:
        InHK.at[i, 'STA'] = dt.datetime.strptime(InHK.at[i, 'STA'], '%H:%M').time()
    else:
        pass


# Clean Data - Convert all STA STD to timedelta (OK)
for i in range(0, InHK_row):
    InHK.loc[i,'STA'] = dt.timedelta(hours = InHK.iloc[i]['STA'].hour, minutes = InHK.iloc[i]['STA'].minute)
for j in range(0, OutHK_row):
    OutHK.loc[j,'STD'] = dt.timedelta(hours = OutHK.iloc[j]['STD'].hour, minutes = OutHK.iloc[j]['STD'].minute)


# To check the number of dt.time in InHK['STA'], make sure all converted to dt.time (OK)
STA_strcount = 0
for i in range(0, InHK_row):
    if type(InHK.at[i,'STA']) == dt.time:
        STA_strcount = STA_strcount + 1
print('STA_strcount = ', STA_strcount)


# To check the number of str in OutHK['STD'], str should return 0 (OK)
STD_strcount = 0
for j in range(0, OutHK_row):
    if type(OutHK.iloc[j]['STD']) == str:
        STD_strcount = STD_strcount + 1
print('STD_strcount = ', STD_strcount)


################ Clean Pattern (DOW/DayofArrival) Data ################
# Clean Data - Convert Pattern to string (OK)
InHK.Pattern = InHK.Pattern.astype(str)
OutHK.Pattern = OutHK.Pattern.astype(str)
print('Datatype of InHK[\'Pattern\']', type(InHK.iloc[0]['Pattern']))
print('Datatype of OutHK[\'Pattern\']', type(OutHK.iloc[0]['Pattern']))


# Clean Data - Remove '.' delimiter (OK)
InHK['Pattern'] = InHK['Pattern'].replace({'\.':''}, regex = True)
OutHK['Pattern'] = OutHK['Pattern'].replace({'\.':''}, regex = True)


########## Update specific pattern for DOW #################
InHK.loc[(InHK['Period'] == '01-11 Jul') & (InHK['Orig'] == 'CGK'), 'Pattern'] = '1234567'

# Update specific pattern for DOW (Specific Flight Number)
OutHK.loc[(OutHK['Period'] == '01-11 Jul') & (OutHK['Dest'] == 'CGK') & (OutHK['FlNo'] == 777), 'Pattern'] = '1234567'
OutHK.loc[(OutHK['Period'] == '01-11 Jul') & (OutHK['Dest'] == 'CGK') & (OutHK['FlNo'] == 797), 'Pattern'] = '1234567'

InHK.loc[(InHK['Period'] == '01-11 Jul') & (InHK['Orig'] == 'CGK') & (InHK['FlNo'] == 776), 'Pattern'] = '1234567'
InHK.loc[(InHK['Period'] == '01-11 Jul') & (InHK['Orig'] == 'CGK') & (InHK['FlNo'] == 798), 'Pattern'] = '1234567'


# Splitting and Exploding column Pattern to multiple rows (OK)
InHK = (InHK.set_index(InHK.columns.drop('Pattern',1).tolist()).
             Pattern.str.split('', expand=True).stack().
             reset_index().rename(columns={0:'Pattern'}).loc[:, InHK.columns])
InHK = InHK[InHK.Pattern != '']
InHK = InHK.reset_index(drop=True)
InHK['Pattern'] = InHK['Pattern'].astype(int)

OutHK = (OutHK.set_index(OutHK.columns.drop('Pattern',1).tolist()).
             Pattern.str.split('', expand=True).stack().
             reset_index().rename(columns={0:'Pattern'}).loc[:, OutHK.columns])
OutHK = OutHK[OutHK.Pattern != '']
OutHK = OutHK.reset_index(drop=True)
OutHK['Pattern'] = OutHK['Pattern'].astype(int)


# Clean Data - Convert In/OutHK['DayChange'] str/nan/int to int (OK)
# 1. convert nan to "" (str)
# 2. convert everything else to str
# 3. change "" to "0"
# 4. convert everything to int

# InHK (OK)
for i in range(0, InHK.shape[0]):
    if (type(InHK.at[i,'DayChange']) == float):
        if (math.isnan(InHK.at[i,'DayChange'])):
            # because we don't want nan for easier processing later
            InHK.at[i, 'DayChange'] = ''
    InHK.at[i, 'DayChange'] = str(InHK.at[i, 'DayChange']).strip()
    if (InHK.at[i, 'DayChange'] == ''):
        # because int("") results in error
        InHK.at[i, 'DayChange'] = 0
    elif (InHK.at[i,'DayChange'] == -1) & (InHK.at[i, 'Pattern'] == 1):
        InHK.at[i,'DayofArrival'] = 7
    else:
        InHK.at[i, 'DayChange'] = int(InHK.at[i, 'DayChange']) 


# Added data into InHK['DayofArrival'] (OK)
for i in range(0, InHK.shape[0]):
    if (InHK.at[i,'DayChange'] == 1) & (InHK.at[i, 'Pattern'] == 7):
        InHK.at[i,'DayofArrival'] = 1
    else: 
        InHK.at[i,'DayofArrival'] = int(InHK.at[i,'Pattern']) + InHK.at[i,'DayChange']


# OutHK (OK)
for i in range(0, OutHK.shape[0]):
    if (type(OutHK.at[i,'DayChange']) == float):
        if (math.isnan(OutHK.at[i,'DayChange'])):
            # because we don't want nan for easier processing later
            OutHK.at[i, 'DayChange'] = ''
    OutHK.at[i, 'DayChange'] = str(OutHK.at[i, 'DayChange']).strip()
    if (OutHK.at[i, 'DayChange'] == ''):
        # because int("") results in error
        OutHK.at[i, 'DayChange'] = 0
    else:
        OutHK.at[i, 'DayChange'] = int(OutHK.at[i, 'DayChange']) 


# Added data into OutHK['DayofArrival'] (OK)
for i in range(0, OutHK.shape[0]):
    if (OutHK.at[i,'DayChange'] == 1) & (OutHK.at[i, 'Pattern'] == 7):
        OutHK.at[i,'DayofArrival'] = 1
    elif (OutHK.at[i,'DayChange'] == -1) & (OutHK.at[i, 'Pattern'] == 1):
        OutHK.at[i,'DayofArrival'] = 7
    else: 
        OutHK.at[i,'DayofArrival'] = int(OutHK.at[i,'Pattern']) + OutHK.at[i,'DayChange']


# Write final to Excel
writer = pd.ExcelWriter(r'C:\Users\Jimosy\Desktop\Python Project\Connection Project\Jun_Jul20 Hubchart Python Cleaned.xlsx', engine  = 'xlsxwriter')
InHK.to_excel(writer, sheet_name='InHK_Clean', index = False)
OutHK.to_excel(writer, sheet_name='OutHK_Clean', index = False)
writer.save()


##############################################################################

# Extract and Calculate
# Create New DataFrame to hold final output
columnnames = ['InHK_Period', 'InHK_AI', 'InHK_FlNo', 'InHK_DepartureDay', 'InHK_ArrivalDay', 'InHK_Region', 'InHK_Orig', 'InHK_Dest', 'InHK_STD', 'InHK_STA', 
               'OutHK_Period', 'OutHK_AI', 'OutHK_FlNo', 'OutHK_DepartureDay', 'OutHK_ArrivalDay', 'OutHK_Region', 'OutHK_Orig', 'OutHK_Dest', 'OutHK_STD', 'OutHK_STA',
               'Layover']

df_OB = pd.DataFrame(columns = columnnames)
df_IB = pd.DataFrame(columns = columnnames)


# Calculate OutBound interested flights Layover Time & write to New DataFrame
# Specified intended period and origin/destination of interest
period = '01-11 Jul'
origin = 'CGK'
destination = 'CGK'

minlayover = dt.timedelta(hours = 0, minutes = 50)
maxlayover = dt.timedelta(hours = 24, minutes = 00)


# For OB flights
OB_InHK = InHK[(InHK['Orig'] == origin) & (InHK['Period'] == period)]
OB_InHK = OB_InHK.reset_index(drop = True)

OB_OutHK = OutHK[OutHK['Period'] == period]
OB_OutHK = OB_OutHK.reset_index(drop = True)
count_OB = 0


dictOne = {'InHK_Period':'Period', 'InHK_AI':'Al', 'InHK_FlNo':'FlNo', 'InHK_DepartureDay':'Pattern', 'InHK_ArrivalDay':'DayofArrival',
           'InHK_Region':'Region', 'InHK_Orig':'Orig', 'InHK_Dest':'Dest', 'InHK_STD':'STD', 'InHK_STA':'STA'}
dictTwo = {'OutHK_Period':'Period', 'OutHK_AI':'Al', 'OutHK_FlNo':'FlNo', 'OutHK_DepartureDay':'Pattern', 'OutHK_ArrivalDay':'DayofArrival',
           'OutHK_Region':'Region', 'OutHK_Orig':'Origin', 'OutHK_Dest':'Dest', 'OutHK_STD':'STD', 'OutHK_STA':'STA'}

for i in range(0, OB_InHK.shape[0]):
    
    for j in range(0, OB_OutHK.shape[0]):
        
        if ((OB_InHK.at[i,'DayofArrival'] == OB_OutHK.at[j,'Pattern']) & 
                (OB_OutHK.at[j,'STD'] - OB_InHK.at[i,'STA'] > minlayover)):
            layover = OB_OutHK.at[j,'STD'] - OB_InHK.at[i,'STA']
            
            df_OB = df_OB.append(pd.Series(np.nan), ignore_index = True)
            df_OB.at[count_OB,'InHK_Period']        = OB_InHK.at[i,dictOne['InHK_Period']]
            df_OB.at[count_OB,'InHK_AI']            = OB_InHK.at[i,dictOne['InHK_AI']]
            df_OB.at[count_OB,'InHK_FlNo']          = OB_InHK.at[i,dictOne['InHK_FlNo']]
            df_OB.at[count_OB,'InHK_DepartureDay']  = OB_InHK.at[i,dictOne['InHK_DepartureDay']]
            df_OB.at[count_OB,'InHK_ArrivalDay']    = OB_InHK.at[i,dictOne['InHK_ArrivalDay']]
            df_OB.at[count_OB,'InHK_Region']        = OB_InHK.at[i,dictOne['InHK_Region']]
            df_OB.at[count_OB,'InHK_Orig']          = OB_InHK.at[i,dictOne['InHK_Orig']]
            df_OB.at[count_OB,'InHK_Dest']          = OB_InHK.at[i,dictOne['InHK_Dest']]
            df_OB.at[count_OB,'InHK_STD']           = OB_InHK.at[i,dictOne['InHK_STD']]
            df_OB.at[count_OB,'InHK_STA']           = OB_InHK.at[i,dictOne['InHK_STA']]
            
            df_OB.at[count_OB,'OutHK_Period']       = OB_OutHK.at[j,dictTwo['OutHK_Period']]
            df_OB.at[count_OB,'OutHK_AI']           = OB_OutHK.at[j,dictTwo['OutHK_AI']]
            df_OB.at[count_OB,'OutHK_FlNo']         = OB_OutHK.at[j,dictTwo['OutHK_FlNo']]
            df_OB.at[count_OB,'OutHK_DepartureDay'] = OB_OutHK.at[j,dictTwo['OutHK_DepartureDay']]
            df_OB.at[count_OB,'OutHK_ArrivalDay']   = OB_OutHK.at[j,dictTwo['OutHK_ArrivalDay']]
            df_OB.at[count_OB,'OutHK_Region']       = OB_OutHK.at[j,dictTwo['OutHK_Region']]
            df_OB.at[count_OB,'OutHK_Orig']         = OB_OutHK.at[j,dictTwo['OutHK_Orig']]
            df_OB.at[count_OB,'OutHK_Dest']         = OB_OutHK.at[j,dictTwo['OutHK_Dest']]
            df_OB.at[count_OB,'OutHK_STD']          = OB_OutHK.at[j,dictTwo['OutHK_STD']]
            df_OB.at[count_OB,'OutHK_STA']          = OB_OutHK.at[j,dictTwo['OutHK_STA']]
            
            df_OB.at[count_OB,'Layover'] = layover
            count_OB = count_OB + 1
            
        elif ((((OB_OutHK.at[j,'Pattern'] - OB_InHK.at[i,'DayofArrival']) == 1) | 
                ((OB_OutHK.at[j,'Pattern'] - OB_InHK.at[i,'DayofArrival']) == -6)) & 
            ((OB_OutHK.at[j,'STD'] + maxlayover - OB_InHK.at[i,'STA']) > minlayover) & 
            ((OB_OutHK.at[j,'STD'] < OB_InHK.at[i,'STA']))):
            layover = OB_OutHK.at[j,'STD'] + maxlayover - OB_InHK.at[i,'STA']
                        
            df_OB = df_OB.append(pd.Series(np.nan), ignore_index = True)
            df_OB.at[count_OB,'InHK_Period']        = OB_InHK.at[i,dictOne['InHK_Period']]
            df_OB.at[count_OB,'InHK_AI']            = OB_InHK.at[i,dictOne['InHK_AI']]
            df_OB.at[count_OB,'InHK_FlNo']          = OB_InHK.at[i,dictOne['InHK_FlNo']]
            df_OB.at[count_OB,'InHK_DepartureDay']  = OB_InHK.at[i,dictOne['InHK_DepartureDay']]
            df_OB.at[count_OB,'InHK_ArrivalDay']    = OB_InHK.at[i,dictOne['InHK_ArrivalDay']]
            df_OB.at[count_OB,'InHK_Region']        = OB_InHK.at[i,dictOne['InHK_Region']]
            df_OB.at[count_OB,'InHK_Orig']          = OB_InHK.at[i,dictOne['InHK_Orig']]
            df_OB.at[count_OB,'InHK_Dest']          = OB_InHK.at[i,dictOne['InHK_Dest']]
            df_OB.at[count_OB,'InHK_STD']           = OB_InHK.at[i,dictOne['InHK_STD']]
            df_OB.at[count_OB,'InHK_STA']           = OB_InHK.at[i,dictOne['InHK_STA']]
            
            df_OB.at[count_OB,'OutHK_Period']       = OB_OutHK.at[j,dictTwo['OutHK_Period']]
            df_OB.at[count_OB,'OutHK_AI']           = OB_OutHK.at[j,dictTwo['OutHK_AI']]
            df_OB.at[count_OB,'OutHK_FlNo']         = OB_OutHK.at[j,dictTwo['OutHK_FlNo']]
            df_OB.at[count_OB,'OutHK_DepartureDay'] = OB_OutHK.at[j,dictTwo['OutHK_DepartureDay']]
            df_OB.at[count_OB,'OutHK_ArrivalDay']   = OB_OutHK.at[j,dictTwo['OutHK_ArrivalDay']]
            df_OB.at[count_OB,'OutHK_Region']       = OB_OutHK.at[j,dictTwo['OutHK_Region']]
            df_OB.at[count_OB,'OutHK_Orig']         = OB_OutHK.at[j,dictTwo['OutHK_Orig']]
            df_OB.at[count_OB,'OutHK_Dest']         = OB_OutHK.at[j,dictTwo['OutHK_Dest']]
            df_OB.at[count_OB,'OutHK_STD']          = OB_OutHK.at[j,dictTwo['OutHK_STD']]
            df_OB.at[count_OB,'OutHK_STA']          = OB_OutHK.at[j,dictTwo['OutHK_STA']]
            
            df_OB.at[count_OB,'Layover'] = layover
            count_OB = count_OB + 1
        
        else:
            pass


# For IB flights
IB_InHK = InHK[InHK['Period'] == period]
IB_InHK = IB_InHK.reset_index(drop = True)

IB_OutHK = OutHK[(OutHK['Dest'] == destination) & (OutHK['Period'] == period)]
IB_OutHK = IB_OutHK.reset_index(drop = True)
count_IB = 0

for i in range(0, IB_OutHK.shape[0]):
    
    for j in range(0, IB_InHK.shape[0]):
        
        if ((IB_OutHK.at[i,'Pattern'] == IB_InHK.at[j,'DayofArrival']) & 
                (IB_OutHK.at[i,'STD'] - IB_InHK.at[j,'STA'] > minlayover)):
            layover = IB_OutHK.at[i,'STD'] - IB_InHK.at[j,'STA']
            
            df_IB = df_IB.append(pd.Series(np.nan), ignore_index = True)
            df_IB.at[count_IB,'InHK_Period']        = IB_InHK.at[j,dictOne['InHK_Period']]
            df_IB.at[count_IB,'InHK_AI']            = IB_InHK.at[j,dictOne['InHK_AI']]
            df_IB.at[count_IB,'InHK_FlNo']          = IB_InHK.at[j,dictOne['InHK_FlNo']]
            df_IB.at[count_IB,'InHK_DepartureDay']  = IB_InHK.at[j,dictOne['InHK_DepartureDay']]
            df_IB.at[count_IB,'InHK_ArrivalDay']    = IB_InHK.at[j,dictOne['InHK_ArrivalDay']]
            df_IB.at[count_IB,'InHK_Region']        = IB_InHK.at[j,dictOne['InHK_Region']]
            df_IB.at[count_IB,'InHK_Orig']          = IB_InHK.at[j,dictOne['InHK_Orig']]
            df_IB.at[count_IB,'InHK_Dest']          = IB_InHK.at[j,dictOne['InHK_Dest']]
            df_IB.at[count_IB,'InHK_STD']           = IB_InHK.at[j,dictOne['InHK_STD']]
            df_IB.at[count_IB,'InHK_STA']           = IB_InHK.at[j,dictOne['InHK_STA']]
            
            df_IB.at[count_IB,'OutHK_Period']       = IB_OutHK.at[i,dictTwo['OutHK_Period']]
            df_IB.at[count_IB,'OutHK_AI']           = IB_OutHK.at[i,dictTwo['OutHK_AI']]
            df_IB.at[count_IB,'OutHK_FlNo']         = IB_OutHK.at[i,dictTwo['OutHK_FlNo']]
            df_IB.at[count_IB,'OutHK_DepartureDay'] = IB_OutHK.at[i,dictTwo['OutHK_DepartureDay']]
            df_IB.at[count_IB,'OutHK_ArrivalDay']   = IB_OutHK.at[i,dictTwo['OutHK_ArrivalDay']]
            df_IB.at[count_IB,'OutHK_Region']       = IB_OutHK.at[i,dictTwo['OutHK_Region']]
            df_IB.at[count_IB,'OutHK_Orig']         = IB_OutHK.at[i,dictTwo['OutHK_Orig']]
            df_IB.at[count_IB,'OutHK_Dest']         = IB_OutHK.at[i,dictTwo['OutHK_Dest']]
            df_IB.at[count_IB,'OutHK_STD']          = IB_OutHK.at[i,dictTwo['OutHK_STD']]
            df_IB.at[count_IB,'OutHK_STA']          = IB_OutHK.at[i,dictTwo['OutHK_STA']]
            
            df_IB.at[count_IB,'Layover'] = layover
            count_IB = count_IB + 1
            
        elif ((((IB_OutHK.at[i,'Pattern'] - IB_InHK.at[j,'DayofArrival']) == 1) | 
                ((IB_OutHK.at[i,'Pattern'] - IB_InHK.at[j,'DayofArrival']) == -6)) & 
            ((IB_OutHK.at[i,'STD'] + maxlayover - IB_InHK.at[j,'STA']) > minlayover) & 
            ((IB_OutHK.at[i,'STD'] < IB_InHK.at[j,'STA']))):
            layover = IB_OutHK.at[i,'STD'] + maxlayover - IB_InHK.at[j,'STA']
            
            df_IB = df_IB.append(pd.Series(np.nan), ignore_index = True)
            df_IB.at[count_IB,'InHK_Period']        = IB_InHK.at[j,dictOne['InHK_Period']]
            df_IB.at[count_IB,'InHK_AI']            = IB_InHK.at[j,dictOne['InHK_AI']]
            df_IB.at[count_IB,'InHK_FlNo']          = IB_InHK.at[j,dictOne['InHK_FlNo']]
            df_IB.at[count_IB,'InHK_DepartureDay']  = IB_InHK.at[j,dictOne['InHK_DepartureDay']]
            df_IB.at[count_IB,'InHK_ArrivalDay']    = IB_InHK.at[j,dictOne['InHK_ArrivalDay']]
            df_IB.at[count_IB,'InHK_Region']        = IB_InHK.at[j,dictOne['InHK_Region']]
            df_IB.at[count_IB,'InHK_Orig']          = IB_InHK.at[j,dictOne['InHK_Orig']]
            df_IB.at[count_IB,'InHK_Dest']          = IB_InHK.at[j,dictOne['InHK_Dest']]
            df_IB.at[count_IB,'InHK_STD']           = IB_InHK.at[j,dictOne['InHK_STD']]
            df_IB.at[count_IB,'InHK_STA']           = IB_InHK.at[j,dictOne['InHK_STA']]
            
            df_IB.at[count_IB,'OutHK_Period']       = IB_OutHK.at[i,dictTwo['OutHK_Period']]
            df_IB.at[count_IB,'OutHK_AI']           = IB_OutHK.at[i,dictTwo['OutHK_AI']]
            df_IB.at[count_IB,'OutHK_FlNo']         = IB_OutHK.at[i,dictTwo['OutHK_FlNo']]
            df_IB.at[count_IB,'OutHK_DepartureDay'] = IB_OutHK.at[i,dictTwo['OutHK_DepartureDay']]
            df_IB.at[count_IB,'OutHK_ArrivalDay']   = IB_OutHK.at[i,dictTwo['OutHK_ArrivalDay']]
            df_IB.at[count_IB,'OutHK_Region']       = IB_OutHK.at[i,dictTwo['OutHK_Region']]
            df_IB.at[count_IB,'OutHK_Orig']         = IB_OutHK.at[i,dictTwo['OutHK_Orig']]
            df_IB.at[count_IB,'OutHK_Dest']         = IB_OutHK.at[i,dictTwo['OutHK_Dest']]
            df_IB.at[count_IB,'OutHK_STD']          = IB_OutHK.at[i,dictTwo['OutHK_STD']]
            df_IB.at[count_IB,'OutHK_STA']          = IB_OutHK.at[i,dictTwo['OutHK_STA']]
            
            df_IB.at[count_IB,'Layover'] = layover
            count_IB = count_IB + 1
        
        else:
            pass



# Write IB & OB to Excel
writer = pd.ExcelWriter(r'C:\Users\Jimosy\Desktop\Python Project\Connection Project\Jun_Jul20 Hubchart Python IB_OB.xlsx', engine  = 'xlsxwriter')
df_OB.to_excel(writer, sheet_name='Orig-Dest_OB', index = False)
df_IB.to_excel(writer, sheet_name='Dest-Orig_IB', index = False)
writer.save()



