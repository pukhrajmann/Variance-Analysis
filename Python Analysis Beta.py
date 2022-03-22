#!/usr/bin/env python
# coding: utf-8

# In[30]:


#import necessarcy libraries
import numpy as np
import pandas as pd
import statistics as st
import xlrd

#Read excel file,  analysis sheet and create dataframes
#Make sure to save file to downloads folder change name to your own
df_main = pd.read_excel(r'C:\Users\pukhraj.mann\Downloads\Weekly Review V2.xlsm')
df_sheet = pd.read_excel('Weekly Review V2.xlsm', sheet_name='Analysis')

#Column of dataframe, use # to switch
i = [0,1,3,4]
#i = [0,1,2,4,5]
#i = [0,1,2,4,5,6]

for x in i:
    if (x<=0):
        df = df_sheet[['Applied']]
        df1 = df_sheet[['Stnd']]
    else:
        df = df_sheet[['Applied' + "." + str(x)]]
        df1 = df_sheet[['Stnd' + "." + str(x)]]
    
    #Convert dataframe columns to numpy arrays
    df_clean =df.dropna()
    df_clean1 =df1.dropna()
    arr = df_clean.to_numpy()
    arr1 = df_clean1.to_numpy()

    #Input Time Standard
    time_stand = arr1[-1].item()
    count = len(arr)

    #Make a nontype an int
    int (count or 0)
    int(0 if count is None else count)

    #Calculate iqr
    arr_75,arr_25 = np.percentile(arr, [75,25])
    iqr = arr_75 - arr_25

    #Identify outliers
    upper_bound = 1.5*iqr + np.percentile(arr,[75])
    lower_bound = np.percentile(arr,[25]) - 1.5*iqr

    #Remove outliers
    arr = arr[ (arr >= lower_bound) & (arr <= upper_bound) ]

    #Count elements left, and determine whether to split or not
    if (int(count) < 20):
        count = count
    else: 
        split = np.array_split(arr, 2)
        first_half = split[0]
        second_half = split[1]
        count = np.count_nonzero(second_half)
        
        #Calculate means
        mean_fh = st.mean(first_half)
        mean_sh = st.mean(second_half)


        #percent difference between mean of first_half and second_half
        p_diff = abs((mean_fh - mean_sh))/((mean_fh+mean_sh)/2)

        #Determine whether to stay split or not
        if(p_diff > 0.5):
            arr = second_half

    #descriptive statistics of new array
    mean_arr = st.mean(arr)
    median_arr = st.median(arr)
    stdev_arr = st.stdev(arr)
    variance_arr = st.variance(arr)

    #Determining new time standards
    if (time_stand*.7<mean_arr<time_stand*1.3):
        mean_arr = mean_arr
    elif (time_stand<mean_arr):
        time_stand = mean_arr*.8
    else:
        time_stand = mean_arr*1.2

    #New Time Standard
    print(round(time_stand,4))
    
    #if (i==1):
       # i+=2
    #else:
    #i+=1


# In[ ]:




