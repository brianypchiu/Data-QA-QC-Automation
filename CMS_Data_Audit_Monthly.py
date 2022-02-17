#!/usr/bin/env python
# coding: utf-8

# In[1]:


import shutil
import pandas as pd
from pandas import DataFrame

#Copy and paste CMS file from Ogee's folder to my folder and create a cvs file.

def CMS_copy_to_csv():
    src_path = r"H:\Consent Decree Work\Cleaning and Inspections (Ogee)\CMS_Export_CD.xlsx"
    dst_path = r"H:\Brian\Python Script\CMS_Monthly Check\CMS_Export_CD_monthlycheck_Brian.xlsx"

    shutil.copy(src_path,dst_path)
    print('copied')

    read_file = pd.read_excel (r"H:\Brian\Python Script\CMS_Monthly Check\CMS_Export_CD_monthlycheck_Brian.xlsx")
    read_file.to_csv (r"H:\Brian\Python Script\CMS_Monthly Check\CMS_CD_Monthlycheck_Brian.csv", index = None, header=True)
    
CMS_copy_to_csv()

#Import csv file to DataFrame
#Filter contract number "WW5200"
#Find duplicate based on Upstream-Downstream-WorkOrderNo-ItemNo


def find_CMSduplicate():
    
    df_CMS = pd.DataFrame(pd.read_csv(r"H:\Brian\Python Script\CMS_Monthly Check\CMS_CD_Monthlycheck_Brian.csv", sep=',', error_bad_lines=False, index_col=False, dtype='unicode'))
    df_new = df_CMS.loc[df_CMS['FILE_NO'] == 'WW5200-02']
    df_new = df_new.copy()
    df_new['UPS_DNS'] = df_new['UPSTREAM'].str.cat(df_new[['DOWNSTREAM', 'WORKORDER_NO','ITEM_NO']], sep='-')
    df_new['UPS_DNS'] = df_new['UPS_DNS'].str.replace(' ','')
    df_duplicated = df_new[df_new.duplicated(subset=['UPS_DNS'],keep='first')].dropna()
    df_duplicated.to_csv(r"H:\Brian\Python Script\CMS_Monthly Check\CMS_Export_CD_monthlycheck_result.csv")
    print('Audited. Please check file.')

find_CMSduplicate()

