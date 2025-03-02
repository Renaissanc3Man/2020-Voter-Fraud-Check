import numpy as np
import pandas as pd
import re
import datetime
import traceback
import time
import os

##import urllib2
##from bs4 import BeautifulSoup

import win32com.client









today = datetime.datetime.now().strftime('%Y-%m-%d')


def get_time_of_file_update(myfile):
    return pd.to_datetime(datetime.datetime.fromtimestamp(os.path.getmtime(myfile)))


def to_newexcel(df,index=False):
    xlApp = win32com.client.Dispatch("Excel.Application")
    xlApp.Visible = True
    xlApp.DisplayAlerts = False
    xlApp.Application.CutCopyMode = False



    df.to_clipboard(index=index)
    unencrypted_workbook = xlApp.Workbooks.Add()
    unencrypted_workbook.Worksheets(1).Activate()
    unencrypted_workbook.Worksheets(1).Cells.NumberFormat = "@"
    unencrypted_workbook.Worksheets(1).Cells(1,1).Select()
    xlApp.Selection.PasteSpecial() #pastespecial to a new page

    xlApp.DisplayAlerts = True



#for later use in replacing misspelled names
def replace_if_in_dict(parm,mydict):
    if not pd.isnull(parm) and parm in list(mydict.keys()):
        parm = mydict[parm]
    return parm


def pivot_concat_helper_function(mygroup,delimiter='__'):
    mygroup = mygroup.drop_duplicates()
    if len(mygroup.dropna()) == 0:
        result = np.nan
    elif len(mygroup) == 1:
        result = mygroup.iloc[0]
    else:
        mygroup = mygroup.map(lambda x: str(x))
        result = delimiter.join(list(mygroup))
    return result


def safe_reset_index(mydataframe):
    try:
        mydataframe = mydataframe.reset_index()
    except:
        pass
    for mycol in ['index','level_0','level_1','']:
        try:
            del mydataframe[mycol]
        except:
            pass
    return mydataframe


#makes two column dataframe out of dict
def pd_dict_to_df(mydict,column_names):
    df = pd.DataFrame(pd.Series(mydict))
    df.columns = [column_names[1]]
    df[column_names[0]] = df.index
    df = safe_reset_index(df)
    df = df[column_names] #put in order
    return df


def fillna_from_other_col(df,from_col,to_col):
    original_df_cols = list(df.columns)
    cols_minus_to_col = list(df.columns)
    cols_minus_to_col.remove(to_col)
    df = pd.concat([df[cols_minus_to_col],df[[to_col]].apply(lambda x: x.fillna(value=df[from_col]))],axis=1)
    df = df[original_df_cols] #preserve column order
    return df


#get rid of double spaces, extra spaces on end, and strage '=-' characters
def clean_fighter_name(x):
    if not pd.isnull(x):
        try:
            x = str(x)
        except:
            x = x.encode('ascii','ignore')
        x = x.strip()
        x = x.replace('  ',' ')
        x = x.replace(r'=-','')
    return x



#between all of the data scraped from different sources, there are many instances
#of the same fighter being referred to as several different names
#take the file mma_data_fighter_names and create master dictionary to
#change all alternative spellings to master fighter name

#fighter_name_standardization_dict = create_fighter_name_standardization_dict()
#mma_df['Fighter'] = mma_df['Fighter'].map(lambda x: replace_if_in_dict(x,fighter_name_standardization_dict))
def create_fighter_name_standardization_dict():
    fighters_df = pd.read_csv(r'../input data/mma_data_fighter_name_standardization.csv',low_memory=False)
    del fighters_df['Last Name'] #only for use when filling out spreadsheet


    #iterate through each fighter to create dict
    fighter_name_standardization_dict = {}
    for i in range(0,fighters_df.shape[0]):
        myfighter_row = fighters_df.iloc[i]
        myfighter_row = myfighter_row.dropna()

        myfighter = myfighter_row['Fighter']
        myfighter_alternatives = list(myfighter_row.drop('Fighter').values)

        for myalternative in myfighter_alternatives:
            try:
                fighter_name_standardization_dict.update({myalternative:myfighter})
            except:
                #i forsee errors where doubling of alternate names breaks dict
                print(myfighter, myalternative)
                print(str(traceback.format_exc()))
                pass

    return fighter_name_standardization_dict






weightclass_dict = {'Catch Weight':0,
                    'Super Heavyweight':0,
                    'Open Weight':0,
                    'Women\'s Strawweight':115,
                    'Women\'s Flyweight':125,
                    'Flyweight':125,
                    'Bantamweight':135,
                    'Women\'s Bantamweight':135,
                    'Featherweight':145,
                    'Women\'s Featherweight':145,
                    'Lightweight':155,
                    'Welterweight':170,
                    'Middleweight':185,
                    'Light Heavyweight':205,
                    'Heavyweight':265}

#convert Division to Weightclass
def division_to_weightclass(x):
    result = np.nan
    if not pd.isnull(x):
        x = x.encode('ascii','ignore')
        if x in weightclass_dict.keys():
            result = weightclass_dict[x]
        elif not pd.isnull(re.match(r'(?i)(Catchweight)',x)):
            numeric_value_check = re.search(r'\d+\.\d+|\d+',x)
            if not pd.isnull(numeric_value_check):
                result = np.round(float(numeric_value_check.group(0)),1)
    return result























# ##############################################################################
# Standardize MMA Event Names
# MMAFighting, Wikipedia, UFC, Sherdog, Fightmetric
def standardize_event_titles(df, data_source = 'Wikipedia'):
    event_standardization_df = pd.read_csv(r'../input data/mma_data_event_standardization.csv',low_memory=False)

    event_standardization_df['Date'] = pd.to_datetime(event_standardization_df['Date'])
    event_standardization_df['Date '+data_source] = pd.to_datetime(event_standardization_df['Date '+data_source])


    event_standardization_df = event_standardization_df[['Date','Event','Date '+data_source,'Event '+data_source]]
    event_standardization_df = event_standardization_df.dropna(subset=['Date '+data_source,'Event '+data_source])

    #change names, join, then fillna from other col
    event_standardization_df = event_standardization_df.rename(columns={'Date':'Date_standardized',
                                                                        'Event':'Event_standardized',
                                                                        'Date '+data_source:'Date_temp2',
                                                                        'Event '+data_source: 'Event_temp2'})

    df = df.rename(columns={'Date':'Date_temp1','Event':'Event_temp1'})

    df = df.merge(event_standardization_df, left_on = ['Date_temp1','Event_temp1'], right_on = ['Date_temp2','Event_temp2'],how='left')

##    #debug
##    df['d1'] = df['Date_temp1'].shift(-1) == df['Date_temp1']
##    df['d2'] = df['Date_temp1'].shift(1) == df['Date_temp1']
##    to_newexcel(df)

    df = fillna_from_other_col(df,'Date_temp1','Date_standardized')
    df = fillna_from_other_col(df,'Event_temp1','Event_standardized')



    df = df.rename(columns={'Date_standardized':'Date' , 'Event_standardized':'Event'})

    del df['Date_temp1']
    del df['Event_temp1']
    del df['Date_temp2']
    del df['Event_temp2']

    return df









#get rid of events that are non-ufc
def filter_non_ufc_event(df):
    df = df[~(df['Event'].fillna('nan').str.contains(r'Bellator|WSOF|Strikeforce|Mayweather|Invicta|PFL|Professional Fighters League|WEC|Pacquiao|World Series of Fighting|Moosin|Canelo|GGG|Tuesday'))]
    return df





#generate layoff between fights
#df = df.groupby(['Fighter']).apply(determine_layoff)
def determine_layoff(df):
    df = df.sort_values(['Date'],ascending=[False])
    df = df.reset_index()
    df['Layoff'] = df['Date'] - df['Date'].shift(-1)
    df['Layoff'] = df['Layoff'].map(lambda x: int(x.days) if not pd.isnull(x) else x)
    return df







def make_date_array(start_date,end_date,strftime_format='%Y_%m_%d'):
    date_array = [pd.to_datetime(start_date + np.timedelta64(i,'D')).strftime(strftime_format) for i in range(0,int((end_date-start_date)/ np.timedelta64(1, 'D'))+1)]
    return date_array



def make_dict_from_cols(mydataframe,key_col,value_col):
    mydataframe = mydataframe[[key_col,value_col]]
    mydataframe = mydataframe.set_index(key_col)
    myseries = mydataframe[value_col]
    mydict = myseries.to_dict()
    return mydict


def replace_if_in_dict(parm,mydict):
    if not pd.isnull(parm) and parm in list(mydict.keys()):
        parm = mydict[parm]
    return parm