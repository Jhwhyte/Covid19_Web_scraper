import pandas as pd
import requests
from bs4 import BeautifulSoup
import datetime

# get the response in the form of html
wikiurl = "https://en.wikipedia.org/wiki/Template:COVID-19_testing_by_country"
table_class = "wikitable sortable jquery-tablesorter"
response = requests.get(wikiurl)

# parse data from the html into a beautifulsoup object
soup = BeautifulSoup(response.text, 'html.parser')
indiatable = soup.find('table',{'class':"wikitable"})

# read in as a pandas wiki_dataframe
wiki_dataframe = pd.read_html(str(indiatable))

# convert list to dataframe
wiki_dataframe = pd.DataFrame(wiki_dataframe[0])

# drop the unwanted columns
refined_wiki_dataframe = wiki_dataframe.drop([ "Units[b]", "Confirmed(cases)","Ref.", "Confirmed /tested,%" , "Tested /population,%", "Confirmed /population,%" ], axis=1)

# rename the remaining desired columns
refined_wiki_dataframe = refined_wiki_dataframe.rename(columns = {"Country or region": "Country","Date[a]": "Date", "Tested": "Tested" })


Norway = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Norway']

Finland = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Finland']

Denmark = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Denmark[e]']

Sweden = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Sweden']

Iceland = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Iceland']

UK = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'United Kingdom']

US = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'United States']

Spain = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Spain']

Italy = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Italy']

France = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'France[f][g]']
           
Germany = refined_wiki_dataframe[refined_wiki_dataframe["Country"] == 'Germany']


concatenated_data_frames = pd.concat([Norway, Finland, Denmark, Sweden, Iceland,
UK, US, Spain, Italy, France, Germany])

# Converts the column from type Object to type Interger32
concatenated_data_frames["Tested"] = concatenated_data_frames["Tested"].astype(int)

# Takes that previously converted column and inserts a commas as a thousands separtor
concatenated_data_frames['Tested'] = concatenated_data_frames['Tested'].apply('{:,}'.format)

# change to file path to directory where python file is being run
out_path = "C:\\Users\\james\\Desktop\\CVD19\\covid19_PMT.xlsx"

concatenated_data_frames['Date'] = pd.to_datetime(concatenated_data_frames['Date'], errors='coerce')

writer = pd.ExcelWriter(out_path, date_format = 'dd-mm-yyyy', datetime_format='mm/dd/yyyy')

formated_for_excel_data = concatenated_data_frames.style.set_properties(**{'text-align': 'center'})

formated_for_excel_data.to_excel(writer, sheet_name='Sheet1',index=False)

workbook  = writer.book
worksheet = writer.sheets['Sheet1']

worksheet.set_column(0, 3, 25)
writer.save()

# starts excel automatically upon creation
import os
os.system(f'start excel.exe "{out_path}"')