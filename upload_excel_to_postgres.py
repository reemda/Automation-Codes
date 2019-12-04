# Import libraries
import pandas as pd
import os
import pandas.io.sql as psql
from sqlalchemy import create_engine
from datetime import datetime


# url is the combination of the folder path and the sheet index inside the folder 
path = "C:\\Users\\reemc\\Desktop\\folder"  
folder = os.listdir(path)
target_sheet= 0  
url = path+"\\"+folder[target_sheet]

# Read excel sheet
xl= pd.io.excel.read_excel(url,sheet_name='Sheet1', index_col=None,na_values=['na'])

# Clean up the sheet as needed
xl["day"]=xl["Meeting"].str.split(',').str[0]
xl["month_day"]=xl["Meeting"].str.split(',').str[1]
xl["year"]=xl["Meeting"].str.split(',').str[2]
xl['in_person_meeting'] =xl['month_day'] +' '+ xl['year']
xl['in_person_meeting'] =pd.to_datetime (xl['in_person_meeting'] )
xl.drop(['Meeting', 'day','month_day','year'], axis=1,inplace=True)
xl.rename(columns={'Manager Name':'manager_name',
                      'Account Name':'account_name',
                      '# In-Person Meetings YTD':'in_person_meetings_ytd',
                       '# Active Opportunities':'active_opportunities',
                       '$ Active Opps Est Revenue':'active_opps_est_revenue',
                        '# Closed Opportunities':'closed_opportunities',
                        '$ Closed Opps Est Revenue':'closed_opps_est_revenue',
                        'year_month':'year_month',
                        'revenue':'revenue',
                        'Key':'key_',
                        'last_year_revenue' :'last_year_revenue',
                        'Close Rev':'close_revenue' },inplace=True)
                        
# Create engine and upload to an existing table 
engine = create_engine('postgresql://user:password@database server:database port/database name')
xl.to_sql('table name',engine,if_exists='replace',index=False)
