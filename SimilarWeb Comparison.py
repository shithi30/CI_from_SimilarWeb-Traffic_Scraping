#!/usr/bin/env python
# coding: utf-8

# import
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import duckdb
from googleapiclient.discovery import build
from google.oauth2 import service_account
import win32com.client
from pretty_html_table import build_table
import time
from datetime import datetime

## Similarweb ##

def scrape_datapoint(attr_ls, val_ls, attr_ind, attr_div, attr_cls, val_div, val_cls, cat, soup_init, p):
    
    # accumulator
    df_temp = pd.DataFrame()
    
    # attributes
    soup = soup_init.find_all(attr_div, attrs={"class": attr_cls})
    soup = eval('soup' + attr_ind)
    for s in soup: attr_ls.append(s.get_text())
        
    # values
    soup = soup_init.find_all(val_div, attrs={"class": val_cls})
    for s in soup: val_ls.append(s.get_text())
        
    # store
    df_temp['value'] = val_ls
    df_temp['attribute'] = attr_ls
    df_temp['category'] = cat
    df_temp['platform'] = p
    
    # return
    return df_temp

def scrape_similarweb(platforms):

    # accumulators
    start_time = time.time()
    df = pd.DataFrame()
    
    # preference
    options = webdriver.ChromeOptions()
    options.add_argument('ignore-certificate-errors')

    # open window
    for p in platforms:
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()
        
        # url
        url = 'https://www.similarweb.com/website/' + p + '/'
        driver.get(url)

        # load
        time.sleep(10)

        # scroll
        SCROLL_PAUSE_TIME = 3
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            time.sleep(SCROLL_PAUSE_TIME)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height: break
            last_height = new_height

        # soup
        soup_init = BeautifulSoup(driver.page_source, 'html.parser')
        
        # close window
        driver.close()

        # marketing channels        
        df = df._append(scrape_datapoint([], [], "[:]", "span", "wa-traffic-source-label__title", "tspan", "wa-traffic-sources__channels-data-label", 'marketing channel', soup_init, p))
        # age
        df = df._append(scrape_datapoint([], [], "[2]", "g", "highcharts-axis-labels highcharts-xaxis-labels", "tspan", "wa-demographics__age-data-label", 'age', soup_init, p))
        # intro
        df = df._append(scrape_datapoint([], [], "[:]", "dt", "app-company-info__list-item app-company-info__list-item--title", "dd", "app-company-info__list-item app-company-info__list-item--value", 'intro', soup_init, p))
        # ranks
        df = df._append(scrape_datapoint([], [], "[:]", "p", "wa-rank-list__title", "p", "wa-rank-list__value", 'rank', soup_init, p))
        # gender
        df = df._append(scrape_datapoint([], [], "[:]", "span", "wa-demographics__gender-legend-item-title", "span", "wa-demographics__gender-legend-item-value", 'gender', soup_init, p))
        # last 3 months' visit
        # df = df._append(scrape_datapoint(['Month - 3', 'Month - 2', 'Month - 1'], [], "[:]", "", "", "tspan", "wa-traffic__chart-data-label", 'visits last 3 months', soup_init, p))
        # last month traffic
        df = df._append(scrape_datapoint([], [], "[:]", "p", "engagement-list__item-name", "p", "engagement-list__item-value", 'last month', soup_init, p))
        # report
        print("Statistics scraped for: " + p)
        
    # cleaning
    qry = '''
    select distinct
        platform,
        category,
        attribute,
        value,
        case 
            when value like '%-%' then null
            when value like '#%' then replace(right(value, length(value)-1), ',', '')::float
            when right(value, 1)='%' then left(value, length(value)-1)::float/100
            when right(value, 1)='M' then left(value, length(value)-1)::float*1000000
            when right(value, 1)='K' then left(value, length(value)-1)::float*1000
            when value like '%:%' then 
                (string_split(value, ':')[1]::int*3600
                +string_split(value, ':')[2]::int*60
                +string_split(value, ':')[3]::int)::float
            when value~'^[0-9\.]+$' then value::float
            else null
        end value_cleaned,
        current_date::text report_date
    from df
    '''
    df = duckdb.query(qry).df().fillna('')
    
    # csv
    folder = r'C:\\Users\\shith\\Unilever Takeaway\\Unilever Codes\\Scraping Scripts\\'
    filename = folder + "similarweb_ecom_compare_data.csv"
    df.to_csv(filename, index=False)
    
    # analysis
    qry = '''
    select 
        platform,
        max(case when attribute='Total Visits' then value else null end) "visits last month",
        max(case when attribute='Avg Visit Duration' then value else null end) "avg. visit duration",
        max(case when attribute='Pages per Visit' then value else null end) "pages/visit",
        max(case when attribute='Female' then value else null end) "female visitors pct",
        max(case when attribute='25 - 34' then value else null end) "age 25 - 34 visitors pct",
        max(case when attribute='Category Rank' then value else null end) "category rank",
        max(case when attribute='Organic Search' then value else null end) "organic search marketing pct",
        max(report_date) "report date"
    from df
    group by 1
    '''
    res_df = duckdb.query(qry).df().fillna('')
    
    # update
    put_to_sheet(df, res_df)
    
    # stats
    print("\nTotal datapoints found: " + str(df.shape[0]))
    elapsed_time = str(round((time.time() - start_time) / 60.00, 2))
    print("Elapsed time to run script (mins): " + elapsed_time)
        
    # return
    return res_df

def put_to_sheet(data_df, summary_df):
    
    # credentials
    SERVICE_ACCOUNT_FILE = 'read-write-to-gsheet-apis-1-04f16c652b1e.json'
    SAMPLE_SPREADSHEET_ID = '1gkLRp59RyRw4UFds0-nNQhhWOaS4VFxtJ_Hgwg2x2A0'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # APIs
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # extract
    values = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='SimilarWeb!A1:F').execute().get('values', [])
    df_prev = pd.DataFrame(values[1:] , columns = values[0])
    values = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='SimilarWeb!H1:P').execute().get('values', [])
    df_sum_prev = pd.DataFrame(values[1:] , columns = values[0])
    
    # transform
    df_now = duckdb.query('''select * from (select * from df_prev where left(report_date::text, 7)!=left(current_date::text, 7) union all select * from data_df) tbl1 order by report_date desc, platform asc, category asc, attribute asc''').df().fillna('')
    df_sum_now = duckdb.query('''select * from (select * from df_sum_prev where left("report date", 7)!=left(current_date::text, 7) union all select * from summary_df) tbl1 order by "report date" desc, platform asc''').df().fillna('')
    
    # load
    res = sheet.values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='SimilarWeb').execute()
    res = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="'SimilarWeb'!A1", valueInputOption='USER_ENTERED', body={'values': [df_now.columns.values.tolist()] + df_now.values.tolist()}).execute()
    res = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="'SimilarWeb'!H1", valueInputOption='USER_ENTERED', body={'values': [df_sum_now.columns.values.tolist()] + df_sum_now.values.tolist()}).execute()

# websites
df = scrape_similarweb(['ushopbd.com', 'chaldal.com', 'shwapno.com', 'daraz.com.bd', 'foodpanda.com.bd', 'shajgoj.com', 'ohsogo.com'])
display(df)

# email
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = 'Similarweb Comparisons ' + time.strftime('%d-%b-%y')
newmail.To = "avra.barua@unilever.com; safa-e.nafee@unilever.com; rafid-al.mahmood@unilever.com; zoya.rashid@unilever.com; samsuddoha.nayeem@unilever.com"
newmail.BCC = "shithi30@outlook.com"

# body
newmail.HTMLbody = f'''
Dear concern,<br><br>
It may be important to analyze e-com platforms' comparative performance in order to maximize ROI. Hence, traffic and engagement data have been fetched from <a href="www.similarweb.com">similarweb.com</a>.
''' + build_table(df, 'green_light') + '''
Note that, the statistics presented are reflections of the mentioned website at the time of scraping. More parameters are available in the attachment. This is an auto email via <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, Cust. Service Excellence<br>
Unilever BD Ltd.<br>
'''

# attachment(s) 
folder = r'C:\\Users\\shith\\Unilever Takeaway\\Unilever Codes\\Scraping Scripts\\'
filename = folder + "similarweb_ecom_compare_data.csv"
newmail.Attachments.Add(filename)

# display
newmail.Send()

