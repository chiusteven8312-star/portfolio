#===========================================================================================
# Sample script to extract, load, and transform data
# Created on October 24, 2025
# Author: Steven Chiu
#===========================================================================================

"""
Importing Libraries
"""

import pandas as pd
import os
import datetime
from datetime import datetime
from datetime import today
import traceback
import win32com.client
from sqlaclchemy import create_engine
import logging
import time

"""
Defining Variables
"""

script_name = 'example_trade_calculations'
today_date = datetime.datetime.strftime(today, '%m/%d/%Y')
database_cnxn = create_engine('mssql+pyodbc://databaseurl/sampledatabasename?driver=ODBC+Driver+Sql+Servber+Example')
email_to = 'chiusteven8312@outlook.com'
CodeType = 'daily calculations and load'
db = 'Example DB'
db_table = 'example_tbl'
folder_main = os.path.dirname(os.path.abspath(__file__))

ex_sql_script = f'''
            selecet * from example_db_tbl where date_trade_confirmed = '{today_date}
            '''

"""
Functions
"""

def sucessEmail():
    ol = win32com.client.dynamic.Dispatch("outlook.application")
    olmailitem = 0x0 #size of new email
    newmail = ol.CreateItem(olmailitem)
    newmail.To = email_to
    newmail.Subject = f"{CodeType} - {script_name} Complete"
    newmail.HTMLBody = f'''Daily {script_name} was successfully loaded to {db} into below table:<br><br>
                                <b>[{db_table}]</b><br>
                                '''
    newmail.Send()

def failEmail(errmsg):
    ol = win32com.client.dynamic.Dispatch("outlook.application")
    olmailitem = 0x0 #size of new email
    newmail = ol.CreateItem(olmailitem)
    newmail.To = f'{os.getlogin()}@outlook.com'
    newmail.Subject = f'Python script: {script_name} failed'
    newmail.Body = "Exception Error: " + str(errmsg)
    newmail.Importance = 2
    #newmail.Display()
    newmail.Send()

def calculate_daily_pnl(df):
    """
    Calculate daily profit and loss (P&L) for eachg trade.
    Generic columns: trade_price, market_price, quantity, side (B/S)
    """
    df = df.copy()
    df['direction'] = df['side'].map({'B': 1, 'S': -1})
    df['pnl'] = (df['market_price'] - df['trade_price']) * df['quantity'] * df['direction']
    return df

def calculate_trade_volume_notional(df):
    """
    Calculate trade volume and notional value for each trade.
    Generic columns: trade_price, quantity
    """
    df = df.copy()
    df['notional'] = df['trade_price'] * df['quantity']
    return df

def calculate_position_exposure(df):
    """
    Calculate end-of-day position and exposure for each security/account.
    Generic columns: quantitym, side (B/S), account_id, security_id
    """
    df = df.copy()
    df['direction'] = df['side'].map({'B': 1, 'S': -1})
    df['signed_quantity'] = df['quantity'] * df['direction']
    # Group by account and security to get net position
    position = df.groupby(['account_id', 'security_id'])['signed_quantity'].sum().reset_index()
    position = position.rename(columns = {'signed_quantity': 'net_position'})
    return position

def calculate_weighted_average_price(df):
    """
    Calculate weighted average price (WAP) for each security/account.
    Generic columns: trade_price, quantity, account_id, security_id
    """
    df = df.copy()
    wap = df.groupby(['account_id', 'security_id']).apply(
        lambda x: pd.Series({
            'weighted_avg_price': (x['trade_price'] * x['quantity']).sum() / x['quantity'].sum()
        })
    ).reset_index()
    return wap

if __name__ == '__main__':
    try:
        logging.info(f'----------------------------------------Beginning {script_name}----------------------------------------------')
        while True:
            logging.info("Checking for data availability")
            today_date_check = pd.read_sql(('SELECT max(date_confirmed) FROM example_db_tbl'), con = database_cnxn).iloc[0, 0].strftime('%m/%d/%Y')

            if today_date_check == today_date:
                logging.info("Reading in Trade Data")
                trade_df = pd.read_sql(ex_sql_script, con = database_cnxn)

                """
                Calculating Daily PnL
                """
                logging.info('Beginning PnL Calculation Process')
                pnl_df = trade_df[['trade_id', 'account_id', 'security_id', 'trade_price', 'market_price', 'quantity', 'side']]
                pnl_result = calculate_daily_pnl(pnl_df)

                """
                Calculating Notional Trade Volume
                """
                logging.info("Beginning Notional Trade Volume Calculation Process")
                volume_df = trade_df[['trade_id', 'account_id', 'security_id', 'trade_price', 'quantity']]
                volume_results = calculate_trade_volume_notional(volume_df)

                """
                Calculating Position Exposure
                """
                logging.info("Beginning Position Exposure Calculation Process")
                position_df = trade_df[['trade_id', 'account_id', 'security_id', 'quantity', 'side']]
                position_results = calculate_position_exposure(position_df)

                """
                Calculating Weighted Average Price
                """
                logging.info("Beginning Weighted Average Price Calculation Process")
                wap_df = trade_df[['trade_id', 'account_id', 'security_id', 'trade_price', 'quantity']]
                wap_results = calculate_weighted_average_price(wap_df)

                logging.info("Uploading to Respective Tables in Database")
                with database_cnxn.connect() as conn:
                    pnl_result.to_sql(name = 'example_pnl_results_table', schema = 'schema_ex', con = database_cnxn, if_exists = 'append', index = False)
                    volume_results.to_sql(name = 'example_volume_results_table', schema = 'schema_ex', con = databse_cnxn, if_exists = 'append', index = False)
                    position_results.to_sql(name = 'example_position_results_table', schema = 'schema_ex', con = database_cnxn, if_exists = 'append', index = False)
                    wap_results.to_sql(name = 'example_wap_results_table', schema = 'schema_ex', con = database_cnxn, if_exists = 'append', index = False)

                sucessEmail()
                break
            else:
                logging.info("Waiting for Trade Information to Upload to Database for Today")
                time.sleep(300)

        logging.info(f'---------------------------------------- {script_name} Complete----------------------------------------------')
    
    except Exception as e:
        errmsg = traceback.format_exc()
        failEmail(errmsg)
        logging.error(errmsg)