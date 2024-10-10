import pandas as pd
import sqlalchemy as sa
import datetime as dt
import pythoncom
import win32com.client
import numpy as np


# build the connection string
connection_string = (
    'Driver=ODBC Driver 17 for SQL Server;'
    'Server=(localdb)\localdb1;'
    'Database=Stock;'
    'Trusted_Connection=yes;'
)
connection_url = sa.engine.URL.create(
    "mssql+pyodbc",
    query=dict(odbc_connect=connection_string)
)
engine = sa.create_engine(connection_url, fast_executemany=True)

# SheetNames=['MostActive','AllTimeHigh','PennyStocks','TopGainers']

# Instead of hardcoding the sheet name get it dynamically from from xl
File_path = r"C:\Users\kiran\OneDrive\Documents\Projects\Stock Projects\StockOverviewNew1.xlsx"

print("Refresh the excel sheet")
# RefreshXLSheet(File_path)
print("Done refreshing now read xl to dataframe")

sheet_to_df_map = pd.read_excel(File_path, sheet_name=None)
print("reading excel complete for fetching sheet names")

xls = pd.ExcelFile(File_path)
SheetNamesList = xls.sheet_names
print(SheetNamesList)

for SheetName in SheetNamesList:
    print('inside for loop for reading excel to dataframe')
    print(SheetName)
    # read the xl
    df = pd.read_excel(File_path, sheet_name=SheetName)
    print(f'read sheet{SheetName} into a dataframe')
    cleanData(df)
    print(f'data cleaned for sheet = {SheetName}')
    # print(df)
    # df.info()
    print('loaded info for the dataframe')
    print('number of rows in dataframe', len(df))
    # Add the date to the dataframe .this will append a new column called StockDate in df and all the rows gets todays date.
    # xl dont have column for date but you are dynamically adding in python .once you load the Df to sql it will add values for date column
    df['StockDate'] = pd.to_datetime('today').date()
    # added currency type new
    df['Currency'] = 'USD'

    # df['StockDate']='2024-02-26'

    # upload the DataFrame to sql
    df.to_sql(f"{SheetName}", engine, schema="dbo", if_exists="append", index=False)
    print('uploaded to sql')

print("outside for loop")
# When an Engine is garbage collected, its connection pool is no longer referred to by that Engine, and assuming
# none of its connections are still checked out, the pool and its connections will also be garbage collected
engine.dispose()
xls.close()
print("MySQL connection is closed")


def cleanData(df):
    print('Inside Clean data')
    print(df.head())
    # Repalce ' ' in column name with '_'
    df.columns = df.columns.str.replace(' ', '_')
    print('new title1')
    print(df.head())
    columnList = df.columns.tolist()
    for column in columnList:
        # values in columns has '-' in values replacing thme with null
        patt = '—'
        df[column] = df[column].astype(str).apply(lambda x: re.sub(patt, "", x))
        # Remove USD from values and create a seperate column as currency with value as USD
        patt = 'USD'
        df[column] = df[column].astype(str).apply(lambda x: x.rstrip('USD'))
        df[column] = df[column].astype(str).apply(lambda x: x.rstrip('%'))
        # print('converting column',column)
        # if column not in['Symbol','Sector','Analyst_Rating','Volume','Market_cap','Vol_*_Price','Change_%','EPS_dil_growth_TTM_YoY']:
        # df[column] = df[column].astype('Int64')

    print(df.dtypes)
    print(df.head())


def cleanData(df):
    print('Inside Clean data')
    print(df.head())
    # Repalce ' ' in column name with '_'
    df.columns = df.columns.str.replace(' ', '_')
    print('new title1')
    print(df.head())
    columnList = df.columns.tolist()
    for column in columnList:
        # values in columns has '-' in values replacing thme with null
        patt = '—'
        df[column] = df[column].astype(str).apply(lambda x: re.sub(patt, "", x))
        # Remove USD from values and create a seperate column as currency with value as USD
        patt = 'USD'
        df[column] = df[column].astype(str).apply(lambda x: x.rstrip('USD'))
        df[column] = df[column].astype(str).apply(lambda x: x.rstrip('%'))
        # print('converting column',column)
        # if column not in['Symbol','Sector','Analyst_Rating','Volume','Market_cap','Vol_*_Price','Change_%','EPS_dil_growth_TTM_YoY']:
        # df[column] = df[column].astype('Int64')

    print(df.dtypes)
    print(df.head())
