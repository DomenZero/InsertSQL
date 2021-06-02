# Author Maksim Merkulov
import pandas as pd
import pyodbc
from pandas import read_excel
import re
from openpyxl import Workbook
from random import randint
import datetime
import csv
import sys
import argparse

def search_value_and_compare(excel_dataframe, path_search, index_row, name_column):
    if (re.search(path_search, excel_dataframe.at[index_row, name_column],
                  flags=re.IGNORECASE)) != None:
        return 1
    else:
        return 0

def generate_random():
    rand_value = randint(1000000000, 2147483643)
    return rand_value

def correct_bool(excel_dataframe, index, column):
    if (bool(excel_dataframe.at[index, column])==1):
        return 1
    else:
        return 0


def createParser():
    parser = argparse.ArgumentParser()
    parser.add_argument ('-ps', '--ps')
    parser.add_argument ('-f', '--file')
    parser.add_argument('-end', '--end', type=int, default=1)

    return parser

if __name__ == "__main__":

    parser = createParser()
    console_args = parser.parse_args(sys.argv[1:])
    print(console_args)
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=server_name;DATABASE=database_name;TrustedConnection=yes;UID=user_name;PWD=' + str(console_args.ps))
    cursor = conn.cursor()

    # Read data from table
    sql_query = pd.read_sql_query('SELECT * FROM database_name.db_name.results', conn)
    print(sql_query)
    print(type(sql_query))

    file_name = str(console_args.file)  # change it to the name of your excel file
    excel_dataframe = read_excel(file_name, sheet_name=0, header=0)
    len_strings = len(excel_dataframe.index)
    rand_value = generate_random()
    now = datetime.datetime.now()
    date_time_value = now.strftime("%Y-%m-%d %H:%M:%S")
    i = 0
    print(bool(excel_dataframe.at[i, 'Stat']))
    print(bool(re.search("TRUE", str(excel_dataframe.at[i, 'Stat']))))

   
    i=0
    # iterating over indixes
    for col in excel_dataframe.index:
        cursor.execute(
            "INSERT INTO database_name.db_name.results([id],[date],[path],[stat],[department1],[department2],[department3],[department4],[department5],[department6],department7],[department8],[department9],[department10],[department11],[department12],[department13]) values (?, ?, ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            rand_value, date_time_value, excel_dataframe.at[i, 'Table_column_name'],
            correct_bool(excel_dataframe, i, 'Stat'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department1", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department2", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department3", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department4", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department5", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department6", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department7", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department8", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department9", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department10", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department11", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department12", i, 'Table_column_name'),
            search_value_and_compare(excel_dataframe, r"\\Local_folder\department13", i, 'Table_column_name'))
        i += 1

       
    conn.commit()
    conn.close()
