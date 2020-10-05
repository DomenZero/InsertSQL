# author Merkulov Maksim
# email: merkulovmx@gmail.com

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

def folder_path_1():
    return r"\\Network Shared Folder\\folder_path_1";
def folder_path_2():
    return r"\\Network Shared Folder\\folder_path_2";
def folder_path_3():
    return r"\\Network Shared Folder\\folder_path_3";
def folder_path_4():
    return r"\\Network Shared Folder\\folder_path_4";
def folder_path_5():
    return r"\\Network Shared Folder\\folder_path_5";
def folder_path_6():
    return r"\\Network Shared Folder\\folder_path_6";
def folder_path_7():
    return r"\\Network Shared Folder\\folder_path_7";
def folder_path_8():
    return r"\\Network Shared Folder\\folder_path_8";

def folder_path_9():
    return r"\\Network Shared Folder\\folder_path_9";
def folder_path_10():
    return r"\\Network Shared Folder\\folder_path_10";
def folder_path_11():
    return r"\\Network Shared Folder\\folder_path_11";
def folder_path_12():
    return r"\\Network Shared Folder\\folder_path_12";
def folder_path_13():
    return r"\\Network Shared Folder\\folder_path_13";

def search_value_and_compare(excel_dataframe, path_search, index_row, name_column):
    if (re.search(path_search, excel_dataframe.at[index_row, name_column],
                  flags=re.IGNORECASE)) != None:
        return 1
    else:
        return 0

def sub_value_for_number(excel_dataframe, index_row, name_column):
    if (re.search(finance(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 1
    elif (re.search(hanil(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 2
    elif (re.search(hs(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 3
    elif (re.search(hr(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 4
    elif (re.search(it(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 5
    elif (re.search(logistic(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 6
    elif (re.search(maintenance(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 7
    elif (re.search(management(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 8
    elif (re.search(production(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 9
    elif (re.search(projects(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 10
    elif (re.search(quality(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 11
    elif (re.search(togliatti(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 12
    elif (re.search(tidoc(), excel_dataframe.at[index_row, name_column],
                   flags=re.IGNORECASE)) != None:
        return 13
    else:
        return 0

def generate_random():
    rand_value = randint(1000000000, 2147483643)
    return rand_value

def correct_bool(excel_dataframe, index, column):
    # print(excel_dataframe.at[index, column])
    if (bool(excel_dataframe.at[index, column])==1):
        return 1
    else:
        return 0
    # excel_dataframe.at[index, column]


def createParser():
    parser = argparse.ArgumentParser()
    parser.add_argument ('-ps', '--ps')
    parser.add_argument ('-f', '--file')
    parser.add_argument('-end', '--end', type=int, default=1)

    return parser

if __name__ == "__main__":

    # ps = str(sys.argv[1])
    parser = createParser()
    console_args = parser.parse_args(sys.argv[1:])
    print(console_args)
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=SERVER_NAME;DATABASE=DATABASE_NAME;TrustedConnection=yes;UID=User_ID;PWD=' + str(console_args.ps)+'')
    cursor = conn.cursor()

    # Read data from table
    sql_query = pd.read_sql_query('SELECT * FROM DATABASE_NAME.db_owner.results_load', conn)
    print(sql_query)
    print(type(sql_query))

    file_name = str(console_args.file)  # change it to the name of your excel file
    excel_dataframe = read_excel(file_name, sheet_name=0, header=0)
    len_strings = len(excel_dataframe.index)
    rand_value = generate_random()
    now = datetime.datetime.now()
    date_time_value = now.strftime("%Y-%m-%d %H:%M:%S")
    i = 0
    # test insert 1 element
    print(bool(excel_dataframe.at[i, 'Stat']))
    print(bool(re.search("TRUE", str(excel_dataframe.at[i, 'Stat']))))

    # Beautiful write in database
    # NEW LOAD
    i = 0
    # iterating over indices
    for col in excel_dataframe.index:
        # print(sub_value_for_number(excel_dataframe, i, 'File link'))
        cursor.execute(
            "INSERT INTO real_retentionpolicy.db_owner.results_load([id],[date],[path],[stat],[department]) values (?, ?, ?,?,?)",
            rand_value, date_time_value, excel_dataframe.at[i, 'File link'],
            correct_bool(excel_dataframe, i, 'Stat'),
            sub_value_for_number(excel_dataframe, i, 'File link'))
        i += 1
    # END NEW LOAD
    conn.commit()
    conn.close()
