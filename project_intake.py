# -*- coding: utf-8 -*-
"""
Created on Wed May 19 13:37:18 2021

@author: jbackes
"""

import pandas as pd
import os
import shutil




file_path = 'M:/jbackes/pytools/project_intake/'
file_name = 'project_file_list.xlsx'


from time import sleep
from datetime import datetime

today = datetime.now()

# dd/mm/YY
d1 = today.strftime("%Y-%m-%d %H:%M:%S")
d2 = today.strftime("%Y%m%d_%H%M%S")

#import math
#import openpyxl 

def create_file():
    data.to_excel(f'{file_path}{file_name}')
    

def check_open():
    try: 
        os.rename(f'{file_path}{file_name}1.xlsx', f'{file_path}{file_name}')
    except OSError:
        print('File is open by another user. Please try again later')

def read_excel():
    df = pd.read_excel(f'{file_path}{file_name}', converters={'Number': lambda x: str.zfill(x,4)})
    return df

def create_project_num():
    project_list = list(data['Number'])
    max(project_list)
    new_project_num = str.zfill(str(int(max(project_list)) +1),4)
    return new_project_num

data = read_excel()
project_num = create_project_num()


while True:
    input_type = input("Do you want to: a) Create a new project. b) List all projects. [a/b]? : ")
    
    if input_type.lower() not in ['a','b','0']:
        exit_input = input("please use answer 'a' or 'b'. Press 'enter' to continue or 0 to exit :")
        if exit_input.lower() == '0':
            break
        else:
            continue
        
    # if 'a' to create a new project. Program will check the status of the Excel file,
    # then create a txt file to let other user know not to open it because it is being edited.
    if input_type.lower() == 'a':
        try: 
            data = pd.read_excel(f'{file_path}{file_name}', converters={'Number': lambda x: str.zfill(x,4)})
            project_list = list(data['Number'])
            new_project_num = str.zfill(str(int(max(project_list)) +1),4)
            shutil.move(f'{file_path}{file_name}', f'{file_path}/backup/{d2}_{file_name}')

        except OSError:
            print('File is open by another user. Please try again later')
            break
        try:
            print('Creating new project')
            sleep(1)
            print('')
            print(f'Your project number is going to be {new_project_num}')
            sleep(1)
            print('')
            Name = input('Please enter name of project : ')
            Team = input('Please enter Team :')
            Creator = input('Please enter your name as creator :')
            Owner = input('Please enter project owner :')
            d = {'Number':new_project_num,'Name':Name, 'Team': Team, 'Creator': Creator, 'Owner':Owner, 'Date':d1}
            df = pd.DataFrame(data=d, index=[0])
            
            new_df = data.append(df)  
            new_df.to_excel(f'{file_path}{file_name}', index=False)
            break
        except:
            print("Oh no!, something didn't work, exiting now")
            break
    if input_type.lower() == 'b':
        print('Listing all projects')
        sleep(1)
        print(data)
        create_input = input("Do you want to create a new project [y/n]? : ")
        if create_input.lower() not in ['y','n']:
            exit_input = input("please use answer 'y' or 'n'. Press 'enter' to continue or '0' to exit :")
        if exit_input.lower() == '0':
            break
        elif exit_input.lower() == 'y':
            continue
        elif exit_input.lower() == 'n':
            continue
    if input_type.lower() == '0':
        print('Exiting')
        break
  
    
os.remove(f'{file_path}{file_name}')
shutil.move(f'{file_path}{file_name}', f'{file_path}{file_name}')