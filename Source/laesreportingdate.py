# -*- coding: utf-8 -*-
"""
Created on Tue Nov 14 09:33:45 2023

@author: B006207
"""
from os.path import join
from os import listdir
import pandas as pd

#Ignore UserWarning fra Openpyxl vedr. Data-validation
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

#Opsaetning og import af egne moduler
import sys
sys.path.append(r'C:\Filkassen\PythonMm\VSCode_projects\LIFEOpsaetning')

import partneroversigt as po



def hent_fr(mappe_path):
    suffix = '.xlsx'
    file_list = listdir(mappe_path)
    xlsx_list = []
    for file in file_list:
        if file.endswith(suffix):
            xlsx_list.append(join(mappe_path, file))
    return xlsx_list


def hent_ult_periode(file_path):
    df = pd.read_excel(file_path, sheet_name='Individual Cost Statement', header=None)
    ult_celle = (3, 4)
    return df.iloc[ult_celle]
    


def oversigt(project):
    for folder in project.distspec.values():
        mappe_path = join(project.pathroot, project.pathbase, folder, project.pathspec)
        xlsx_list = hent_fr(mappe_path)
        print(folder)
        for file in xlsx_list:
            print(f'****{hent_ult_periode(file)}')


def create_oversigt(pr):
    match pr.lower():
        case "nat":
            project = po.Natureman()
        case "open":
            project = po.Openwood()
        case "for":
            project = po.Forfit()
        case _:
            print("Forkert angivelse af projekt: Gyldige arg.: Nat - Open - For")
            return
        
    oversigt(project)


if __name__ == '__main__':
    create_oversigt('open')
