"""
#######################################################################################################################
###											L_R loader											###
###																													###
###		Code developed by Negar Shams (negar.shams@PSCconsulting.com, +44 7436 544893) as part of PSC 		 		###
###		project JU8149 - Orsted - Ocean Wind																###
###																													###
#######################################################################################################################
"""

# Generic Imports
# from __future__ import division
from operator import truediv, floordiv
import pandas as pd
import numpy as np
# Unique imports
import common_functions as common
import constants as constants
import collections
# import load_est
# import load_est.psse as psse
import math
import time
import os
import dill
import csv
import plotter as plotter

import sys

sys.path.append(r"C:\DigSILENT15p1p7\python")

import os
import sys
import subprocess

user_name = 'NegarShams'
protection = 0
harmonics = 1
arcflash = 0

pf_version = '2020'

DIG_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5'
DIG_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP5'
DIG_PATH_2019 = r'C:\Program Files\DIgSILENT\PowerFactory 2019'
DIG_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020'
DIG_PYTHON_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5\Python\3.5'
DIG_PYTHON_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP7\Python\3.5'
DIG_PYTHON_PATH_2019 = r'C:\Program Files\DIgSILENT\PowerFactory 2019\Python\3.5'
DIG_PYTHON_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020\Python\3.8'
DIG_PYTHON_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020\Python\3.8'


if pf_version == '2016':
    DIG_PATH = DIG_PATH_2016
    DIG_PYTHON_PATH = DIG_PYTHON_PATH_2016
elif pf_version == '2018':
    DIG_PATH = DIG_PATH_2018
    DIG_PYTHON_PATH = DIG_PYTHON_PATH_2018
elif pf_version == '2019':
    DIG_PATH = DIG_PATH_2019
    DIG_PYTHON_PATH = DIG_PYTHON_PATH_2019
elif pf_version == '2020':
    DIG_PATH = DIG_PATH_2020
    DIG_PYTHON_PATH = DIG_PYTHON_PATH_2020
else:
    print('ERROR python version not found')
    raise SyntaxError('ERROR')

sys.path.append(DIG_PATH)
# #sys.path.append(DIG_PATH_2018)
sys.path.append(DIG_PYTHON_PATH)
# #sys.path.append(DIG_PYTHON_PATH_2018)

os.environ['PATH'] = os.environ['PATH'] + ';' + DIG_PATH
# noqa
import powerfactory

def loader(scenario_list,project):
    """
            Function returns dataframe with mapped values.
        """
    sc_num=len(scenario_list)
    #import powerfactory
    app = powerfactory.GetApplication()

    print('PowerFactory Opened')


    # activate project
    project = app.ActivateProject(project)
    prj = app.GetActiveProject()

    for i in range(0,sc_num):
        scenario_name=scenario_list[i]
        size = len(scenario_name)
        # # Slice string to remove last 3 characters from string
        new_str = scenario_name[:size - 2]
        excel_file_name=new_str+'_points.xlsx'
        excel_file_path = common.get_local_file_path_withfolder(file_name=excel_file_name, folder_name='Output_excel')
        points_df = common.import_excel(excel_file_path)
        L_points_I=points_df.loc[:,'worst_I_L']
        R_points_I=points_df.loc[:,'worst_I_R']
        L_points_V = points_df.loc[:, 'worst_V_L']
        R_points_V = points_df.loc[:, 'worst_V_R']
        L_I=L_points_I.tolist()
        R_I = R_points_I.tolist()
        L_V = L_points_V.tolist()
        R_V = R_points_V.tolist()

        k=1

        folders = prj.GetContents('*.IntPrjfolder')
        temp_filtered = filter(lambda folders: folders.loc_name == 'Study Cases', folders)
        studyCases_folder = list(temp_filtered)
        studyCases = studyCases_folder[0].GetContents('*.IntCase')
        temp_filtered = filter(lambda folders: folders.loc_name == 'Library', folders)
        Library = list(temp_filtered)
        Library_temp = Library[0]
        Library_subfolders = Library_temp.GetContents('*.')
        temp_filtered = filter(lambda Library_subfolders: Library_subfolders.loc_name == 'Operational Library',
                               Library_subfolders)
        oper_folder = list(temp_filtered)
        oper_temp = oper_folder[0]
        oper_subfolders = oper_temp.GetContents('*.')
        temp_filtered = filter(lambda oper_subfolders: oper_subfolders.loc_name == 'Characteristics',
                               oper_subfolders)
        char_folder = list(temp_filtered)
        char_temp = char_folder[0]
        char_subfolders = char_temp.GetContents('*.')

        # temp_filtered = filter(lambda char_subfolders: char_subfolders.loc_name == 'Oyster Creek 800MW',
        #                        char_subfolders)

        temp_filtered = filter(lambda char_subfolders: char_subfolders.loc_name == 'BL England 400MW',
                               char_subfolders)

        # temp_filtered = filter(lambda char_subfolders: char_subfolders.loc_name == 'BL England 400MW_PreFilter',
        #                        char_subfolders)

        scen_folder = list(temp_filtered)
        scen_temp = scen_folder[0]
        scen_subfolders = scen_temp.GetContents('*.')
        temp_filtered = filter(lambda scen_subfolders: scen_subfolders.loc_name == scenario_name, scen_subfolders)
        scen_param_folder = list(temp_filtered)
        scen_param_temp = scen_param_folder[0]
        scen_param_subvalues = scen_param_temp.GetContents('*.')
        # studyCases = folders[4].GetContents('*.IntCase')
        temp_filtered_L = filter(lambda scen_param_subvalues: scen_param_subvalues.loc_name == 'L',
                                 scen_param_subvalues)
        temp_filtered_R = filter(lambda scen_param_subvalues: scen_param_subvalues.loc_name == 'R',
                                 scen_param_subvalues)
        L = list(temp_filtered_L)
        R = list(temp_filtered_R)
        if scenario_name.endswith('_V'):
            L[0].vector=L_V
            R[0].vector=R_V
        if scenario_name.endswith('_I'):
            L[0].vector=L_I
            R[0].vector=R_I

        if scenario_name.endswith('_A'):
            L[0].vector=L_I
            R[0].vector=R_I

        k=1



    k=1



    return k



if __name__ == '__main__':


    #FILE_NAME_INPUT_1 = 'scenario_names_dig.xlsx'
    FILE_NAME_INPUT_1 = 'scenario_names_dig_BLE.xlsx'
    Project_name='Test1'
    #Project_name = '20210225-OCW01-New_VImp_Main_BLE_test1'
    #Project_name = 'New_OC_BLE_Filter1_test1'
    #Project_name ='New_OC_BLE_Filter2_withHigh_loci'
    # Project_name ='New_OC_BLE_Filter4'
    # Project_name = 'New_OC_BLE_IEC_1'
    # Project_name='OC_BLE_Injection_New_1'
    # Project_name = 'OC_BLE_Injection_New_4'
    # Project_name = 'OC_BLE_Injection_New_5'
    Project_name = 'OC_BLE_Injection_New_Main_New_4'

    # Amp_F_loader=True # if True it means that it's going to load the impedance values
    # # of AmPF scenarios which might be calculated using a different h values in point finder code
    # # if false it's going to load normal V and I scenario values.


    FILE_PTH_INPUT_1 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_1, folder_name='dig_results')
    sc_dig_df = common.import_excel(FILE_PTH_INPUT_1)
    scenario_dig_list = list(sc_dig_df['scenario'])
    output_excel_dig_list = ['{}_{}'.format(x, 'points.xlsx') for x in scenario_dig_list]
    scenario_dig_list_V = ['{}_{}'.format(x, 'V') for x in scenario_dig_list]
    scenario_dig_list_I = ['{}_{}'.format(x, 'I') for x in scenario_dig_list]
    scenario_dig_list_A = ['{}_{}'.format(x, 'A') for x in scenario_dig_list]
    dig_scenarios=scenario_dig_list_V + scenario_dig_list_I + scenario_dig_list_A

    pf_version = '2020'

    DIG_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5'
    DIG_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP5'
    DIG_PATH_2019 = r'C:\Program Files\DIgSILENT\PowerFactory 2019'
    DIG_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020'
    DIG_PYTHON_PATH_2016 = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5\Python\3.5'
    DIG_PYTHON_PATH_2018 = r'C:\Program Files\DIgSILENT\PowerFactory 2018 SP7\Python\3.5'
    DIG_PYTHON_PATH_2019 = r'C:\Program Files\DIgSILENT\PowerFactory 2019\Python\3.5'
    DIG_PYTHON_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020\Python\3.8'
    DIG_PYTHON_PATH_2020 = r'C:\Program Files\DIgSILENT\PowerFactory 2020\Python\3.8'

    if pf_version == '2016':
        DIG_PATH = DIG_PATH_2016
        DIG_PYTHON_PATH = DIG_PYTHON_PATH_2016
    elif pf_version == '2018':
        DIG_PATH = DIG_PATH_2018
        DIG_PYTHON_PATH = DIG_PYTHON_PATH_2018
    elif pf_version == '2019':
        DIG_PATH = DIG_PATH_2019
        DIG_PYTHON_PATH = DIG_PYTHON_PATH_2019
    elif pf_version == '2020':
        DIG_PATH = DIG_PATH_2020
        DIG_PYTHON_PATH = DIG_PYTHON_PATH_2020
    else:
        print('ERROR python version not found')
        raise SyntaxError('ERROR')

    sys.path.append(DIG_PATH)
    # #sys.path.append(DIG_PATH_2018)
    sys.path.append(DIG_PYTHON_PATH)
    # #sys.path.append(DIG_PYTHON_PATH_2018)

    os.environ['PATH'] = os.environ['PATH'] + ';' + DIG_PATH
    # noqa
    import powerfactory

    #app = powerfactory.GetApplication()

    df_out=loader(scenario_list=dig_scenarios,project=Project_name)

    # size = len(dig_scenarios[1])
    # # Slice string to remove last 3 characters from string
    # new_str = dig_scenarios[1][:size - 2]


