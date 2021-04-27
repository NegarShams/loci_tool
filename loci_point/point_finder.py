"""
#######################################################################################################################
###											scale											###
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


# class loci:
#     """
#     		General constants
#     	"""
#
#     def __init__(self, df):
#         """
#             Just to avoid error message
#         """
#         self.df = df.copy()
#         self.range = (1, 2)
#         pass


def point_adder(loci, p_num):
    """
        Function returns dataframe with mapped values.
    """
    loci.columns = ['X', 'R']
    # loci['Z']=np.nan
    # loci['V factor']=np.nan
    # loci['I factor']=np.nan
    row_num = len(loci) + (p_num * (len(loci) - 1))

    loci = loci.reset_index(drop=True)
    index = range(0, row_num)
    columns = ['X', 'R']
    loci_new = pd.DataFrame(index=index, columns=columns)
    for i in range(0, len(loci)):
        row = (i * p_num) + i
        loci_new.loc[row, columns] = loci.loc[i, columns]
        if row < row_num - 1:
            for n in range(0, p_num):
                row_in = row + n + 1
                p_num_1 = float(p_num)
                step = ((loci.loc[i + 1, 'X'] - loci.loc[i, 'X']) / (p_num_1))
                # step=assert truediv(10, 8)
                x_cur = loci.loc[i, 'X'] + (n + 1) * ((loci.loc[i + 1, 'X'] - loci.loc[i, 'X']) / (p_num_1 + 1))
                r_cur = loci.loc[i, 'R'] + (n + 1) * ((loci.loc[i + 1, 'R'] - loci.loc[i, 'R']) / (p_num_1 + 1))

                loci_new.loc[row_in, 'X'] = x_cur
                loci_new.loc[row_in, 'R'] = r_cur


    # loci_new=loci

    return loci_new


def worst_point(loci, R, X):
    """
        Function returns dataframe with mapped values.
    """

    loci.columns = ['X', 'R']
    loci['Z'] = np.nan
    loci['V factor'] = np.nan
    loci['I factor'] = np.nan

    k=1

    loci = loci.reset_index(drop=True)
    R_N = -1 * R
    X_N = -1 * X

    for i in range(0, len(loci)):
        loci.loc[i, 'Z'] = math.sqrt((loci.loc[i, 'R'] ** 2) + (loci.loc[i, 'X'] ** 2))
        loci.loc[i, 'I factor'] = math.sqrt(((loci.loc[i, 'R'] - R_N) ** 2) + ((loci.loc[i, 'X'] - X_N) ** 2))
        loci.loc[i, 'V factor'] = loci.loc[i, 'Z'] / loci.loc[i, 'I factor']
        max_row = loci['V factor'].argmax()
        min_row = loci['I factor'].argmin()

    R_worst_V = loci.loc[max_row, 'R']
    X_worst_V = loci.loc[max_row, 'X']
    R_worst_I = loci.loc[min_row, 'R']
    X_worst_I = loci.loc[min_row, 'X']


    return R_worst_V, X_worst_V, R_worst_I, X_worst_I


def finder(loci_excel, scan_excel, h_excel, added_point, output_excel,output_fig):
    """
            Function returns dataframe with mapped values.
        """
    FILE_PTH_INPUT_1 = common.get_local_file_path_withfolder(file_name=loci_excel, folder_name='Excel')
    FILE_PTH_INPUT_2 = common.get_local_file_path_withfolder(file_name=scan_excel, folder_name='dig_results')
    FILE_PTH_INPUT_3 = common.get_local_file_path_withfolder(file_name=h_excel, folder_name='Excel')
    FILE_PTH_OUTPUT_1 = common.get_local_file_path_withfolder(file_name=output_excel, folder_name='Output_excel')
    #FILE_PTH_OUTPUT_2 = common.get_local_file_path_withfolder(file_name=output_fig, folder_name='plots')

    loci_df = common.import_excel(FILE_PTH_INPUT_1)
    scan_df = common.import_excel(FILE_PTH_INPUT_2)
    h_df = common.import_excel(FILE_PTH_INPUT_3)

    h_df_len = len(h_df)
    loci_num = (len(loci_df.columns)) / 2
    df_list = list()
    range_list = list()

    scan_df['Loci Range'] = np.nan
    scan_df['worst_V_R'] = np.nan
    scan_df['worst_V_L'] = np.nan
    scan_df['worst_V_X'] = np.nan
    scan_df['worst_I_R'] = np.nan
    scan_df['worst_I_L'] = np.nan
    scan_df['worst_I_X'] = np.nan

    for i in range(0, loci_num * 2, 2):
        df = loci_df.iloc[1:, [i, i + 1]]
        df_list.append(df)
        range_list.append(df.columns)

    for i in range(1, len(scan_df) + 1):
        h = scan_df.loc[i, 'H']
        for r in range(0, loci_num):
            cur_range = range_list[r]
            if h >= cur_range[0]:
                if h <= cur_range[1]:
                    scan_df.loc[i, 'Loci Range'] = r

    for i in range(1, len(scan_df) + 1):
        loci_cur = df_list[int(scan_df.loc[i, 'Loci Range'])]
        loci_cur_new = point_adder(loci=loci_cur, p_num=added_point)
        r_w = scan_df.loc[i, 'R']
        x_w = scan_df.loc[i, 'X']
        R_worst_v, X_worst_v, R_worst_i, X_worst_i = worst_point(loci=loci_cur_new, R=r_w, X=x_w)
        if constants.plotting.save_plot:
            fig_title='h='+str(scan_df.loc[i, 'H'])
            plot_name=output_fig+str(scan_df.loc[i, 'H'])+constants.plotting.figure_type
            FILE_PTH_OUTPUT_2=common.get_local_file_path_withfolder(file_name=plot_name, folder_name=constants.plotting.folder)

            plotter.loci_scatter_plotter(loci_df=loci_cur_new, R=r_w, X=x_w, R_V=R_worst_v, X_V=X_worst_v, R_I=R_worst_i, X_I=X_worst_i,
                                         save_o=constants.plotting.save_plot,
                                         p_show=constants.plotting.show_plot,
                                         fig_path=FILE_PTH_OUTPUT_2,
                                         fig_name=fig_title)
            k=1

        scan_df.loc[i, 'worst_V_R'] = R_worst_v
        scan_df.loc[i, 'worst_V_X'] = X_worst_v
        scan_df.loc[i, 'worst_I_R'] = R_worst_i
        scan_df.loc[i, 'worst_I_X'] = X_worst_i
        scan_df.loc[i, 'worst_V_L'] = (X_worst_v / (2 * math.pi * 60 * scan_df.loc[i, 'H']))*1000
        scan_df.loc[i, 'worst_I_L'] = (X_worst_i / (2 * math.pi * 60 * scan_df.loc[i, 'H']))*1000

    if h_df_len > 0:
        idx = (scan_df['H'].isin(h_df['H'].tolist()))  # filters the buses that have GSPs inside the given GSP list
        scan_df = scan_df.loc[idx == True, :]

    scan_df.to_excel(FILE_PTH_OUTPUT_1)

    return scan_df


def finder_csv(loci_excel, csv_name, h_excel, added_point, output_excel,output_fig):
    """
            Function returns dataframe with mapped values.
        """
    FILE_PTH_INPUT_1 = common.get_local_file_path_withfolder(file_name=loci_excel, folder_name='Excel')
    #FILE_PTH_INPUT_2 = common.get_local_file_path_withfolder(file_name=scan_excel, folder_name='dig_results')
    FILE_PTH_INPUT_3 = common.get_local_file_path_withfolder(file_name=h_excel, folder_name='Excel')
    FILE_PTH_OUTPUT_1 = common.get_local_file_path_withfolder(file_name=output_excel, folder_name='Output_excel')
    #FILE_PTH_OUTPUT_2 = common.get_local_file_path_withfolder(file_name=output_fig, folder_name='plots')

    loci_df = common.import_excel(FILE_PTH_INPUT_1)

    scan_df=CSV_reader_df_maker(csv_name=csv_name, csv_folder_name=constants.folder_file_names.dig_result_folder,
                                object_name=constants.Digsilent.object_name, param_name_list=constants.Digsilent.param_name_list,
                                header_list=constants.Digsilent.header_list,number_or_name=constants.Digsilent.number_or_name_toggle,
                                param_number_list=constants.Digsilent.param_number_list)

    # csv_name, csv_folder_name, object_name, param_name_list, header_list

    #scan_df = common.import_excel(FILE_PTH_INPUT_2)
    h_df = common.import_excel(FILE_PTH_INPUT_3)



    h_df_len = len(h_df)
    loci_num = int((len(loci_df.columns)) / 2)
    df_list = list()
    range_list = list()

    # if h_df_len > 0:
    #     idx = (scan_df['H'].isin(h_df['H'].tolist()))  # filters the buses that have GSPs inside the given GSP list
    #     scan_df = scan_df.loc[idx == True, :]


    scan_df['Loci Range'] = np.nan
    scan_df['worst_V_R'] = np.nan
    scan_df['worst_V_L'] = np.nan
    scan_df['worst_V_X'] = np.nan
    scan_df['worst_I_R'] = np.nan
    scan_df['worst_I_L'] = np.nan
    scan_df['worst_I_X'] = np.nan

    for i in range(0, int(loci_num) * 2, 2):
        df = loci_df.iloc[1:, [i, i + 1]]
        df_list.append(df)
        range_list.append(df.columns)

    for i in range(1, len(scan_df) + 1):
        h = float(scan_df.loc[i, 'H'])
        for r in range(0, loci_num):
            cur_range = range_list[r]
            if h >= cur_range[0]:
                if h <= cur_range[1]:
                    scan_df.loc[i, 'Loci Range'] = r

    for i in range(1, len(scan_df) + 1):
        loci_cur = df_list[int(scan_df.loc[i, 'Loci Range'])]
        loci_cur_new = point_adder(loci=loci_cur, p_num=added_point)
        r_w = scan_df.loc[i, 'R']
        x_w = scan_df.loc[i, 'X']
        R_worst_v, X_worst_v, R_worst_i, X_worst_i = worst_point(loci=loci_cur_new, R=r_w, X=x_w)
        if constants.plotting.save_plot:
            fig_title='h='+str(scan_df.loc[i, 'H'])
            plot_name=output_fig+str(scan_df.loc[i, 'H'])+constants.plotting.figure_type
            FILE_PTH_OUTPUT_2=common.get_local_file_path_withfolder(file_name=plot_name, folder_name=constants.plotting.folder)

            plotter.loci_scatter_plotter(loci_df=loci_cur_new, R=r_w, X=x_w, R_V=R_worst_v, X_V=X_worst_v, R_I=R_worst_i, X_I=X_worst_i,
                                         save_o=constants.plotting.save_plot,
                                         p_show=constants.plotting.show_plot,
                                         fig_path=FILE_PTH_OUTPUT_2,
                                         fig_name=fig_title)
            k=1

        scan_df.loc[i, 'worst_V_R'] = R_worst_v
        scan_df.loc[i, 'worst_V_X'] = X_worst_v
        scan_df.loc[i, 'worst_I_R'] = R_worst_i
        scan_df.loc[i, 'worst_I_X'] = X_worst_i
        scan_df.loc[i, 'worst_V_L'] = (X_worst_v / (2 * math.pi * 60 * scan_df.loc[i, 'H']))*1000
        scan_df.loc[i, 'worst_I_L'] = (X_worst_i / (2 * math.pi * 60 * scan_df.loc[i, 'H']))*1000

    if h_df_len > 0:
        idx = (scan_df['H'].isin(h_df['H'].tolist()))  # filters the buses that have GSPs inside the given GSP list
        scan_df = scan_df.loc[idx == True, :]

    scan_df.to_excel(FILE_PTH_OUTPUT_1)

    return scan_df



def CSV_reader_df_maker(csv_name,csv_folder_name,object_name,param_name_list,header_list,number_or_name,param_number_list):
    """
        Function returns dataframe with mapped values.
    """
    FILE_PTH_INPUT = common.get_local_file_path_withfolder(file_name=csv_name, folder_name=csv_folder_name)
    df=pd.read_csv(FILE_PTH_INPUT)

    column_list=list()
    column_list.append(df.columns[0])
    column_list_o = list(filter(lambda local_x: local_x.startswith(object_name), df.columns))
    #column_list_o = list(filter(lambda local_x: local_x.startswith('POI'), df.columns))
    column_list=column_list+column_list_o

    df_ob=df.loc[:,column_list]
    df_ob.columns=df_ob.loc[0,:]
    df_new=df_ob.loc[1:,:]

    column_list_1=list()
    column_list_1.append(df_new.columns[0])
    if number_or_name:
        column_list_1=column_list_1+df_new.columns[param_number_list].tolist()
        # column_list_1.extend(df_new.columns[param_number_list])
        k=1

    else:
        column_list_1=column_list_1+param_name_list

    df_new=df_new.loc[:,column_list_1]
    df_new.columns=header_list

    df_new['R']=df_new['R'].astype(float)
    df_new['X'] = df_new['X'].astype(float)
    df_new['H'] = df_new['H'].astype(float)

    k=1

    return df_new



if __name__ == '__main__':

   #FILE_NAME_INPUT_1 = 'loci.xlsx'
    FILE_NAME_INPUT_1 = 'loci_BLE.xlsx'
    #FILE_NAME_INPUT_2 = 'intact_scan.xlsx'
    FILE_NAME_INPUT_3 = 'h.xlsx'
    FILE_NAME_INPUT_3 = 'h_new.xlsx'

    FILE_NAME_INPUT_4 = 'scenario_names.xlsx'
    #FILE_NAME_INPUT_6 = 'scenario_names_dig.xlsx'
    FILE_NAME_INPUT_6 = 'scenario_names_dig_BLE.xlsx'
    FILE_NAME_INPUT_5='OC_Intact.csv'
    dig_results_folder='dig_results'
    object_name_dig = 'POI 230kV FirstEnergy'
    param_name_list_dig = ['R, Re(Z) in Ohm', 'X, Im(Z) in Ohm']
    header_list_new=['H','R','X']
    csv_or_excel=True# if True then uses csv input if False uses excel input

    # scan_df_new=CSV_reader_df_maker(csv_name=FILE_NAME_INPUT_5,csv_folder_name=dig_results_folder,
    #                                 object_name=object_name_dig,param_name_list=param_name_list_dig,header_list=header_list_new)
    df=1

    #FILE_NAME_OUTPUT_1 = 'intact.xlsx'
    # FILE_PTH_INPUT_1 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_1, folder_name='Excel')
    # FILE_PTH_INPUT_2 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_2, folder_name='dig_results')
    FILE_PTH_INPUT_3 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_3, folder_name='Excel')
    FILE_PTH_INPUT_4 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_4, folder_name='dig_results')
    FILE_PTH_INPUT_6 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_6, folder_name='dig_results')

    # FILE_PTH_OUTPUT_1 = common.get_local_file_path_withfolder(file_name=FILE_NAME_OUTPUT_1, folder_name='Output_excel')
    point_number = 20

    # loci_df = common.import_excel(FILE_PTH_INPUT_1)
    # scan_df = common.import_excel(FILE_PTH_INPUT_2)
    h_df = common.import_excel(FILE_PTH_INPUT_3)
    h_df_len = len(h_df)
    sc_df = common.import_excel(FILE_PTH_INPUT_4)
    sc_dig_df = common.import_excel(FILE_PTH_INPUT_6)

    scenario_list = list(sc_df['scenario'])
    scenario_dig_list = list(sc_dig_df['scenario'])

    if csv_or_excel:
        sc_num=len(sc_dig_df)
        output_fig_list = ['{}_{}'.format(x, constants.plotting.string_added) for x in scenario_dig_list]
    else:
        sc_num = len(sc_df)
        output_fig_list = ['{}_{}'.format(x, constants.plotting.string_added) for x in scenario_list]

    scenario_excel_list = ['{}_{}'.format(x, 'scan.xlsx') for x in scenario_list]
    scenario_csv_list = ['{}{}'.format(x, '.csv') for x in scenario_dig_list]
    output_excel_list = ['{}_{}'.format(x, 'points.xlsx') for x in scenario_list]
    output_excel_dig_list = ['{}_{}'.format(x, 'points.xlsx') for x in scenario_dig_list]
    #output_fig_list = ['{}_{}'.format(x, 'fig.png') for x in scenario_list]

    save_fig=True
    show_fig=True

    for i in range(0, sc_num):

        # scan_df_new = CSV_reader_df_maker(csv_name=FILE_NAME_INPUT_5, csv_folder_name=dig_results_folder,
        #                                   object_name=object_name_dig, param_name_list=param_name_list_dig,
        #
        #                                  header_list=header_list_new)
        if csv_or_excel==False:

            scan_df_cur = finder(loci_excel=FILE_NAME_INPUT_1, scan_excel=scenario_excel_list[i], h_excel=FILE_NAME_INPUT_3,
                                 added_point=point_number
                                 , output_excel=output_excel_list[i],output_fig=output_fig_list[i])
        if csv_or_excel == True:
            scan_df_cur=finder_csv(loci_excel=FILE_NAME_INPUT_1, csv_name=scenario_csv_list[i], h_excel=FILE_NAME_INPUT_3, added_point=point_number,
                                   output_excel=output_excel_dig_list[i], output_fig=output_fig_list[i])
            k=1




    k = 1

    # # scan_df_1 = finder(loci_df=loci_df, scan_df=scan_df, h_df=h_df,added_point=point_number,output_path=FILE_PTH_OUTPUT_1)
    # # k=1
    #
    # h_df_len = len(h_df)
    #
    # loci_num = (len(loci_df.columns)) / 2
    # df_list = list()
    # range_list = list()
    # range_object_list = list()
    # scan_df['Loci Range'] = np.nan
    # scan_df['worst_V_R'] = np.nan
    # scan_df['worst_V_L'] = np.nan
    # scan_df['worst_V_X'] = np.nan
    # scan_df['worst_I_R'] = np.nan
    # scan_df['worst_I_L'] = np.nan
    # scan_df['worst_I_X'] = np.nan
    #
    # for i in range(0, loci_num * 2, 2):
    #     df = loci_df.iloc[1:, [i, i + 1]]
    #     df_list.append(df)
    #     range_list.append(df.columns)
    #     columns = df.columns
    #     columns = columns.astype(int)
    #
    # for i in range(1, len(scan_df) + 1):
    #     h = scan_df.loc[i, 'H']
    #     for r in range(0, loci_num):
    #         cur_range = range_list[r]
    #         if h >= cur_range[0]:
    #             if h <= cur_range[1]:
    #                 scan_df.loc[i, 'Loci Range'] = r
    #
    # for i in range(1, len(scan_df) + 1):
    #     loci_cur = df_list[int(scan_df.loc[i, 'Loci Range'])]
    #     loci_cur_new = point_adder(loci=loci_cur, p_num=point_number)
    #     k = 1
    #     r_w = scan_df.loc[i, 'R']
    #     x_w = scan_df.loc[i, 'X']
    #     R_worst_v, X_worst_v, R_worst_i, X_worst_i = worst_point(loci=loci_cur_new, R=r_w, X=x_w)
    #     scan_df.loc[i, 'worst_V_R'] = R_worst_v
    #     scan_df.loc[i, 'worst_V_X'] = X_worst_v
    #     scan_df.loc[i, 'worst_I_R'] = R_worst_i
    #     scan_df.loc[i, 'worst_I_X'] = X_worst_i
    #     scan_df.loc[i, 'worst_V_L'] = (X_worst_v / (2 * math.pi * 60 * scan_df.loc[i, 'H'])) * 1000
    #     scan_df.loc[i, 'worst_I_L'] = (X_worst_i / (2 * math.pi * 60 * scan_df.loc[i, 'H'])) * 1000
    #
    #     k = 1
    #
    # # if h_df_len > 0:
    # #     idx = (scan_df['H'].isin(list(gsp)))  # filters the buses that have GSPs inside the given GSP list
    # #     df_loads = df_loads.loc[idx == True, :]
    #
    # if h_df_len > 0:
    #     idx = (scan_df['H'].isin(h_df['H'].tolist()))  # filters the buses that have GSPs inside the given GSP list
    #     scan_df = scan_df.loc[idx == True, :]
    #
    # scan_df.to_excel(FILE_PTH_OUTPUT_1)

    k = 1
