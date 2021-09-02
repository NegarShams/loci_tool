# matplotlib inline

# Generic Imports
# from __future__ import division
from operator import truediv, floordiv
import pandas as pd
import numpy as np
# Unique imports
import common_functions as common
import point_finder as p_finder
import constants as constants
import collections
# import load_est
# import load_est.psse as psse
import math
import time
import os
import dill
import csv
import matplotlib.pyplot as plt


# time = [0, 1, 2, 3]
# position = [0, 100, 200, 300]
#
# #plt.plot(time, position)
# plt.scatter(time, position)
# plt.xlabel('Time (hr)')
# plt.ylabel('Position (km)')
# plt.show()


def CSV_reader_df_maker_General(csv_name, csv_folder_name, object_name, param_name_list, header_list, number_or_name,
                                param_number_list, first_col_in):
    """
        Function returns dataframe with mapped values.
    """
    FILE_PTH_INPUT = common.get_local_file_path_withfolder(file_name=csv_name, folder_name=csv_folder_name)
    df = pd.read_csv(FILE_PTH_INPUT)

    column_list = list()
    if first_col_in:
        column_list.append(df.columns[0])

    column_list_o = list(filter(lambda local_x: local_x.startswith(object_name), df.columns))
    # column_list_o = list(filter(lambda local_x: local_x.startswith('POI'), df.columns))
    column_list = column_list + column_list_o

    df_ob = df.loc[:, column_list]
    df_ob.columns = df_ob.loc[0, :]
    df_new = df_ob.loc[1:, :]

    column_list_1 = list()
    if first_col_in:
        column_list_1.append(df_new.columns[0])

    if number_or_name:
        column_list_1 = column_list_1 + df_new.columns[param_number_list].tolist()
        # column_list_1.extend(df_new.columns[param_number_list])
        k = 1

    else:
        column_list_1 = column_list_1 + param_name_list

    df_new = df_new.loc[:, column_list_1]
    df_new.columns = header_list
    for i in range(0, len(header_list)):
        df_new[header_list[i]] = df_new[header_list[i]].astype(float)

    # df_new[header_list[0]] = df_new[header_list[0]].astype(float)
    # df_new[header_list[1]] = df_new[header_list[1]].astype(float)
    # df_new[header_list[2]] = df_new[header_list[2]].astype(float)
    # df_new[header_list[3]] = df_new[header_list[3]].astype(float)
    k = 1

    return df_new


def RMS_analyser(df, sw_time, POW):
    """
            Function returns dataframe with mapped values.
        """
    time_step = df.loc[2, 'Time'] - df.loc[1, 'Time']
    f = 50
    t_length = 1 / f
    n_points = t_length / time_step

    st_time = sw_time + (0.001 * POW) + 0.02
    time_df = (st_time / time_step)

    # A_RMS_max = df.loc[time_df:len(df), 'A_RMS'].max()
    # B_RMS_max = df.loc[time_df:len(df), 'B_RMS'].max()
    # C_RMS_max = df.loc[time_df:len(df), 'C_RMS'].max()

    A_RMS_max = df.loc[1:len(df), 'A_RMS'].max()
    B_RMS_max = df.loc[1:len(df), 'B_RMS'].max()
    C_RMS_max = df.loc[1:len(df), 'C_RMS'].max()

    A_I_max = df.loc[1:len(df), 'Phase A'].max()
    B_I_max = df.loc[1:len(df), 'Phase B'].max()
    C_I_max = df.loc[1:len(df), 'Phase C'].max()

    A_I_min = df.loc[1:len(df), 'Phase A'].min()
    B_I_min = df.loc[1:len(df), 'Phase B'].min()
    C_I_min = df.loc[1:len(df), 'Phase C'].min()

    # max_row = loci['V factor'].argmax()
    # min_row = loci['I factor'].argmin()

    k = 1
    return A_RMS_max, A_RMS_max, C_RMS_max, A_I_max, B_I_max, C_I_max, A_I_min, B_I_min, C_I_min


def RMS_Maker(df, TOV_L):
    """
            Function returns dataframe with mapped values.
        """
    time_step = df.loc[2, 'Time'] - df.loc[1, 'Time']
    f = 50
    t_length = 1 / f
    n_points = t_length / time_step

    df.loc[:, 'A'] = (df.loc[:, 'Phase A'] ** 2)
    df.loc[:, 'B'] = (df.loc[:, 'Phase B'] ** 2)
    df.loc[:, 'C'] = (df.loc[:, 'Phase C'] ** 2)
    df.loc[:, 'TOV_Limit_P'] = TOV_L
    df.loc[:, 'TOV_Limit_N'] = TOV_L * -1
    df.loc[:, 'TOV_Limit_P_RMS'] = TOV_L / math.sqrt(2)
    df.loc[:, 'TOV_Limit_N_RMS'] = TOV_L / math.sqrt(2) * -1

    # df['A_RMS'] = df['Time'].apply(lambda val: math.sqrt(sum(df.loc[i:i + n_points, 'A']) / n_points)  if val>5 else 4)

    # for i in range(1, len(df) + 1):
    #     df.loc[i, 'A'] = (df.loc[i, 'Phase A'] ** 2)
    #     df.loc[i, 'B'] = (df.loc[i, 'Phase B'] ** 2)
    #     df.loc[i, 'C'] = (df.loc[i, 'Phase C'] ** 2)

    for i in range(1, len(df) + 1):
        if i <= n_points:
            df.loc[i, 'A_RMS'] = 0
            df.loc[i, 'B_RMS'] = 0
            df.loc[i, 'C_RMS'] = 0
        else:
            df.loc[i, 'A_RMS'] = math.sqrt(sum(df.loc[i:i + n_points, 'A']) / n_points)
            df.loc[i, 'B_RMS'] = math.sqrt(sum(df.loc[i:i + n_points, 'B']) / n_points)
            df.loc[i, 'C_RMS'] = math.sqrt(sum(df.loc[i:i + n_points, 'C']) / n_points)

    return df


def TOV_Plotter(df_plt, save_o_TOV, p_show_TOV, fig_name, is_RMS, max_time):
    """
            Function returns dataframe with mapped values.
        """
    x_a = df_plt.loc[:, 'Time'] * 1000
    x_b = df_plt.loc[:, 'Time'] * 1000
    x_c = df_plt.loc[:, 'Time'] * 1000
    x_p = df_plt.loc[:, 'Time'] * 1000
    x_n = df_plt.loc[:, 'Time'] * 1000

    if is_RMS:
        y_a = df_plt.loc[:, 'A_RMS']
    else:
        y_a = df_plt.loc[:, 'Phase A']

    if is_RMS:
        y_b = df_plt.loc[:, 'B_RMS']
    else:
        y_b = df_plt.loc[:, 'Phase B']

    if is_RMS:
        y_c = df_plt.loc[:, 'C_RMS']
    else:
        y_c = df_plt.loc[:, 'Phase C']

    if is_RMS:
        y_p = df_plt.loc[:, 'TOV_Limit_P_RMS']
    else:
        y_p = df_plt.loc[:, 'TOV_Limit_P']

    if is_RMS:
        y_n = df_plt.loc[:, 'TOV_Limit_N_RMS']
    else:
        y_n = df_plt.loc[:, 'TOV_Limit_N']

    # # y_b = df_plt.loc[:, 'Phase B']
    #
    # y_b = df_plt.loc[:, 'B_RMS']
    #
    # x_c = df_plt.loc[:, 'Time']
    # # y_c = df_plt.loc[:, 'Phase C']
    #
    # y_c = df_plt.loc[:, 'C_RMS']

    # x_p = df_plt.loc[:, 'Time']
    # y_p = df_plt.loc[:, 'TOV_Limit_P']
    #
    # x_n = df_plt.loc[:, 'Time']
    # y_n = df_plt.loc[:, 'TOV_Limit_N']

    if is_RMS:
        #plt.axis(xmin=int(600), xmax=(1000), ymin=int(0), ymax=int(400))
        plt.axis(xmin=int(100), xmax=(500), ymin=int(0), ymax=int(400))
    else:
        #plt.axis(xmin=int(600), xmax=int(1000), ymin=int(-600), ymax=int(600))
        plt.axis(xmin=int(100), xmax=int(500), ymin=int(-600), ymax=int(600))

    # plt.axis(xmin=int(0), xmax=int(2))

    plt.xlabel('Time (ms)')
    plt.ylabel('Voltage (kV)')

    plt.plot(x_a, y_a, 'r')
    plt.plot(x_b, y_b, 'b')
    plt.plot(x_c, y_c, 'g')

    plt.plot(x_p, y_p, 'k--')
    plt.plot(x_n, y_n, 'k--')

    plt.legend()
    # plt.legend(['grid loci','windfarm R&X','windfarm flipped R&X','worst R&X point for V','worst R&X point for I'])
    plt.legend(constants.plotting.TOV_legend)

    plt.grid(True)

    plt.title(fig_name)

    if is_RMS:
        plot_name = fig_name + '_RMS' + constants.plotting.figure_type
    else:
        plot_name = fig_name + constants.plotting.figure_type

    # fig_title = 'h=' + str(scan_df.loc[i, 'H'])
    # plot_name = output_fig + str(scan_df.loc[i, 'H']) + constants.plotting.figure_type
    fig_path = common.get_local_file_path_withfolder(file_name=plot_name,
                                                     folder_name=constants.plotting.TOV_plots_folder)

    if save_o_TOV:
        plt.savefig(fig_path)
    if p_show_TOV:
        plt.show()

    if not is_RMS:
        plt.axis(xmin=int(0), xmax=int(max_time), ymin=int(-600), ymax=int(600))
        plot_name = fig_name + '_Total' + constants.plotting.figure_type
        fig_path = common.get_local_file_path_withfolder(file_name=plot_name,
                                                         folder_name=constants.plotting.TOV_plots_folder)
        if save_o_TOV:
            plt.savefig(fig_path)

    plt.close()

    # if save_o == True:
    #     plt.savefig(fig_path)
    # if p_show == True:
    #     plt.show()
    #
    # plt.close()

    # plt.show()


def TOV_SC_Runner(TOV_List, object_Name, output_name, is_pow, TOV_Limit, plot_TOV, save_o_TOV, p_show_TOV,
                  first_col_in,param_num_list,max_time):
    """
        Function returns dataframe with mapped values.
    """
    if first_col_in:
        TOV_header_List=constants.Digsilent.TOV_header_list
    else:
        TOV_header_List = constants.Digsilent.TOV_header_list_no_fc

    TOV_POW_List = np.arange(0, 360 + 18, 18)

    if is_pow:
        final_Df = pd.DataFrame(TOV_POW_List, columns=['POW'])
        csv_List = ['{}{}'.format(x, '.csv') for x in TOV_POW_List]
        results_folder = constants.folder_file_names.POW_folder

    else:
        final_Df = pd.DataFrame(TOV_List, columns=['Scenario'])
        csv_List = ['{}{}'.format(x, '.csv') for x in TOV_List]
        results_folder = constants.folder_file_names.TOV_folder

    output_TOV_list = ['{}{}'.format(x, '.xlsx') for x in output_name]

    final_Df['A_max_RMS'] = np.nan
    final_Df['B_max_RMS'] = np.nan
    final_Df['C_max_RMS'] = np.nan
    final_Df['A_max_I'] = np.nan
    final_Df['B_max_I'] = np.nan
    final_Df['C_max_I'] = np.nan
    final_Df['A_min_I'] = np.nan
    final_Df['B_min_I'] = np.nan
    final_Df['C_min_I'] = np.nan
    final_Df['TOV_limit_P'] = np.nan
    final_Df['TOV_limit_N'] = np.nan

    for i in range(0, len(csv_List)):
        scan_df = CSV_reader_df_maker_General(csv_name=csv_List[i],
                                              csv_folder_name=results_folder,
                                              object_name=object_Name,
                                              param_name_list=constants.Digsilent.param_name_list,
                                              header_list=TOV_header_List,
                                              number_or_name=constants.Digsilent.number_or_name_toggle,
                                              param_number_list=param_num_list, first_col_in=first_col_in)
        if is_pow:
            POW = i * 18
        else:
            POW = 0

        if not first_col_in:
            scan_df['Time']= list(range(0,len(scan_df)))
            scan_df['Time']=scan_df['Time']/10000

        df_RMS = RMS_Maker(scan_df, TOV_Limit)

        if plot_TOV:
            is_RMS = False
            TOV_Plotter(df_RMS, save_o_TOV, p_show_TOV, TOV_List[i], is_RMS,max_time)
            is_RMS = True
            TOV_Plotter(df_RMS, save_o_TOV, p_show_TOV, TOV_List[i], is_RMS,max_time)

        A_RMS_max, B_RMS_max, C_RMS_max, A_I_max, B_I_max, C_I_max, A_I_min, B_I_min, C_I_min = RMS_analyser(df_RMS,
                                                                                                             swi_time,
                                                                                                             POW)

        final_Df.loc[i, 'A_max_RMS'] = A_RMS_max
        final_Df.loc[i, 'B_max_RMS'] = B_RMS_max
        final_Df.loc[i, 'C_max_RMS'] = C_RMS_max
        final_Df.loc[i, 'A_max_I'] = A_I_max
        final_Df.loc[i, 'B_max_I'] = B_I_max
        final_Df.loc[i, 'C_max_I'] = C_I_max
        final_Df.loc[i, 'A_min_I'] = A_I_min
        final_Df.loc[i, 'B_min_I'] = B_I_min
        final_Df.loc[i, 'C_min_I'] = C_I_min

        final_Df.loc[i, 'TOV_limit_P'] = TOV_Limit
        final_Df.loc[i, 'TOV_limit_N'] = TOV_Limit * -1

    FILE_PTH_OUTPUT_1 = common.get_local_file_path_withfolder(file_name=output_TOV_list[0], folder_name='TOV_Output')

    final_Df.to_excel(FILE_PTH_OUTPUT_1)
    return final_Df


def loci_scatter_plotter(loci_df, R, X, R_V, X_V, R_I, X_I, save_o, p_show, fig_path, fig_name):
    """
        Function returns dataframe with mapped values.
    """
    r_ = loci_df.loc[:, 'R']
    x_ = loci_df.loc[:, 'X']
    axis_space = 5;
    # max_x=
    max_X = max(X, -1 * X, x_.max())
    min_X = min(X, -1 * X, x_.min())
    max_R = max(R, -1 * R, r_.max())
    min_R = min(R, -1 * R, r_.min())
    # list_x=list(x_)+list(X)+list(X_V)+list(X_I)
    # List_r=r_+R

    # plt.plot(time, position)

    # plt.plot(r_, x_,label='grid loci')
    # plt.scatter(R, X,lable='windfarm R&X')
    # plt.scatter(-1*R, -1*X,lable='windfarm flipped R&X')
    # plt.scatter(R_V, X_V,lable='worst R&X point for V')
    # plt.scatter(R_I, X_I,lable='worst R&X point for I')
    #
    plt.plot(r_, x_)
    plt.scatter(R, X)
    plt.scatter(-1 * R, -1 * X)
    plt.scatter(R_V, X_V)
    plt.scatter(R_I, X_I)
    plt.xlabel(constants.plotting.x_label)
    plt.ylabel(constants.plotting.y_lable)
    # plt.xlim(int(min_R-20), int(max_R+20))
    # plt.ylim(int(min_X-20), int(max_X+20))
    # plt.xlim(-10, 10)
    # plt.ylim(-40, 40)
    plt.title(fig_name)
    # plt.axis('scaled')
    # plt.axis("equal")
    x_space = (int(max_R) - int(min_R)) / 5
    y_space = (int(max_X) - int(min_X)) / 5

    plt.axis(xmin=int(min_R - x_space), xmax=int(max_R + x_space),
             ymin=int(min_X - y_space),
             ymax=int(max_X + y_space))

    # plt.axis(xmin=int(min_R-constants.plotting.axis_space), xmax=int(max_R+constants.plotting.axis_space), ymin=int(min_X-constants.plotting.axis_space),
    #          ymax=int(max_X+constants.plotting.axis_space))

    x_step = (int(max_R + constants.plotting.axis_space) - int(min_R - constants.plotting.axis_space)) / 7
    y_step = (int(max_X + constants.plotting.axis_space) - int(min_X - constants.plotting.axis_space)) / 7

    # plt.xticks(np.arange(int(min_R - constants.plotting.axis_space), int(max_R + constants.plotting.axis_space),
    #                      x_step))
    #
    # plt.yticks(np.arange(int(min_X - constants.plotting.axis_space), int(max_X + constants.plotting.axis_space),
    #                      y_step))

    plt.xticks(np.arange(int(min_R) - x_space, int(max_R) + x_space,
                         x_step))
    plt.yticks(np.arange(int(min_X) - y_space, int(max_X) + y_space,
                         y_step))

    # plt.xticks(np.arange(int(min_R-constants.plotting.axis_space), int(max_R+constants.plotting.axis_space), constants.plotting.axis_step))
    # plt.yticks(np.arange(int(min_X - constants.plotting.axis_space), int(max_X + constants.plotting.axis_space), constants.plotting.axis_step))
    plt.legend()
    # plt.legend(['grid loci','windfarm R&X','windfarm flipped R&X','worst R&X point for V','worst R&X point for I'])
    plt.legend(constants.plotting.legend)
    plt.grid(True)
    # if save_o==True:
    #     plt.savefig(fig_path)
    # if p_show==True:
    #     plt.show()
    #
    # plt.close()

    return


if __name__ == '__main__':

    #object_Name_TOV = 'WOO'
    first_col_in = True
    param_num_list = [1, 2, 3]

    object_Name_TOV = 'DSN'
    first_col_in = False
    param_num_list=[0,1,2]

    first_col_in = True
    param_num_list = [1, 2, 3]

    max_time=2000
    #max_time = 500

    TOV_sc_list = ['SV_WOO4_3ph_W_Old']  # \
    # this is the list of the names of the final output (general scenario name for POW case)
    TOV_List = ['W_SV_PRF_Old_3p', 'W_SV_PRF_Por_3p']
    TOV_List = ['W_SV_WOOT_Old_3p', 'W_SV_WOOT_Old_1p', 'W_SV_WOOT_Por_3p', 'W_SV_WOOT_Dun_3p', 'W_SV_PRF_Tur_3p',
                'W_SV_PRF_Tur_1p', \
                'W_SV_PRF_Old_3p', 'W_SV_PRF_Old_1p', 'W_SV_PRF_Por_3p', 'W_SV_PRF_Dun_3p', 'W_SV_PRF_Dun_1p']

    TOV_List = ['W_SV_PRF_Dun_3p', 'W_SV_PRF_Dun_1p']

    TOV_List = ['D_SV_WOOT_Coo_3p', 'D_SV_WOOT_Coo_1p', 'D_SV_WOOT_Woo_3p', 'D_SV_WOOT_Woo_1p', 'D_SV_PRF_Coo_3p',
                'D_SV_PRF_Coo_1p', \
                'D_SV_PRF_Woo_3p', 'D_SV_PRF_Woo_1p']

    # TOV_List = ['D_SV_WOOT_Coo_3p', 'D_SV_WOOT_Coo_1p', 'D_SV_WOOT_Woo_3p', 'D_SV_WOOT_Woo_1p', 'D_SV_PRF_Coo_3p',
    #             'D_SV_PRF_Coo_1p']

    TOV_List = ['D_SV_PRF_Woo_3p', 'D_SV_PRF_Woo_1p']

    TOV_List = ['D_SV_PRF_Woo_3p', 'D_SV_PRF_Woo_1p']

    TOV_List = ['W_SV_PRF_Dun_Lin_0',	'W_SV_PRF_Dun_Lin_18',	'W_SV_PRF_Dun_Lin_36',	'W_SV_PRF_Dun_Lin_54',\
                'W_SV_PRF_Dun_Lin_72',	'W_SV_PRF_Dun_Lin_90',	'W_SV_PRF_Dun_Lin_108',	'W_SV_PRF_Dun_Lin_126',\
                'W_SV_PRF_Dun_Lin_144',	'W_SV_PRF_Dun_Lin_162',	'W_SV_PRF_Dun_Lin_180',	'W_SV_PRF_Dun_Lin_198',\
                'W_SV_PRF_Dun_Lin_216',	'W_SV_PRF_Dun_Lin_234',	'W_SV_PRF_Dun_Lin_252',	'W_SV_PRF_Dun_Lin_270',\
                'W_SV_PRF_Dun_Lin_288',	'W_SV_PRF_Dun_Lin_306',	'W_SV_PRF_Dun_Lin_324',	'W_SV_PRF_Dun_Lin_342',	'W_SV_PRF_Dun_Lin_360']

    TOV_List = ['W_SV_PRF_Dun_Lin_0', 'W_SV_PRF_Dun_Lin_18', 'W_SV_PRF_Dun_Lin_36', 'W_SV_PRF_Dun_Lin_54', \
                'W_SV_PRF_Dun_Lin_72', 'W_SV_PRF_Dun_Lin_90', 'W_SV_PRF_Dun_Lin_108', 'W_SV_PRF_Dun_Lin_126', \
                'W_SV_PRF_Dun_Lin_144', 'W_SV_PRF_Dun_Lin_162', 'W_SV_PRF_Dun_Lin_180', 'W_SV_PRF_Dun_Lin_198', \
                'W_SV_PRF_Dun_Lin_216', 'W_SV_PRF_Dun_Lin_234', 'W_SV_PRF_Dun_Lin_252', 'W_SV_PRF_Dun_Lin_270', \
                'W_SV_PRF_Dun_Lin_288', 'W_SV_PRF_Dun_Lin_306', 'W_SV_PRF_Dun_Lin_324', 'W_SV_PRF_Dun_Lin_342',
                'W_SV_PRF_Dun_Lin_360']

    TOV_List = ['W_SV_PRF_Tx4_0', 'W_SV_PRF_Tx4_18', 'W_SV_PRF_Tx4_36', 'W_SV_PRF_Tx4_54', \
                'W_SV_PRF_Tx4_72', 'W_SV_PRF_Tx4_90', 'W_SV_PRF_Tx4_108', 'W_SV_PRF_Tx4_126', \
                'W_SV_PRF_Tx4_144', 'W_SV_PRF_Tx4_162', 'W_SV_PRF_Tx4_180', 'W_SV_PRF_Tx4_198', \
                'W_SV_PRF_Tx4_216', 'W_SV_PRF_Tx4_234', 'W_SV_PRF_Tx4_252', 'W_SV_PRF_Tx4_270', \
                'W_SV_PRF_Tx4_288', 'W_SV_PRF_Tx4_306', 'W_SV_PRF_Tx4_324', 'W_SV_PRF_Tx4_342',
                'W_SV_PRF_Tx4_360']

    TOV_List = ['D_SV_WOOT_Coo_Lin','D_SV_PRF_Coo_Lin']
    TOV_List = ['W_SV_WOOT_Old_Lin', 'W_SV_PRF_Old_Lin']
    TOV_List = ['D_SV_PRF_Woo_Lin']
    TOV_List = ['W_SV_PRF_Dun_Lin']
    TOV_List = ['W_SV_PRF_Dun_Lin_5s']
    TOV_List = ['D_SV_PRF_Tx2']
    #TOV_List = ['W_SV_PRF_Tx1','W_SV_PRF_Tx2','W_SV_PRF_Tx4']

    is_pow = False
    plot_TOV = True
    save_o_TOV = True
    p_show_TOV = False

    output_name = ['all_fault_cases']
    output_name = ['missing_fault_cases']
    output_name = ['missing_cases_D']
    # output_name = ['all_cases_D']
    output_name = ['W_SV_PRF_Dun_Lin_All']
    output_name = ['W_SV_PRF_Tx4_All']
    output_name = ['D_Lines_All']
    output_name = ['W_final_lines']
    output_name = ['D_SV_PRF_Woo_Lin_Results']
    output_name = ['W_SV_PRF_Dun_Lin_Results']
    output_name = ['W_SV_PRF_Dun_Lin_5s_Results']
    output_name = ['W_TX_Results']
    output_name = ['D_TX_Results']
    output_name = ['D_TX_Results_1']

    v_level = 380
    TOV_ratio = 1.6
    swi_time = 0.69
    TOV_Limit = (v_level * math.sqrt(2) / math.sqrt(3)) * TOV_ratio
    # TOV_Limit = (v_level * math.sqrt(1) / math.sqrt(3)) * TOV_ratio

    Final_DF = TOV_SC_Runner(TOV_List, object_Name_TOV, output_name, is_pow, TOV_Limit, plot_TOV, save_o_TOV,
                             p_show_TOV, first_col_in,param_num_list,max_time)

    k = 1

    TOV_POW_List = np.arange(0, 360 + 18, 18)

    csv_list = ['{}{}'.format(x, '.csv') for x in TOV_POW_List]
    final_df = pd.DataFrame(TOV_POW_List, columns=['POW'])
    output_TOV_list = ['{}{}'.format(x, '.xlsx') for x in TOV_sc_list]

    final_df['A_max_RMS'] = np.nan
    final_df['B_max_RMS'] = np.nan
    final_df['C_max_RMS'] = np.nan
    final_df['A_max_I'] = np.nan
    final_df['B_max_I'] = np.nan
    final_df['C_max_I'] = np.nan
    final_df['A_min_I'] = np.nan
    final_df['B_min_I'] = np.nan
    final_df['C_min_I'] = np.nan
    final_df['TOV_limit_P'] = np.nan
    final_df['TOV_limit_N'] = np.nan

    for i in range(0, len(csv_list)):
        scan_df = CSV_reader_df_maker_General(csv_name=csv_list[i],
                                              csv_folder_name=constants.folder_file_names.POW_folder,
                                              object_name=object_Name_TOV,
                                              param_name_list=constants.Digsilent.param_name_list,
                                              header_list=constants.Digsilent.TOV_header_list,
                                              number_or_name=constants.Digsilent.number_or_name_toggle,
                                              param_number_list=[1, 2, 3])
        POW = i * 18

        df_RMS = RMS_Maker(scan_df, TOV_Limit)
        A_RMS_max, B_RMS_max, C_RMS_max, A_I_max, B_I_max, C_I_max, A_I_min, B_I_min, C_I_min = RMS_analyser(df_RMS,
                                                                                                             swi_time,
                                                                                                             POW)

        final_df.loc[i, 'A_max_RMS'] = A_RMS_max
        final_df.loc[i, 'B_max_RMS'] = B_RMS_max
        final_df.loc[i, 'C_max_RMS'] = C_RMS_max
        final_df.loc[i, 'A_max_I'] = A_I_max
        final_df.loc[i, 'B_max_I'] = B_I_max
        final_df.loc[i, 'C_max_I'] = C_I_max
        final_df.loc[i, 'A_min_I'] = A_I_min
        final_df.loc[i, 'B_min_I'] = B_I_min
        final_df.loc[i, 'C_min_I'] = C_I_min

        final_df.loc[i, 'TOV_limit_P'] = TOV_Limit
        final_df.loc[i, 'TOV_limit_N'] = TOV_Limit * -1

    FILE_PTH_OUTPUT_1 = common.get_local_file_path_withfolder(file_name=output_TOV_list[0], folder_name='TOV_Output')

    final_df.to_excel(FILE_PTH_OUTPUT_1)
    k = 1

    scan_df = CSV_reader_df_maker_General(csv_name=csv_list[0], csv_folder_name=constants.folder_file_names.TOV_folder,
                                          object_name=object_Name_TOV,
                                          param_name_list=constants.Digsilent.param_name_list,
                                          header_list=constants.Digsilent.TOV_header_list,
                                          number_or_name=constants.Digsilent.number_or_name_toggle,
                                          param_number_list=[1, 2, 3])
    POW = 0

    df_RMS = RMS_Maker(scan_df, TOV_Limit)
    A_RMS_max, A_RMS_max, C_RMS_max, A_I_max, B_I_max, C_I_max, A_I_min, B_I_min, C_I_min = RMS_analyser(df_RMS,
                                                                                                         swi_time, POW)

    TOV_Plotter(df_RMS)
    x = df_RMS.loc[:, 'Time']
    y = df_RMS.loc[:, 'A']

    plt.plot(x, y)
    plt.show()

    time_step = scan_df.loc[2, 'Time'] - scan_df.loc[1, 'Time']
    f = 50
    t_length = 1 / f
    n_points = t_length / time_step

    for i in range(1, len(scan_df) + 1):
        scan_df.loc[i, 'A'] = (scan_df.loc[i, 'Phase A'] ** 2)
        scan_df.loc[i, 'B'] = (scan_df.loc[i, 'Phase B'] ** 2)
        scan_df.loc[i, 'C'] = (scan_df.loc[i, 'Phase C'] ** 2)

    for i in range(1, len(scan_df) + 1):
        if i <= n_points:
            scan_df.loc[i, 'A_RMS'] = 0
            scan_df.loc[i, 'B_RMS'] = 0
            scan_df.loc[i, 'C_RMS'] = 0
        else:
            scan_df.loc[i, 'A_RMS'] = math.sqrt(sum(scan_df.loc[i:i + n_points, 'A']) / n_points)
            scan_df.loc[i, 'B_RMS'] = math.sqrt(sum(scan_df.loc[i:i + n_points, 'B']) / n_points)
            scan_df.loc[i, 'C_RMS'] = math.sqrt(sum(scan_df.loc[i:i + n_points, 'C']) / n_points)

    # sum=sum(scan_df.loc[1:2,'A'])

    # for i in range(1, len(scan_df) + 1):

    # scan_df.loc[i, 'A'] = math.sqrt((scan_df.loc[i, 'Phase A'] ** 2) + (loci.loc[i, 'X'] ** 2))
    # loci.loc[i, 'I factor'] = math.sqrt(((loci.loc[i, 'R'] - R_N) ** 2) + ((loci.loc[i, 'X'] - X_N) ** 2))

    FILE_NAME_INPUT_1 = 'loci_test.xlsx'
    FILE_NAME_INPUT_2 = 'test_plot.png'
    # FILE_NAME_INPUT_2 = 'intact_scan.xlsx'
    r, x, r_v, x_v, r_i, x_i = 13.320539, -73.3139, 25, 48, 12.23809524, 41.19047619
    plot_name = 'h=2'
    FILE_PTH_INPUT_1 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_1, folder_name='Excel')
    FILE_PTH_INPUT_2 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_2, folder_name='plots')
    loci = common.import_excel(FILE_PTH_INPUT_1)
    save = True
    figure_show = True

    loci_scatter_plotter(loci_df=loci, R=r, X=x, R_V=r_v, X_V=x_v, R_I=r_i, X_I=x_i, save_o=save, p_show=figure_show,
                         fig_path=FILE_PTH_INPUT_2,
                         fig_name='h=2')

    # scan_df_new=CSV_reader_df_maker(csv_name=FILE_NAME_INPUT_5,csv_folder_name=dig_results_folder,
    #                                 object_name=object_name_dig,param_name_list=param_name_list_dig,header_list=header_list_new)
