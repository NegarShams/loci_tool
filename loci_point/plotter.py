#matplotlib inline

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
import matplotlib.pyplot as plt
# time = [0, 1, 2, 3]
# position = [0, 100, 200, 300]
#
# #plt.plot(time, position)
# plt.scatter(time, position)
# plt.xlabel('Time (hr)')
# plt.ylabel('Position (km)')
# plt.show()

def loci_scatter_plotter(loci_df,R,X,R_V,X_V,R_I,X_I,save_o,p_show,fig_path,fig_name):
    """
        Function returns dataframe with mapped values.
    """
    r_ = loci_df.loc[:,'R']
    x_ = loci_df.loc[:,'X']
    axis_space=5;
    #max_x=
    max_X=max(X,-1*X,x_.max())
    min_X=min(X,-1*X,x_.min())
    max_R = max(R, -1 * R, r_.max())
    min_R = min(R, -1 * R, r_.min())
    #list_x=list(x_)+list(X)+list(X_V)+list(X_I)
    #List_r=r_+R

    # plt.plot(time, position)


    #plt.plot(r_, x_,label='grid loci')
    #plt.scatter(R, X,lable='windfarm R&X')
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
    #plt.axis('scaled')
    #plt.axis("equal")
    x_space = (int(max_R) - int(min_R ))/ 5
    y_space = (int(max_X) - int(min_X)) / 5

    plt.axis(xmin=int(min_R - x_space), xmax=int(max_R + x_space),
             ymin=int(min_X - y_space),
             ymax=int(max_X + y_space))


    # plt.axis(xmin=int(min_R-constants.plotting.axis_space), xmax=int(max_R+constants.plotting.axis_space), ymin=int(min_X-constants.plotting.axis_space),
    #          ymax=int(max_X+constants.plotting.axis_space))

    x_step=(int(max_R+constants.plotting.axis_space)-int(min_R-constants.plotting.axis_space))/7
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
    #plt.legend(['grid loci','windfarm R&X','windfarm flipped R&X','worst R&X point for V','worst R&X point for I'])
    plt.legend(constants.plotting.legend)
    plt.grid(True)
    if save_o==True:
        plt.savefig(fig_path)
    if p_show==True:
        plt.show()

    plt.close()

    return



if __name__ == '__main__':

    FILE_NAME_INPUT_1 = 'loci_test.xlsx'
    FILE_NAME_INPUT_2='test_plot.png'
    #FILE_NAME_INPUT_2 = 'intact_scan.xlsx'
    r, x, r_v, x_v, r_i, x_i=13.320539,-73.3139,25,48,12.23809524,41.19047619
    plot_name='h=2'
    FILE_PTH_INPUT_1 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_1, folder_name='Excel')
    FILE_PTH_INPUT_2 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_2, folder_name='plots')
    loci=common.import_excel(FILE_PTH_INPUT_1)
    save=True
    figure_show=True

    loci_scatter_plotter(loci_df=loci,R=r,X=x,R_V=r_v,X_V=x_v,R_I=r_i,X_I=x_i,save_o=save,p_show=figure_show,fig_path=FILE_PTH_INPUT_2,
    fig_name='h=2')




    # scan_df_new=CSV_reader_df_maker(csv_name=FILE_NAME_INPUT_5,csv_folder_name=dig_results_folder,
    #                                 object_name=object_name_dig,param_name_list=param_name_list_dig,header_list=header_list_new)
