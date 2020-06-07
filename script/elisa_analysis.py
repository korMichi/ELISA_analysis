'''
title: elisa_analysis.py
author: Michael Korenkov
date: 2020/06/07
'''

import numpy as np
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import matplotlib.pyplot as plt


def load_data(file, n):
    """
    Loads OD values from excel file into numpy arrays
    """
    workbook = load_workbook(filename=file)
    data, data1, data2, data3, data4 = [], [], [], [], []
    for sheets in workbook.sheetnames:
        data.append(pd.read_excel(file, sheet_name=sheets, usecols="A:M", index_col=0, skiprows=10, nrows=8))
    if n == 1:
        data1 = data
        return data1
    elif n == 2:
        data1 = data[0::n]
        data2 = data[1::n]
        return data1, data2
    elif n == 3:
        data1 = data[0::n]
        data2 = data[1::n]
        data3 = data[2::n]
        return data1, data2, data3
    elif n == 4:
        data1 = data[0::n]
        data2 = data[1::n]
        data3 = data[2::n]
        data4 = data[3::n]
        return data1, data2, data3, data4
    print(str(file) + " has been loaded")


def get_standard(array, i):
    """
    Gets the values for the antibody used as a standard on the plate and saves it into a new array while deleting itsself from the old array
    """
    axis = int(input("Enter the axis of the standard (1=1 OR 2=H): "))
    standard_array = []
    value_array = []
    for x in range(i):
        for n in range(len(array[x])):
            if axis == 1:
                standard_array.append(array[x][n].iloc[0:7, 0])
                value_array.append(array[x][n].drop(1, axis=1))
            else:
                standard_array.append(array[x][n].iloc[[0]])
                value_array.append(array[x][n].drop(["H"]))
    print(value_array)


def analysis():
    """
    Chooses which timepoints to plot
    """


def create_plot(plot_data):
    """
    Plots ELISA values with x=concentration and y=OD(405nm)
    """
    plt.plot(np.logspace(0.001, 1, 7), plot_data[0].iloc[0:7, 0])
    plt.xscale("log")
    plt.xlabel("concentration")
    plt.ylabel("OD 405nm")
    plt.show()


excel_file = input("Please enter the name of the excel file: ")
plates = int(input("Please enter the number of plates: "))
data = list(load_data(excel_file, plates))
get_standard(data, plates)
