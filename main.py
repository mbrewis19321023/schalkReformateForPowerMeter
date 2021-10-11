import os
import sys
from numpy import mod
from openpyxl.styles import NamedStyle, Font, Border, Side
import tkinter as tk
from tkinter import filedialog, Text
import tkinter.font as font
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter.filedialog import asksaveasfile
from openpyxl.styles import Alignment
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import NamedStyle, Font, Border, Side, numbers
import pandas as pd
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
import re

f = open("myfile.csv", "w")

digitReplace = re.compile(r"(\d+)(,)(\d+)")
dateModify = re.compile(r"(\d{2})\-(\d{2})\-\d{2}(\d{2})")
regGroup = re.compile(
    r"\"\",\"(.*?)\",\"(.*?)\",\"(.*?)\",\"(.*?)\",\"(.*?)\",\"(.*?)\",\"(.*?)\",\"(.*?)\",")

monthList = ['Jan', 'Feb', 'Mar', 'Apr', 'May',
             'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
csvFile = open("q2.csv")
lines = csvFile.readlines()





j = 0
f.write(r'"Load Profile Data","","","","","","","",""' + '\n')
f.write(r'"","Date","Start","End","Import W","blank W","Q4 Cap Exp","Export W","las","Total VA",' + '\n')
f.write(r'"","","","","(kW)","(kW)","(kVAR)","(kW)","","(kVA)" ' + '\n')
for x in lines:
    p1 = x.replace(';', '","')
    p3 = p1.replace("/", '-')
    p2 = p3.replace(" ", "")
    p4 = p2[:10] + " " + p2[11:]
    r1 = re.findall(dateModify, p4)
    try:
        if r1[0][1]:
            if r1[0][1] == '01':
                p5 = re.sub(dateModify, r"\1-Jan-\3", p4)
            if r1[0][1] == '02':
                p5 = re.sub(dateModify, r"\1-Feb-\3", p4)
            if r1[0][1] == '03':
                p5 = re.sub(dateModify, r"\1-Mar-\3", p4)
            if r1[0][1] == '04':
                p5 = re.sub(dateModify, r"\1-Apr-\3", p4)
            if r1[0][1] == '05':
                p5 = re.sub(dateModify, r"\1-May-\3", p4)
            if r1[0][1] == '06':
                p5 = re.sub(dateModify, r"\1-Jun-\3", p4)
            if r1[0][1] == '07':
                p5 = re.sub(dateModify, r"\1-Jul-\3", p4)
            if r1[0][1] == '08':
                p5 = re.sub(dateModify, r"\1-Aug-\3", p4)
            if r1[0][1] == '09':
                p5 = re.sub(dateModify, r"\1-Sep-\3", p4)
            if r1[0][1] == '10':
                p5 = re.sub(dateModify, r"\1-Oct-\3", p4)
            if r1[0][1] == '11':
                p5 = re.sub(dateModify, r"\1-Nov-\3", p4)
            if r1[0][1] == '12':
                p5 = re.sub(dateModify, r"\1-Dec-\3", p4)
            p6 = p5.replace(" ", '","')
            supString = '"","'
            p7 = supString + p6
            r2 = re.findall(regGroup, p7)
            temp5 = float(r2[0][2]) * 2
            r3 = [(r2[0][0], r2[0][1], temp5/1000, float(r2[0][3])/1000,
                   float(r2[0][4])/1000, float(r2[0][5])*2/1000, float(r2[0][6])/1000, float(r2[0][7])/1000)]
            p9 = p7[:22] + ',"blank",' + p7[23:]
            # print(r2)
            print(r3)
            # print(p9)
            p10 = '"","' + str(r3[0][0]) + '","' + str(r3[0][1]) + '","blank' + '","' + str(r3[0][2]) + '","' + str(
                r3[0][3]) + '","' + str(r3[0][4]) + '","' + str(r3[0][5]) + '","' + str(r3[0][6]) + '","' + str(r3[0][7]) + '"'
            f.write(p10 + '\n')
    except IndexError:
      print("oops")
