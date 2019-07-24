#header file
import sys
import csv
import xlsxwriter
import datetime
import os.path
import glob #get latest update file
import re
import os
import xlrd
import openpyxl
from openpyxl import Workbook
import xlwings as xw

from splitIntoSheet import *

global output_path;#global variable, used by all function this file
output_path='C:\\Users\melody\\Desktop\\Rong'

global new_arrival_list_path;
new_arrival_list_path='Z:\SEA-AIR INFO CHECK'

global input_folder
input_folder='C:\\Users\\melody\\Desktop\\Python Tools\\Check Transaction\\'
