import pandas as pd
import os
#from pathlib import Path, PureWindowsPath
#import glob
import tkinter as tk #required download package
from tkinter import filedialog

#create an instance of the pick path dialog where PDFs are
in_application_window = tk.Tk()
folder_root = filedialog.askdirectory(parent=in_application_window,
                                 initialdir=os.getcwd(),
                                 title="Please select the Box Sync folder where the NBHD style files reside:")
#Root structure for files
#folder_root = PureWindowsPath("C:\\Users\MBOND4\\Box Sync\\NDDCSNKRS_TEST\\")

#get the list of files in the folder
dir_list = os.listdir(folder_root)

in_application_window.destroy() #end the dialog

#Create lists
lst_catalyst = list()
lst_inline = list()

for f in dir_list:
    if f.endswith(".xlsx")  and not f.startswith('~'):
        file_to_open = os.path.join(folder_root,f)
        #print(file_to_open)
        wrkbk = pd.ExcelFile(file_to_open, engine='openpyxl')
        for sheets in wrkbk.sheet_names:
            if 'CATALYST' in sheets:
                catalyst_instance = pd.read_excel(file_to_open, engine='openpyxl', sheet_name=sheets)
                catalyst_instance['SEASON'] = f[0:4]
                catalyst_instance['MARKETING_TYPE'] = 'SNKRS + NBHD'
                lst_catalyst.append(catalyst_instance[['SEASON', 'MARKETING_TYPE','STYLE COLOR ', 'NRG COMMENTARY',
                                                       'FLOW DATE', 'CLASSIFICATION', 'NRG CLASSIFICATION',
                                                       'PROJECT NAME', 'PLUG', 'MARKETING NAME',
                                                       'DISTRIBUTION DETAILS', 'DISTRIBUTION INTENT',
                                                       'LAUNCH', ' NA TOTAL UNITS', ' EMEA TOTAL UNITS',
                                                       ' GC TOTAL UNITS', ' APLA TOTAL UNITS']])

            elif 'INLINE' in sheets:
                inline_instance = pd.read_excel(file_to_open, engine='openpyxl', sheet_name=sheets)
                inline_instance['SEASON'] = f[0:4]
                inline_instance['MARKETING TYPE'] = 'SNKRS + NBHD INLINE'
                lst_inline.append(inline_instance[['SEASON', 'MARKETING TYPE','Marketing Name', 'Product Code']])

dt_catalyst = pd.concat(lst_catalyst)
dt_inline = pd.concat(lst_inline)

dt_catalyst2 = dt_catalyst.rename({'STYLE COLOR ': 'PRODUCT_CD', 'NRG COMMENTARY': 'NRG COMMENTARY', 'FLOW DATE': 'FLOW_DATE',
                                   'NRG CLASSIFICATION': 'NRG_CLASSIFICATION', 'PROJECT NAME': 'PROJECT_NAME',
                                   'MARKETING NAME': 'MARKETING_NAME', 'DISTRIBUTION DETAILS': 'DISTRIBUTION_DETAILS',
                                   ' NA TOTAL UNITS': 'NA_TOTAL_UNITS', ' EMEA TOTAL UNITS': 'EMEA_TOTAL_UNITS',
                                   ' GC TOTAL UNITS': 'GC_TOTAL_UNITS', ' APLA TOTAL UNITS': 'APLA_TOTAL_UNITS'},
                                  axis=1)

dt_inline2 = dt_inline.rename({'MARKETING TYPE' : 'MARKETING_TYPE', 'Marketing Name': 'MARKETING_NAME',
                               'Product Code': 'PRODUCT_CD' }, axis=1)

# Create an instance of the pick path dialog where the merged PDF needs to go
out_application_window = tk.Tk()
folder_out = filedialog.askdirectory(parent=out_application_window,
                                 initialdir=os.getcwd(),
                                 title="Please select folder where you want to put the concatenated files:")

dt_catalyst2.to_csv(os.path.join(folder_out, 'Catalyst.csv'), index=False)
dt_inline2.to_csv(os.path.join(folder_out, 'Inline.csv'), index=False)

in_application_window.destroy() #end the dialog