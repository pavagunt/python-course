import pandas as pd
import numpy as np
#import re as str_edit
from tkinter import filedialog
from pandas.io.excel import ExcelWriter
import logging
import os
import xlwt
import re

##########################################################################
#iALM information
########################################################################
export_enable = True  #True - FaultMatrix data exported from iALM, False- FaultMatrix excel needs to be provided from local machine.
query="FM_BMW_HDT"
hostname = "mpt-ptc-ialm.magna.global"
port="7001"
output_file = 'C:\TMP\FaultMatrix_HDT.xls'  #Path to store the exported FaultMatrix data from iALM.
fields='"Category","State","ID","Reuses","Text","Fault Detection Criteria","SW Fault Handling Time Interval (Design) [ms]","Fault Reaction","Fault Reaction 2","Fault Reaction 3","Active Discharge","DTC Code","DTC Name","Error Event Name","Fault Recovery Criteria","SW Fault Recovery Time Interval (Design) [ms]","Parameterized By","Implementation Status","Operation Cycle","Enable Condition","Document ID","Customer Application Project","Monitor Type","Engineering Notes","Element","Verified By"'
#command_to_export_query = f'im exportissues --outputFile={output_file} --fields="Category","State","ID","Reuses","Text","Fault Detection Criteria","SW Fault Handling Time Interval (Design) [ms]","Fault Reaction","Fault Reaction 2","Fault Reaction 3","Active Discharge","DTC Code","DTC Name","Error Event Name","Fault Recovery Criteria","SW Fault Recovery Time Interval (Design) [ms]","Parameterized By","Implementation Status","Operation Cycle","Enable Condition","Document ID","Customer Application Project","Monitor Type","Engineering Notes","Element","Verified By" --query={query} --hostname={hostname} --port={port} --gui'
################################################################

################################################################
#Logger
################################################################
def logger(path):
    # Create and configure logger
    logging.basicConfig(filename=os.path.join(os.path.dirname(path), "FaultMatrix_updates.log"),                    
                        level=logging.DEBUG,
                        format='%(asctime)s %(message)s',
                        datefmt="%Y-%m-%d %H:%M:%S :",
                        filemode='w')

path = output_file if export_enable else filedialog.askopenfilename(filetypes=[('FaultMatrix File', '*.xls')])
logger(path)

def export_ialm_data():
    '''Exporting FaultMatrix iALM data'''
    command_to_export_query = f'im exportissues --outputFile={output_file} --noopenOutputFile --fields={fields} --query={query} --hostname={hostname} --port={port} --gui'
    try:
        logging.info(f"Exporting FaultMatrix data from iALM server '{hostname}'.....")
        print(f"Exporting FaultMatrix data from iALM server '{hostname}'.....")
        os.system(command_to_export_query)
        logging.info("Exported FaultMatrix data from iALM!!")
        print("Exported FaultMatrix data from iALM!!")
    except:
        print('Exception to export FaultMatrix iALM Data')
        logging.error("Exception to export FaultMatrix iALM Data")

export_ialm_data() if export_enable else ''

logging.info("FaultMatrix file: " + path)

if os.path.isfile(path):
    #Read FaultMatrix excel
    df = pd.read_excel(path, sheet_name=0, engine='xlrd')

    #Rename column names
    logging.info("\n\n########################################\n# Renaming column names\n########################################")
    df.rename(columns={'SW Fault Handling Time Interval (Design) [ms]':'Fault Detection Time Interval [ms]'}, inplace=True)
    logging.info("column name updated: 'SW Fault Handling Time Interval (Design) [ms]' ==> 'Fault Detection Time Interval [ms]'")
    df.rename(columns={'SW Fault Recovery Time Interval (Design) [ms]':'Fault Recovery Time Interval [ms]'}, inplace=True)
    logging.info("column name updated: 'SW Fault Recovery Time Interval (Design) [ms]' ==> 'Fault Recovery Time Interval [ms]'")
    df.rename(columns={'Error Event Name':'MonitorStateSignal'}, inplace=True)
    logging.info("column name updated: 'Error Event Name' ==> 'MonitorStateSignal'")

    #Remove '?' from reuses
    # logging.info("\n\n########################################\n# Removing '?' from 'Reuses' column\n########################################")
    # for i in range(0, len(list(df['Reuses']))):
    #     if '?' in df['Reuses'][i]:
    #         logging.info(f"Excel row {i+2}: Removing '?' from {df['Reuses'][i]}")
    #         df.loc[i, 'Reuses'] = int(df['Reuses'][i].replace('?', '').strip())

    df['Reuses'] = df['Reuses'].apply(lambda x: int(str(x).replace('?', '').strip()) if '?' in str(x) else x)

    df['Parameterized By'] = df['Parameterized By'].str.replace('?', '')

    #Add time in ms
    logging.info("\n\n################################################\n# Adding only time value in 'Fault Detection Time Interval [ms]' column\n####################################")
    FDTI_list = list(df['Fault Detection Time Interval [ms]'])
    for i in range(0, len(FDTI_list)):
        fdti = re.findall(r"([-+]?\d+)", FDTI_list[i].split("FDTI[ms]")[1].split('FRTI[ms]')[0].strip())
        df.loc[i, 'Fault Detection Time Interval [ms]'] = fdti[0] if fdti and 'no debouncing, detection as fast as possible' not in FDTI_list[i] else '0'
        frti = 0
        if "FRTI[ms]" in FDTI_list[i]:
            frti = FDTI_list[i].split("FRTI[ms]")[1].split('=')[1].replace('(no debouncing, detection as fast as possible)', '').strip()
        df.loc[i, 'Fault Reaction Time Interval [ms]'] = frti
        logging.info(f"Fault Detection Time Interval [ms]: '{FDTI_list[i]} ==> {fdti[0] if fdti and 'no debouncing, detection as fast as possible' not in FDTI_list[i] else 0}'")
        logging.info(f"Fault Reaction Time Interval [ms]: '{FDTI_list[i]} ==> {frti}'")

    #Add FRTI in ms
    logging.info("\n\n########################################\n# Fault Recovery Time Interval [ms] column\n########################################")
    FRTI_list = list(df['Fault Recovery Time Interval [ms]'])
    for i in range(0, len(FRTI_list)):
        if FRTI_list[i] is np.nan:
            df.loc[i, 'Fault Recovery Time Interval [ms]'] = 0
            logging.info(f"Excel row {i+2}: {FRTI_list[i]} ==> 0")
        elif 'Recovery time' in FRTI_list[i]:
            frti = str(FRTI_list[i].split("=")[1]).replace('ms', '')
            df.loc[i, 'Fault Recovery Time Interval [ms]'] = int(frti)
            logging.info(f"Excel row {i+2}: {FRTI_list[i]} ==> {frti}")
        elif 'Recovery_time_1' in FRTI_list[i]:
            frti = FRTI_list[i].split(",")
            for val in frti:
                if 'Recovery_time_1' in val:
                    df.loc[i, 'Fault Recovery Time Interval [ms]'] = int(val.split('=')[1].replace('ms' , '').strip())
                    logging.info(f"Excel row {i+2}: {FRTI_list[i]} ==> {val.split('=')[1].replace('ms' , '').strip()}")

    #Replace 'EVENT_' with 'V_RTE_SftyDiagEve_'
    logging.info("\n\n########################################\n# MonitorStateSignal\n########################################")
    pt_mo = list(df['MonitorStateSignal'])
    for i in range(len(pt_mo)):
        if 'EVENT_' in pt_mo[i]:
                df.loc[i, 'MonitorStateSignal'] = re.sub('EVENT_', 'V_RTE_SftyDiagEve_', pt_mo[i])
                logging.info(f"Excel row {i+2}: {pt_mo[i]} ==> {df['MonitorStateSignal'][i]}")
                
    # get the column to move           
    col_to_move = df.pop(df.columns[26])
    # insert the column at the desired index
    df.insert(7, col_to_move.name, col_to_move)  

    wb = xlwt.Workbook()
    sheet = wb.add_sheet('Sheet0')
    col_num = 0
    for column in list(df.columns):
        row_num = 1        
        sheet.write(0, col_num, column)
        for row in list(df[column]):
            sheet.write(row_num, col_num, row if (row is not np.nan) else '')
            row_num += 1
        col_num += 1

    #wb.save('C:\TMP\FaultMatrix_test.xls')
    wb.save(path)   #Path to store the updated FaultMatrix data.
    logging.info("\n\n!!FaultMatrix file is successfully updated!!")
    print('\n\nFaultMatrix file is successfully updated!!')
else:
    print('Please select the FaultMatrix file!!')