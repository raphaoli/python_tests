import numpy as np
import pandas as pd
import os
import xlwings as xw
import datetime as dt
import time


list_sheets_to_extract_kumba = ['Tot_Summary', 'Sish_Sched', 'Kolo_Sched' ,'Sish_Fleet_Capex','Kolo_Fleet_Capex']
dir_fol_kumba = r'C:\Users\IRACIC\OneDrive - Anglo American\Documents\Anglo American\Projects\working\Kumba'


caseSummaries_startsWith = "sc_"
dir_fol_kumba = r'C:\Users\IRACIC\OneDrive - Anglo American\Documents\Anglo American\Projects\working\Kumba'
dir_file_kumba = r'C:\Users\IRACIC\OneDrive - Anglo American\Documents\Anglo American\Projects\working\Kumba\20230818_LOM23_v1_template_VG.xlsx'


def get_latest_folder(directory):
    # Get all entries in the directory
    all_entries = os.listdir(directory)

    # Filter only directories
    folders = [entry for entry in all_entries if os.path.isdir(os.path.join(directory, entry))]

    # If there are no folders, return None
    if not folders:
        return None

    # Sort directories based on their creation time
    latest_folder = max(folders, key=lambda folder: os.path.getctime(os.path.join(directory, folder)))
    print(latest_folder)
    return latest_folder



def func_xl_ex_stack_kumba(dir_folder_read, dir_folder_output):
    # Program Start
    time_start = time.time()
    now = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    for f in os.listdir(dir_folder_read):
        file = os.path.join(dir_folder_read, f)
        print(f"f:{f}")
        print(f"file:{file}")
        print("_"*40)

        list_sheets_to_extract_kumba = ['Tot_Summary', 'Sish_Sched', 'Kolo_Sched', 'Sish_Fleet_Capex',
                                        'Kolo_Fleet_Capex']

        dir_name = now + " Kumba input"
        dir_kumba_input = os.path.join(dir_folder_output, dir_name)
        os.mkdir(dir_kumba_input)
        print(dir_name)
        print(dir_kumba_input)

        for sheet in list_sheets_to_extract_kumba:


            dir_file_kumba = file
            df_tot_sum = pd.read_excel(dir_file_kumba, list_sheets_to_extract_kumba[0]).fillna("")
            df_sish = pd.read_excel(dir_file_kumba, list_sheets_to_extract_kumba[1]).fillna("")
            df_kol = pd.read_excel(dir_file_kumba, list_sheets_to_extract_kumba[2]).fillna("")

            blank_df = pd.DataFrame(columns=["column1"], index=range(3))
            df_concat = pd.concat([df_tot_sum, blank_df, df_sish, blank_df, df_kol, blank_df, ], axis=0)
            position = 0

            # Inserting extra columns to get alignment on the years
            for i in range(5):
                df_concat.insert(position + i, f'blank_{i + 1}', np.nan)
            df_concat.fillna('')

            # _________
            df_sish_capex = pd.read_excel(dir_file_kumba, sheet_name=list_sheets_to_extract_kumba[3]).fillna("")
            df_kol_capex = pd.read_excel(dir_file_kumba, sheet_name=list_sheets_to_extract_kumba[4]).fillna("")
            df_capex = pd.DataFrame()

            df_capex = pd.concat([df_sish_capex, blank_df, df_kol_capex])
            df_all = df_concat.copy()

            start_row = df_all.shape[0]
            start_col = 0

            end_row = start_row + df_capex.shape[0]
            end_col = start_col + df_capex.shape[1]

            # Add rows if needed
            num_rows_to_add = end_row - df_all.shape[0]
            if num_rows_to_add > 0:
                df_all = pd.concat([df_all, pd.DataFrame(index=range(num_rows_to_add))], axis=0)

            # Add columns if needed
            num_cols_to_add = end_col - df_all.shape[1]
            if num_cols_to_add > 0:
                new_cols = pd.DataFrame(columns=range(df_all.shape[1], end_col))
                df_all = pd.concat([df_all, new_cols], axis=1)

            df_all.iloc[start_row:end_row, start_col:end_col] = df_capex.values

            df_all.fillna("")

            print(df_all)
            print(f"dir_kumba_input:{dir_kumba_input}")
            file_name = os.path.join(dir_kumba_input, now + " Kumba.xlsx")
            df_all.to_excel(file_name, index=False, engine="openpyxl")

    time_end = time.time()
    duration_min = round((time_end - time_start) / 60, 2)
    print(f"Program complete\nTotal run time {duration_min} min")


# func_xl_ex_stack_kumba(dir_folder_read=r"C:\Users\IRACIC\OneDrive - Anglo American\Documents\Anglo American\Projects\working\Kumba\Period Schedules\202309081215",
#                        dir_folder_output=r"C:\Users\IRACIC\OneDrive - Anglo American\Documents\Anglo American\Projects\working\Kumba")


def func_xl_ex_case_summaries(dir_folder_read, dir_folder_output):
    # Program Start
    time_start = time.time()
    now = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    list_files = []
    counter = 1

    noww = now
    # noww = ""

    for f in os.listdir(dir_folder_read):
        file = os.path.join(dir_folder_read, f)
        print(f"f:{f}")
        print(f"file:{file}")
        print("_"*40)

        sheets_start_with = "CS_"
        brNo_brNa_scNo_sc_Na_loc = "A6"

        dir_name = now + " Case Summaries input"
        dir_cs_input = os.path.join(dir_folder_output, dir_name)
        os.mkdir(dir_cs_input)
        print(dir_name)
        print(dir_cs_input)

        df_xl = pd.ExcelFile(file)
        sheet_names = df_xl.sheet_names
        relevant_sheets = [sheet for sheet in df_xl.sheet_names if sheet.startswith(sheets_start_with)]
        df_xl.close()


        print("_"*50)
        print(f"sheet_names:{sheet_names}")
        print("_"*25)
        print(f"relevant_sheets:{relevant_sheets}")
        print("="*50)

        for sheet in relevant_sheets:
            df = pd.read_excel(file, sheet, header=None).fillna("")

            description = df.iloc[5, 0]
            print(f"====>>> description:{description}")
            print(f"sheet:{sheet}")
            print()
            print(f"df.shape:{df.shape}")
            print(df.head())
            print(df.tail())
            # file_name = os.path.join(dir_cs_input, now + " Case Summaries.xlsx")
            file_name = os.path.join(dir_cs_input,noww + description + f"{ sheet}" + ".xlsx")
            print("_"*20)
            print(file_name)

            df.to_excel(file_name, index=False, engine="openpyxl")

            list_files.append(file_name)
            counter += 1
    for i in list_files:
        print(i)

    time_end = time.time()
    duration_min = round((time_end - time_start) / 60, 2)
    print(f"Program complete\nTotal run time {duration_min} min")

func_xl_ex_case_summaries(dir_folder_read=r"C:\Users\raphael.oliveira\OneDrive - Anglo American\Desktop\Test_Igor_Code\File",
                          dir_folder_output=r"C:\Users\raphael.oliveira\OneDrive - Anglo American\Desktop\Test_Igor_Code")