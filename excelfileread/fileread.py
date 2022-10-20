import pathlib
import pandas as pd
import dataset
import os
from tkinter import *
from tkinter import filedialog
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler


def main():
    # read current working directory of the project
    path = str(pathlib.Path().resolve())
    current_date_time = str(datetime.now()).split(" ")[0]
    path_exists = os.path.exists(path)
    if not path_exists:
        os.makedirs(path)
    try:
        # answer = filedialog.askopenfilename(multiple=True)
        file_list = ["D:/DB1.xlsx", "D:/DB2.xlsx", "D:/DB.xlsx"]

        # check for file extension and read data
        for extension in file_list:
            file_extension = extension.split(".")[1]
            if file_extension == "xlsx":
                new_excel_list = []
                # append all the files in list as dataframe
                for file in file_list:
                    new_excel_list.append(pd.read_excel(file))
                merged_excel = pd.DataFrame()
                # merge all the appended dataframe to new
                for latest_merged_excel in new_excel_list:
                    merged_excel = merged_excel.append(latest_merged_excel, ignore_index=True)
                file_name = "LATEST_ROUND_OFF_DATA_" + current_date_time
                # convert the dataframe to new excel file
                merged_excel.to_excel(path + "\\" + file_name + ".xlsx", sheet_name=dataset.fileread_sheet_name)
                answer = path + "\\" + file_name + ".xlsx"

            else:
                pass
    except FileNotFoundError:
        print("File Not Found")

    worksheet = dataset.fileread_sheet_name

    df = pd.read_excel(answer, worksheet)
    df2 = df.groupby(['Purpose', 'Center'])['R off'].sum()
    print("File has been saved to current working directory !!!")
    filename = answer.split(":")[1].split("\\")[-1].split(".xlsx")[0]
    df2.to_excel(path + "\\" + "PURPOSE_AND_CENTER_BASED_ROUND_OFF_" + current_date_time + ".xlsx")


# scheduler = BlockingScheduler()
# scheduler.add_job(main, 'interval', seconds=30)
# scheduler.start()

if __name__ == "__main__":
    main()
