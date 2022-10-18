import pathlib
from subprocess import call
import matplotlib.pyplot as plt
import pandas as pd

# jepzzavjhsrsaasl


def main():
    try:
        print("Enter the file path : ")
        answer = input()
    except FileNotFoundError:
        print("File Not Found")
    try:
        print("Enter the Sheet Name :")
        worksheet = input()
    except IOError:
        print("Worksheet Not Found")

    # answer = "D:/C_Blood_MIS.xlsx"
    # filename = answer.split(":")[1].split("\\")[-1].split(".xlsx")[0]
    path = str(pathlib.Path().resolve())

    # pandas read Excel file along with workbook
    df = pd.read_excel(answer, worksheet)

    # get the total number of delivery dates from receipt date and booking date
    df['Receipt date'] = pd.to_datetime(df['Receipt date'], dayfirst=True)
    df['Booking Date'] = pd.to_datetime(df['Booking Date'], dayfirst=True)
    df['BD_TAT'] = df['Receipt date'] - df['Booking Date']

    # converting dataframe values to integer and replacing values
    df['BD_TAT'] = df['BD_TAT'].astype(str)
    df['BD_TAT'] = df['BD_TAT'].map(lambda x: x.rstrip('days'))
    df['BD_TAT'] = df['BD_TAT'].astype(int)

    # creating dataframe for time limit between 48 hours to 72 hours
    df48 = df[df['BD TAT'] == 2].groupby(['CENTER_CODE'])['BD_TAT'].count().reset_index(name="48_hrs_sample")
    df72 = df[df['BD TAT'] == 3].groupby(['CENTER_CODE'])['BD_TAT'].count().reset_index(name="72_hrs_sample")
    df_48 = df[df['BD TAT'] < 2].groupby(['CENTER_CODE'])['BD_TAT'].count().reset_index(name="less_48_hrs_sample")
    df_72 = df[df['BD TAT'] > 3].groupby(['CENTER_CODE'])['BD_TAT'].count().reset_index(name="greater_72_hrs_sample")

    # merge two dataframe on center_code
    df48_72 = pd.merge(df48, df72, on="CENTER_CODE")
    df_0_48 = pd.merge(df_48, df_72, on="CENTER_CODE")

    # add new column as 48_to_72hrs which is addition of columns (48_hrs_sample,72_hrs_sample)
    df48_72['48_to_72hrs'] = df48_72['48_hrs_sample'] + df48_72['72_hrs_sample']

    df_0_48['0_to_48hrs'] = df_0_48['less_48_hrs_sample']

    df_72['72_to_higher'] = df_72['greater_72_hrs_sample']

    # dataframe group-by center_code along with count
    df2 = df.groupby(['CENTER_CODE'])['CENTER_CODE'].count().reset_index(name="NO_OF_SAMPLES") \
        .sort_values('NO_OF_SAMPLES', ascending=False).reset_index(level=0, drop=False)

    # merge the dataframe df2 along with df48_72 for time between 48 and 72 hours
    df_48_to_72_hrs = pd.merge(df2, df48_72, on="CENTER_CODE").sort_values('NO_OF_SAMPLES', ascending=False)

    # merge the dataframe df2 along with df_0_48 for time between 0 and 48 hours
    df_0_48_hrs = pd.merge(df2, df_0_48, on="CENTER_CODE").sort_values('NO_OF_SAMPLES', ascending=False)

    # merge the dataframe df2 along with df_72 for time greater than 72 hours
    df_72_hrs = pd.merge(df2, df_72, on="CENTER_CODE").sort_values('NO_OF_SAMPLES', ascending=False)

    # create new column in dataframe for calculating the percentage and round off the values for time between 48 and 72
    df_48_to_72_hrs['48_to_72_hrs_%'] = (df_48_to_72_hrs['48_to_72hrs'] / df_48_to_72_hrs['NO_OF_SAMPLES']) * 100
    df_48_to_72_hrs['48_to_72_hrs_%'] = df_48_to_72_hrs['48_to_72_hrs_%'].apply(lambda x: round(x, 2))

    # create new column in dataframe for calculating the percentage and round off the values for time less than 48
    df_0_48_hrs['0_to_48_hrs_%'] = (df_0_48_hrs['0_to_48hrs'] / df_0_48_hrs['NO_OF_SAMPLES']) * 100
    df_0_48_hrs['0_to_48_hrs_%'] = df_0_48_hrs['0_to_48_hrs_%'].apply(lambda x: round(x, 2))

    # create new column in dataframe for calculating the percentage and round off the values for time greater than 72
    df_72_hrs['72_hrs_higher_%'] = (df_72_hrs['72_to_higher'] / df_72_hrs['NO_OF_SAMPLES']) * 100
    df_72_hrs['72_hrs_higher_%'] = df_72_hrs['72_hrs_higher_%'].apply(lambda x: round(x, 2))

    # rearranging the dataframe for 0 t0 48 hours
    df_0_to_48_hrs_df = df_0_48_hrs.loc[:, ['CENTER_CODE', '0_to_48hrs', '0_to_48_hrs_%', 'NO_OF_SAMPLES']] \
        .reset_index(level=0, drop=True)
    df_0_to_48_hrs_df = df_0_to_48_hrs_df.head(10)
    df_0_to_48_hrs_df.to_excel(path + "\\" + "SAMPLE_RECEIVED_LESS_THAN_48_HRS" + ".xlsx")
    df_0_to_48_hrs_df.plot.barh(x="CENTER_CODE")
    plt.title('TOP 10 CENTERS BY SAMPLE RECEIVED LESS THAN 48 HRS')
    plt.savefig(path + "\\" + "SAMPLE_RECEIVED_LESS_THAN_48_HRS" + ".png", dpi=100)

    # rearranging the dataframe for greater than 72 hours
    df_72_hrs_to_higher_df = df_72_hrs.loc[:, ['CENTER_CODE', '72_to_higher', '72_hrs_higher_%', 'NO_OF_SAMPLES']] \
        .reset_index(level=0, drop=True)
    df_72_hrs_to_higher_df = df_72_hrs_to_higher_df.head(10)
    df_72_hrs_to_higher_df.to_excel(path + "\\" + "SAMPLE_RECEIVED_AFTER_72_HRS" + ".xlsx")
    df_72_hrs_to_higher_df.plot.barh(x="CENTER_CODE")
    plt.title('TOP 10 CENTERS BY SAMPLE RECEIVED AFTER 72 HRS')
    plt.savefig(path + "\\" + "SAMPLE_RECEIVED_AFTER_72_HRS" + ".png", dpi=100)

    # rearranging the dataframe for time between 48 and 72 hours
    df_48_to_72_hrs_df = df_48_to_72_hrs.loc[:, ['CENTER_CODE', '48_to_72hrs', '48_to_72_hrs_%', 'NO_OF_SAMPLES']] \
        .reset_index(level=0, drop=True)
    df_48_to_72_hrs_df = df_48_to_72_hrs_df.head(10)
    df_48_to_72_hrs_df.to_excel(path + "\\" + "SAMPLE_RECEIVED_BETWEEN_48_TO_72_HRS" + ".xlsx")
    df_48_to_72_hrs_df.plot.barh(x='CENTER_CODE')
    plt.title('TOP 10 CENTERS BY SAMPLE RECEIVED MORE THAN 48-72 HRS')
    plt.savefig(path + "\\" + "SAMPLE_RECEIVED_BETWEEN_48_TO_72_HRS" + ".png", dpi=100)
    # plt.show()
    print("Files Have Been Saved Successfully In Current Working Directory !!")

    # calling mailsender to send all files
    call(["python", "mailSender.py"])


if __name__ == "__main__":
    main()
