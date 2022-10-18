import pathlib
import pandas as pd


def main():
    try:
        print("Enter the file path : ")
        answer = input()
        print("Enter the Sheet Name :")
        worksheet = input()
    except FileNotFoundError:
        print("File Not Found")

    df = pd.read_excel(answer, worksheet)
    df2 = df.groupby(['Purpose', 'Center'])['R off'].sum()
    path = str(pathlib.Path().resolve())
    print("File has been saved to current working directory !!!")
    filename = answer.split(":")[1].split("\\")[-1].split(".xlsx")[0]
    df2.to_csv(path + "\\" + filename + ".csv")


if __name__ == "__main__":
    main()
