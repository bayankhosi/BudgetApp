import datetime
import pandas as pd

df = pd.read_excel(
    './Files/budget.xlsx',
    sheet_name='Sheet1',
    index_col=0)

class monthly():

    def summary():

        month = int(input("Enter month number: "))

        switcher = {
            1: "JAN",
            2: "FEB",
            3: "MAR",
            4: "APR",
            5: "MAY",
            6: "JUN",
            7: "JUL",
            8: "AUG",
            9: "SEP",
            10: "OCT",
            11: "NOV",
            12: "DEC"}

        print('\n',switcher.get(month, "Invalid month"), '\n')
        mnt_name = switcher.get(month)

        print(df.loc[mnt_name])


# monthly.summary()