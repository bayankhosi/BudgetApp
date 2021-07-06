import datetime
import openpyxl as opx
import pandas as pd


budget = opx.load_workbook('./Files/budget.xlsx')
records = budget.worksheets[0]
month = int(datetime.datetime.now().strftime("%m"))  # month number


class month:

    def monthly_total():
        ro = int(input("""
        Enter Number of Month you wanna view
        (eg. May = 5)
        """))

        income_list = []
        income_type_list = []

        for row in records.iter_rows(min_row=5, max_row=7):

            cash = row[ro].value        # the cash
            income_list.append(cash)

            income_type = row[0].value
            income_type_list.append(income_type)

        income_dict = {'Amount': income_list}

        income_dataframe = pd.DataFrame(income_dict, index=income_type_list)
        print(income_dataframe)


#month.monthly_total()
