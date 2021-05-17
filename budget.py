import openpyxl as opx
import datetime

budget = opx.load_workbook('budget.xlsx')
expenses = budget.worksheets[1]
income = budget.worksheets[2]
month = int(datetime.datetime.now().strftime("%m"))  # month number

loop = 2


def rec_income():
    value = int(input("\nEnter Amount of Income\n"))
    income_type = int(input("""Which Type of Income\n
            [1] - Pay Slip
            [2] - Tips
            [3] - Bonus
            [4] - Commission
            [5] - Other
        """))
    cur_value = income.cell(row=3 + income_type, column=3 + month).value
    new_value = value + cur_value
    income.cell(row=3 + income_type, column=3 + month).value = new_value

    print("kak income")


def rec_expense():
    print("kak expense")


def data():
    print("kak data")


while loop == 2:
    print("\n**************************************************************************************************\n")
    choice = int(
        input("Choose Operation:\n1. Record Expense  2. Record Income  3. View Data\n"))
    if choice == 1:
        rec_income()
    elif choice == 2:
        rec_income()
    elif choice == 3:
        data()

    budget.save('budget.xlsx')
    loop = int(input("1. Exit  2. Restart\n"))
