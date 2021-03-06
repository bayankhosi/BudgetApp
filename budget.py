import openpyxl as opx
import datetime
import upload
import analysis

budget = opx.load_workbook('./Files/budget.xlsx')
records = budget.worksheets[0]
month = int(datetime.datetime.now().strftime("%m"))  # month number

loop = 2


def rec_income():
    value = int(input("\nEnter Amount of Income\n"))
    income_type = int(input("""Which Type of Income\n
            [1] - Wages
            [2] - Interest/Dividends
            [3] - Misc
        """))
    cur_value = records.cell(row=4 + income_type, column=1 + month).value
    new_value = value + cur_value
    records.cell(row=4 + income_type, column=1 + month).value = new_value


def rec_expense():
    value = int(input("\nEnter Amount of Expense\n"))

    expense_catagory = int(input("""Which Type of Catagory\n
            [1] - Everyday Expenses
            [2] - Home
            [3] - Transport
            [4] - Vacation
            [5] - Recreation
            [6] - Subscriptions
            [7] - Personal
            [8] - Financial Obligation
            [9] - Misc
        """))

    if expense_catagory == 1:  # Everyday Expenses
        expense_type = int(input("""Which Type of Expense\n
            [1] - Groceries
            [2] - Restaurants
        """))
        rw = 19 + expense_type

    if expense_catagory == 2:  # Home
        expense_type = int(input("""Which Type of Expense\n
            [1] - Rent/Mortgage
            [2] - Insurance
            [3] - Repairs/Improvements
            [4] - Services
            [5] - Utilities
        """))
        rw = 11 + expense_type

    if expense_catagory == 3:  # Transport
        expense_type = int(input("""Which Type of Expense\n
            [1] - Public Transit

        """))
        rw = 28 + expense_type

    if expense_catagory == 4:  # Vacation
        expense_type = int(input("""Which Type of Expense\n
            [1] - Plane fare
            [2] - Accommodation
            [3] - Food
            [4] - Souvenirs
            [5] - Pet Boarding
            [6] - Transport
        """))
        rw = 54 + expense_type

    if expense_catagory == 5:  # Recreation
        expense_type = int(input("""Which Type of Expense\n
            [1] - Car
            [2] - Sax
            [3] - Misc
        """))
        rw = 63 + expense_type

    if expense_catagory == 6:  # Subscriptions
        expense_type = int(input("""Which Type of Expense\n
            [1] - Phone
            [2] - Internet
            [3] - Online Services
        """))
        rw = 70 + expense_type

    if expense_catagory == 7:  # Personal
        expense_type = int(input("""Which Type of Expense\n
            [1] - Clothing
            [2] - Barber
            [3] - Toiletry
            [4] - Gifts
            [5] - Charity
        """))
        rw = 80 + expense_type

    if expense_catagory == 8:  # FINANCIAL OBLIGATIONS

        expense_type = int(input("""Which Type of Expense\n
            [1] - Long-term savings
        """))
        rw = 88 + expense_type

    if expense_catagory == 9:  # Misc

        expense_type = int(input("""Which Type of Expense\n
            [1] - Loan Outs
        """))
        rw = 96 + expense_type

    cur_value = records.cell(row=rw, column=1 + month).value
    new_value = value + cur_value
    records.cell(row=rw, column=1 + month).value = new_value


while loop == 2:
    print("\n**************************************************************************************************\n")
    choice = int(
        input("Choose Operation:\n1. Record Expense  2. Record Income  3. View Data\n"))
    if choice == 1:
        rec_expense()
    elif choice == 2:
        rec_income()
    elif choice == 3:
        choice = int(input("""
                            [1] - Monthly
                            [2] - Yearly
                            """))
        if choice ==1:
            choice = int(input("""
                                [1] - Total
                                [2] - Catagory
                                """))
            analysis.month.monthly_total()

    budget.save('./Files/budget.xlsx')
    loop = int(input("1. Exit  2. Restart\n"))

upload.main()
