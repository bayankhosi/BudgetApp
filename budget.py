#!/usr/bin/env python3
import openpyxl as opx
import datetime
import upload
import analysis

budget = opx.load_workbook(analysis.budget)
records = budget.worksheets[0]
month = int(datetime.datetime.now().strftime("%m"))  # month number

loop = 2

def rec_income():
    value = int(input("\nEnter Amount of Income\n"))
    income_type = int(input("""Which Type of Income\n
            [1] - Wages
            [2] - Interest/Dividends
        """))
    cur_value = records.cell(row=1
     + month, column=1 + income_type).value
    new_value = value + cur_value
    records.cell(row=1 + month, column=1 + income_type).value = new_value

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
        """))
        col = 3 + expense_type

    if expense_catagory == 2:  # Home
        expense_type = int(input("""Which Type of Expense\n
            [1] - Rent/Mortgage
            [2] - Insurance
            [3] - Repairs/Improvements
            [4] - Utilities
        """))
        col = 4 + expense_type

    if expense_catagory == 3:  # Transport
        expense_type = int(input("""Which Type of Expense\n
            [1] - Public Transit
            [2] - Fuel
            [3] - Car Repairs
            [4] - Parking

        """))
        col = 8 + expense_type

    if expense_catagory == 4:  # Vacation
        expense_type = int(input("""Which Type of Expense\n
            [1] - Transport
            [2] - Accommodation
            [3] - Food
        """))
        col = 17 + expense_type

    if expense_catagory == 5:  # Recreation
        expense_type = int(input("""Which Type of Expense\n
            [1] - Hobbies
            [2] - Books/Games
            [3] - Movies/Concerts
        """))
        col = 12 + expense_type

    if expense_catagory == 6:  # Subscriptions
        expense_type = int(input("""Which Type of Expense\n
            [1] - Phone
            [2] - Internet
        """))
        col = 21 + expense_type

    if expense_catagory == 7:  # Personal
        expense_type = int(input("""Which Type of Expense\n
            [1] - Clothing
            [2] - Barber
            [3] - Toiletry
            [4] - Gifts
        """))
        col = 23 + expense_type

    if expense_catagory == 8:  # FINANCIAL OBLIGATIONS

        expense_type = int(input("""Which Type of Expense\n
            [1] - Long-term savings
            [2] - tax
        """))
        col = 26 + expense_type

    if expense_catagory == 9:  # Misc

        expense_type = int(input("""Which Type of Expense\n
            [1] - Loan Outs
        """))
        col = 28 + expense_type

    cur_value = records.cell(row=1 + month, column=col).value
    new_value = value + cur_value
    records.cell(row=1 + month, column=col).value = new_value


while loop == 2:
    print("\n********************************************************************************\n")
    choice = int(
        input("Choose Operation:\n 1. Record Expense\n 2. Record Income\n 3. View Data\n"))
    if choice == 1:
        print("\n********************************************************************************\n")
        rec_expense()
    elif choice == 2:
        print("\n********************************************************************************\n")
        rec_income()
    elif choice == 3:
        print("\n********************************************************************************\n")
        choice = int(input("""
            [1] - Monthly
            [2] - Yearly
        """))
        if choice ==1:
            choice = int(input("""
            [1] - Monthly Overview
            [2] - Catagory
            """))
            analysis.monthly.summary()

    budget.save(analysis.budget)
    loop = int(input("1. Exit  2. Restart\n"))

#upload.main()
