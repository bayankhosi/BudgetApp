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


def rec_expense():
    value = int(input("\nEnter Amount of Expense\n"))

    expense_catagory = int(input("""Which Type of Catagory\n
            [1] - Everyday Expenses
            [2] - Home
            [3] - Transport
            [4] - Utilities
            [5] - Education
            [6] - Entertainment
            [7] - Gifts
        """))

    if expense_catagory == 1:  # Everyday Expenses
        expense_type = int(input("""Which Type of Expense\n
            [1] - Groceries
            [2] - Restaurants
            [3] - Personal supplies
            [4] - Clothes
        """))
        rw = 44 + expense_type

    if expense_catagory == 2:  # Home
        expense_type = int(input("""Which Type of Expense\n
            [1] - Rent/mortgage
        """))
        rw = 69 + expense_type

    if expense_catagory == 3:  # Transport
        expense_type = int(input("""Which Type of Expense\n
            [1] - Public Transit

        """))
        rw = 105 + expense_type

    if expense_catagory == 4:  # Utilities
        expense_type = int(input("""Which Type of Expense\n
            [1] - Phone
            [2] - TV
            [3] - Internet
            [4] - Electricity
            [5] - Heat/Gas
            [6] - Water
        """))
        rw = 124 + expense_type

    if expense_catagory == 5:  # Education
        expense_type = int(input("""Which Type of Expense\n
            [1] - Tuition
            [2] - Books
        """))
        rw = 22 + expense_type

    if expense_catagory == 6:  # Entertainment
        expense_type = int(input("""Which Type of Expense\n
            [1] - Books
            [2] - Concerts
            [3] - Games
            [4] - Hobbies
            [5] - Movies
            [6] - Music
            [7] - Outdoor activities
            [8] - Photography
            [9] - Sport
            [10] - Theatre/plays
            [11] - TV
            [12] - Other
        """))
        rw = 29 + expense_type

    if expense_catagory == 7:  # Gifts
        expense_type = int(input("""Which Type of Expense\n
            [1] - Gifts
            [2] - Charity
            [3] - Other
        """))
        rw = 55 + expense_type

    cur_value = expenses.cell(row=rw, column=3 + month).value
    new_value = value + cur_value
    expenses.cell(row=rw, column=3 + month).value = new_value


def data():
    print("kak data")


while loop == 2:
    print("\n**************************************************************************************************\n")
    choice = int(
        input("Choose Operation:\n1. Record Expense  2. Record Income  3. View Data\n"))
    if choice == 1:
        rec_expense()
    elif choice == 2:
        rec_income()
    elif choice == 3:
        data()

    budget.save('budget.xlsx')
    loop = int(input("1. Exit  2. Restart\n"))
