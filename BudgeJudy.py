import openpyxl
import pandas as pd
from openpyxl import Workbook, load_workbook
from datetime import date, datetime


#### Stored Data ####

# Transaction Descriptions #
descriptions = {

    "income": {"Income": ["TEND EXCHANGE", "ADVANTAGE WORK"]
    },

    "savings": {"Savings": []
    },

    "rent": {"Rent": []
    },

    "food": {
        "Groceries": ["RALPH", "VON", "WHOLEFDS", "TRADER JOE", "PAM", \
        "PAVIOLIONS", "STAR MARKET"],
        "Restaurant": ["CAFE", "CURRY", "SUSHI", "BELLA", "BISTRO", \
        "CUCINA", "TRATTORIA", "OSTERIA", "RESTAUR", "DOORDASH", "GRUBHUB",\
        "UBEREAT", "CLOVER", "BURGER", "YELLOW DOOR", "STARBUCKS", "TAPIOCA EXPRESS",\
        "KITCHEN", "DIRTY BIRDS", "CUISIN", "DOMINO\"S", "TACO", "PANDA EXPRESS",\
        "PLATES", "COFFEE", "JUICE"]
    },

    "drink": {
        "Counter": ["BEVERAGES AND MORE", "DISCOUNT WIN", "BEVERAGES", "LIQUOR"],
        "Bar": ["PUB MKT", "BEL VINO", "BREW", "CLURICAUNE", "REFRESHMENT",\
        "TIPSY CROW"]
    },

    "pharmacy": {"Pharmacy": ["CVS", "TARGET", "WALGREENS", "PHARM"]
    },

    "transportation": {
        "Parking": ["PARKING"],
        "Plane": ["JETBLUE", "DELTA"],
        "Train": ["AMTRAK", "TRENITALIA"],
        "Rideshare": ["UBER", "LYFT", "RIDEMOVI"]
    }
}


# Category Totals #
totals = {

    "income": {
        "Income": 0,
        "Total": 0
    },

    "savings": {
        "Savings": 0,
        "Total": 0
    },

    "rent": {
        "Rent": 0,
        "Total": 0
    },

    "food": {
        "Groceries": 0,
        "Restaurant": 0,
        "Total": 0
    },

    "drink": {
        "Counter": 0,
        "Bar": 0,
        "Total": 0
    },

    "pharmacy": {
        "Pharmacy": 0,
        "Total": 0
    },

    "transportation": {
        "Parking": 0,
        "Plane": 0,
        "Train": 0,
        "Rideshare": 0,
        "Total": 0
    },

    "miscellaneous": {
        "Miscellaneous": 0,
        "Total": 0}
}


# Categories tracked by month
bymonth_totals = {
    "Income": 0, 
    "Savings": 0, 
    "Rent": 0, 
    "Food": 0, 
    "Drink": 0, 
    "Pharmacy": 0, 
    "Transportation": 0, 
    "Miscellaneous": 0}


# Months and abbreviations #
months={
    "January": "jan", 
    "February": "feb",
    "March": "mar",
    "April": "apr",
    "May": "may",
    "June": "jun",
    "July": "jul",
    "August": "aug",
    "September": "sep",
    "October": "oct",
    "November": "nov",
    "December": "dec"}


# Converts number of month to name of month #
def month_name(month):
    names = {
        "01": "January",
        "02": "February",
        "03": "March",
        "04": "April",
        "05": "May",
        "06": "June",
        "07": "July",
        "08": "August",
        "09": "September",
        "10": "October",
        "11": "November",
        "12": "December",
        }
    return names[month] 


# Records bymonthly totals into new sheet and reset totals #
def record_bymonthly_totals():
    this_month = dater(row)[0:2]
    next_month = dater(row + 1)[0:2]

    if this_month != next_month:
        this_row = 20
        col_number = 73

        # write month name in bymonth chart
        month_column = budget_months.index(this_month)
        newsheet[chr(col_number + month_column) + str(this_row)] \
            = month_name(this_month)
        this_row += 1
        
        # write monthly totals in bymonth chart
        for bymonth_total in list(bymonth_totals.values()):
            newsheet[chr(col_number + month_column) + str(this_row)] \
            = bymonth_total
            this_row += 1

        # reset monthly totals
        for bymonth_total in list(bymonth_totals.keys()):
            bymonth_totals[bymonth_total] = 0


# Returns date without time #
def dater(row):
    day = sheet["A"+str(row)].value.strftime("%m/%d/%Y")
    return day


# If type is ATM returns ATM otherwise returns description #
def typer(row):
    if sheet["B"+str(row)].value == "ATM":
        return "ATM"
    else:
        return sheet["D"+str(row)].value


# Returns check number #
def checker(row):
    return sheet["C"+str(row)].value


# Returns description if found, otherwise deemed miscellaneous #
def descriptor(row):

    date = sheet["A"+str(row)].value.strftime("%m/%d/%Y")
    storecode = sheet["D"+str(row)].value

    if sheet["B"+str(row)].value == "ATM":
        return "ATM"

    for category in list(descriptions.items()):
        for subcategory in list(category[1].items()):
            for code in subcategory[1]:
                if code.lower() in storecode.lower():

                    # update by-month totals
                    bymonth_totals[category[0][0].upper() + category[0][1:]] \
                    += amounter(row)

                    record_bymonthly_totals()

                    # update subcategory total
                    totals[category[0]][subcategory[0]] += amounter(row)

                    # update category total
                    totals[category[0]]["Total"] += amounter(row)

                    # return category, subcategory and charge description
                    return [category[0], subcategory[0], storecode]

    record_bymonthly_totals()

    # update by-month totals
    bymonth_totals["Miscellaneous"] += amounter(row)

    # update category total
    totals["miscellaneous"]["Total"] += amounter(row)

    # updates miscellaneous category
    totals["miscellaneous"]["Miscellaneous"] += amounter(row)
    return ["miscellaneous", "Miscellaneous", storecode]


# Returns amount given or receieved (pay is pos, get is neg) #
def amounter(row):
    withdrawal = sheet["E"+str(row)].value
    if isinstance(withdrawal, float) or isinstance(withdrawal, int):
        return withdrawal
    else:
        return -sheet["F"+str(row)].value


# Returns remaining balance #
def balancer(row):
    return sheet["G"+str(row)].value








"""
######## User input: ########
Accesses spreadsheet
Changes name of original sheet
Makes new spreadsheet
Confirms columns
Sets start and end dates
"""

"""
# Accesses spreadsheet #
accessed = False
while accessed == False:
    try:
        print("What is the filepath for your sheet?")
        filepath = input() + ".xlsx"
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active
        accessed = True
    except FileNotFoundError as e:
        print("Sorry, that file could not be located. Please try again")
"""
# original_sheet_name = "PracticeCheckingSheet.xlsx"
# wb = openpyxl.load_workbook(original_sheet_name)
# sheet = wb.active

# Changes name of original sheet #
#print("Would you like to rename this sheet from "+str(wb.sheetnames[0])+"?")
#print("Provide the new name or press enter to keep original name")
#rename = input()
#if len(rename)>0:
#   sheet.title=rename

# Makes new spreadsheet
# print("What would you like to call your new sheet?")
# newsheet_name = input()
# wb.create_sheet(newsheet_name)
# wb.save(original_sheet_name)
# newsheet = wb[newsheet_name]


#### User enters start and end dates ####
startday = 0
startmonth = 0
startyear = 0
endyear = 0
endmonth = 0
endday = 0


# Start date info
print("Enter the starting date for your budget:")

# Collect starting year
while startyear == 0:
    print("Starting Year (YYYY)")
    year = input().strip()
    if len(year) == 2:
        startyear = "20" + year
    elif len(year) == 4:
        startyear = year
    else:
        print("Sorry! Your date was not in (YYYY) format. Please try again.")

# Collect starting month
while startmonth==0:
    print("Starting Month (MM)")
    month = input().strip()
    if len(month) == 1:
        month = "0" + month
    if 1<=int(month)<=12:
        startmonth = month

    else:
        print("Sorry! "+month+" isn\"t a valid month. Please try again.")

# Collect starting day
while startday == 0:
    print("Starting Day (DD)")
    day = input().strip()
    if len(day) == 1:
        day = "0" + day
    if 1 <= int(day) <= 31:
        startday = day
    else:
        print("Sorry! Your day is not in the range of the days in the month. Please try again.")


first_fiscal = str(startyear)+"-"+str(startmonth)+"-"+str(startday)





# End date info
print("Enter the ending date for your budget:")

print("Do you want to see your spending up to the present? (Y/N)")
answer = input()
if "y" in answer.lower():
    last_fiscal = str(date.today())
else:

    # Collect ending year
    while endyear == 0:
        print("Final Year (YYYY)")
        year = input().strip()
        if len(year) == 2:
            endyear = "20" + year
        elif len(year) == 4:
            endyear = year
        else:
            print("Sorry! Your date was not in (YYYY) format. Please try again.")

    # Collect ending month
    while endmonth==0:
        print("Final Month (MM)")
        month = input().strip()
        if len(month) == 1:
            month = "0" + month
        if 1<=int(month)<=12:
            endmonth = month
        else:
            print("Sorry! "+month+" isn\"t a valid month. Please try again.")

    # Collect ending day
    while endday == 0:
        print("Final Day (DD)")
        day = input().strip()
        if len(day) == 1:
            day = "0" + day
        if 1 <= int(day) <= 31:
            endday = day
        else:
            print("Sorry! Your day is not in the range of the days in the month. Please try again.")

    last_fiscal = str(endyear)+"-"+str(endmonth)+"/"+str(endday)





# The range of dates selected by the user
budget_days = pd.date_range(start=first_fiscal, end=last_fiscal)

# Makes a list of the months in budget_days 
budget_months = []
for day in budget_days:
    if str(day)[5:7] not in budget_months:
        budget_months.append(str(day)[5:7])











