import openpyxl

# load the workbook
wb = openpyxl.load_workbook('spam_data.xlsx')

# select the active worksheet
ws = wb.active

# ask for user input
x = input("Enter a phone number: ")

# search for the input value in the worksheet
found = False
for row in ws.iter_rows(min_row=2):
    if row[0].value == x:
        found = True
        spam_value = row[8].value
        value = spam_value / 10
        print(f"{value} %. Chance of this number to be a spam")
        ans = input("Do you want to mark the number as spam (yes/no)? ")
        if ans.lower() == "yes":
            spam_value = row[8].value
            row[8].value = spam_value + 2
            print(f"Your respons has been saved.Thanks for your contribution ! ")
        break

if not found:
    print("Data not found.")

# save the updated workbook
wb.save('spam_data.xlsx')
