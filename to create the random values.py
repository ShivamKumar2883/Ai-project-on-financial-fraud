import random
from openpyxl import Workbook

# create a new workbook
wb = Workbook()

# select the active worksheet
ws = wb.active

# write headers
ws.append(['Phone Number', 'Location Value', 'Activation Date', 'Calls Per Hour', 
           'Diversity Value', 'Outgoing vs Incoming', 'Number Saved by People', 
           'Spam Report', 'Spam Value'])

# generate 200 rows of random data
for i in range(200):
    phone_num = random.choice(['6', '7', '8', '9']) + ''.join(random.choices('0123456789', k=9))
    location_value = random.randint(0, 100)
    activation_date = random.randint(0, 100)
    calls_per_hour = random.randint(0, 200)
    diversity_value = random.randint(0, 100)
    outgoing_vs_incoming = random.randint(0, 200)
    number_saved_by_people = random.randint(0, 100)
    spam_report = random.randint(0, 300)
    spam_value = location_value + activation_date + calls_per_hour + diversity_value + \
        outgoing_vs_incoming + number_saved_by_people + spam_report
    
    # write the row to the worksheet
    ws.append([phone_num, location_value, activation_date, calls_per_hour, 
               diversity_value, outgoing_vs_incoming, number_saved_by_people, 
               spam_report, spam_value])

# save the workbook
wb.save('spam_data.xlsx')

# ask the user for input
x = input("Enter the phone number : ")

# search for the input value in the worksheet
for row in ws.iter_rows(min_row=2):
    if row[0].value == x:
        spam_value = row[8].value
        print(f"The spam value for {x} is {spam_value}.")
        y = spam_value // 10
        print(f"{y}% of the spam value is {y/100}.")
        break
else:
    print("Data not found.")
