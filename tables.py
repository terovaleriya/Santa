import random
import csv
import pandas

Surnames = pandas.read_excel('TCS.xlsx', sheet_name='Sheet1', usecols=['Surname'])

Surnames = Surnames['Surname'].tolist()

From = random.sample(Surnames, len(Surnames))
To = random.sample(Surnames, len(Surnames))

FromTo = dict(zip(From, To))

w = csv.writer(open("output.csv", "w"))
for key, val in FromTo.items():
    w.writerow([key, val])

print(FromTo)
