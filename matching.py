import random

import pandas

matching_table = 'Assignment.xlsx'
data_table = 'test.xlsx'

Surnames = pandas.read_excel(data_table, sheet_name='answers', usecols=['Фамилия'])

Surnames = Surnames['Фамилия'].tolist()

From = random.sample(Surnames, len(Surnames))
To = random.sample(Surnames, len(Surnames))

# check that random works
for i in range(len(Surnames)):
    if From[i] == To[i]:
        exit("Random sucks, re-run!")

FromToSheet = pandas.DataFrame({"From": From, "To": To})

# print(FromToSheet)

FromToSheet.to_excel(matching_table)
