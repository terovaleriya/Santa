import random

import pandas

matching_table = 'Assignment.xlsx'
data_table = 'test.xlsx'

Surnames = pandas.read_excel(data_table, sheet_name='answers', usecols=['Фамилия'])

Surnames = Surnames['Фамилия'].tolist()

From = random.sample(Surnames, len(Surnames))
To = random.sample(Surnames, len(Surnames))

FromToSheet = pandas.DataFrame({"From": From, "To": To})

# print(FromToSheet)

FromToSheet.to_excel(matching_table)
