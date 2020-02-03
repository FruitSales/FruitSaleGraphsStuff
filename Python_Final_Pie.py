import pyodbc
import plotly.graph_objects as go
from plotly import offline

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\tlam512q12/Documents\Python\Chpt17\data\2017AnthisFruit.accdb;'
)
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()
crsr.execute('select * from FruitSale')
clementines, small_bask, large_bask, small_grape, large_grape, large_orange, santa_orange, small_orange, red_apples, gold_apples, gala_apples, amount_owed \
    = [], [], [], [], [], [], [], [], [], [], [], []
columns = [column[0] for column in crsr.description]
for row in crsr.fetchall():
    clementines.append(row[4])
    small_bask.append(row[5])
    large_bask.append(row[6])
    small_grape.append(row[7])
    large_grape.append(row[8])
    large_orange.append(row[9])
    santa_orange.append(row[10])
    small_orange.append(row[11])
    red_apples.append(row[12])
    gold_apples.append(row[13])
    gala_apples.append(row[15])
    amount_owed.append(row[14])
columns.remove('AmountOwed')
columns.remove('ID')
columns.remove('teacherCode')
columns.remove('StudentLastName')
columns.remove('StudentFirstName')
columns.remove('Sheet')
x_label = columns
print(x_label)

acc_clementines, acc_small_bask, acc_large_bask, acc_small_grape, acc_large_grape, acc_large_orange, acc_santa_orange, acc_small_orange, acc_red_apples, acc_gold_apples, acc_gala_apples, acc_amount_owed\
    = [1590.0, 1667.0, ], [1341.0, 1510.0, ], [471.0, 217, ], [211, 475.0, ], [47.0, 49.0, ], [47.0, 66.0, ], [310.0,
                                                                                                               366.0, ], [
          59.0, 98.0, ], [77, 89, ], [64.0, 71.0, ], [161.0, 173.0], []

for value in clementines:
    if value is not None:
        acc_clementines.append(value)

for value in small_bask:
    if value is not None:
        acc_small_bask.append(value)

for value in large_bask:
    if value is not None:
        acc_large_bask.append(value)

for value in small_grape:
    if value is not None:
        acc_small_grape.append(value)

for value in large_grape:
    if value is not None:
        acc_large_grape.append(value)

for value in large_orange:
    if value is not None:
        acc_large_orange.append(value)

for value in santa_orange:
    if value is not None:
        acc_santa_orange.append(value)

for value in small_orange:
    if value is not None:
        acc_small_orange.append(value)

for value in red_apples:
    if value is not None:
        acc_red_apples.append(value)

for value in gold_apples:
    if value is not None:
        acc_gold_apples.append(value)

for value in gala_apples:
    if value is not None:
        acc_gala_apples.append(value)

for value in amount_owed:
    if value is not None:
        acc_amount_owed.append(value)
sums = []
total_raised = []
sums.append(sum(acc_clementines) * 8)
sums.append(sum(acc_small_bask) * 25)
sums.append(sum(acc_large_bask) * 39)
sums.append(sum(acc_small_grape) * 20)
sums.append(sum(acc_large_grape) * 35)
sums.append(sum(acc_large_orange) * 20)
sums.append(sum(acc_santa_orange) * 30)
sums.append(sum(acc_small_orange) * 40)
sums.append(sum(acc_red_apples) * 25)
sums.append(sum(acc_gold_apples) * 25)
sums.append(sum(acc_gala_apples) * 25)
total_raised.append(sum(acc_amount_owed))
print(sums)

#print(acc_amount_owed)
fig = go.Figure(data=[go.Pie(labels=x_label, values=sums)])
offline.plot(fig, filename='Fruit_Sale_Pie.html')
