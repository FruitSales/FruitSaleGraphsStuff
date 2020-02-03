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

acc_clementines, acc_small_bask, acc_large_bask, acc_small_grape, acc_large_grape, acc_large_orange, acc_santa_orange, acc_small_orange, acc_red_apples, acc_gold_apples, acc_gala_apples,  \
    = [], [], [], [], [], [], [], [], [], [], [],



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

sums = []
sums.append(sum(acc_clementines))
sums.append(sum(acc_small_bask))
sums.append(sum(acc_large_bask))
sums.append(sum(acc_small_grape))
sums.append(sum(acc_large_grape))
sums.append(sum(acc_large_orange))
sums.append(sum(acc_santa_orange))
sums.append(sum(acc_small_orange))
sums.append(sum(acc_red_apples))
sums.append(sum(acc_gold_apples))
sums.append(sum(acc_gala_apples))
print(sums)

sums_18 = [1590.0, 1341.0, 471.0, 211, 47.0, 47.0, 310.0, 59.0, 77, 64.0, 161.0]
sums_19 = [1667.0, 1510.0, 475.0, 217, 49.0, 66.0, 366.0, 98.0, 89, 71.0, 173.0]


fig = go.Figure()
fig.add_trace(go.Bar(
    x=x_label,
    y=sums,
    name='2017 Data',
    marker_color='royalblue'
))
fig.add_trace(go.Bar(
    x=x_label,
    y=sums_18,
    name='2018 Data',
    marker_color='crimson'
))
fig.add_trace(go.Bar(
    x=x_label,
    y=sums_19,
    name='2019 Data',
    marker_color='darkseagreen'
))

offline.plot(fig, filename='Cumulative_Fruit_Bar.html')
