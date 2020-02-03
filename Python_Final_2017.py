import pyodbc
from plotly.graph_objs import Bar
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

data = [{
    'type': 'bar',
    'x': x_label,
    'y': sums,
    'marker': {
        'color': 'rgb(60, 100, 150)',
        'line': {'width': 1.5, 'color': 'rgb(25, 25, 25)'}
    },
    'opacity': 0.6,
}]

my_layout = {
    'title': '2017 Fruit Sale',
    'titlefont': {'size': 28},
    'xaxis': {
        'titlefont': {'size': 24},
        'tickfont': {'size': 14},
    },
    'yaxis': {
        'title': "Fruit Sold",
        'titlefont': {'size': 24},
        'tickfont': {'size': 14},
    },
}

fig = {'data': data, 'layout': my_layout}
offline.plot(fig, filename='fruit_sale_2017.html')
