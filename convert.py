import pandas as pd
import csv

cur = pd.read_csv('currencies_df.csv', index_col='date').dropna()

convert = pd.DataFrame()

def add_info(name, salary, area_name, published_at):
    global convert
    convert = convert.append(pd.DataFrame([[salary, area_name, published_at]], [name], ['salary', 'area_name', 'published_at']))

with open('vacancies_dif_currencies.csv', encoding='utf-8-sig') as file:
        reader = csv.reader(file)
        head = []
        is_first = True
        for row in reader:
                if is_first:  
                    is_first = False
                    head = row
                else:
                    date = row[5][:7]
                    if date in cur.index.array and row[3] != '':
                        multi = 1 if row[3] == 'RUR' else cur.loc[date][row[3]]
                        if row[1] != '' and row[2] != '':
                            salary = (float(row[1]) + float(row[2])) / 2 * multi
                            add_info(row[0], salary, row[4], row[5])

                        if row[1] == '' and row[2] != '':
                            salary = float(row[2]) * multi
                            add_info( row[0], salary, row[4], row[5])

                        if row[1] != '' and row[2] == '':
                            salary = float(row[1]) * multi
                            add_info(row[0], salary, row[4], row[5])

convert.index.rename('name', inplace=True)

convert.to_csv('convert.csv')
