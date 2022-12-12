import csv
import os

def separate(file_str):
    with open(file_str, encoding='utf-8-sig') as file:
        reader = csv.reader(file)
        head = []
        is_first = True
        years = {}
        for row in reader:
                if is_first:  
                    is_first = False
                    head = row
                else:
                    if not "" in row and len(row) == len(head):
                        year = row[5][:4]
                        if year in years.keys():
                            years[year].append(row)
                        else:
                            years[year] = [row]

    os.mkdir('years')

    for year in years.keys():
        with open(f'years//{year}.csv', 'w', encoding='utf-8-sig', newline='') as csvfile:
            filewriter = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_MINIMAL)
            filewriter.writerow(head)
            for item in years[year]:
                filewriter.writerow(item)