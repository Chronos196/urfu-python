from main import *

choise = input('Что вывести?')

report = Report()

match choise:
    case 'Вакансии':
        report.generate_excel(data[:5], data[5:])
    case 'Статистика':
        report.generate_image(data)
#Изменение 1