from main import Report, InputConect

choise = input('Что вывести?')

report = Report()

ic = InputConect()
data = ic.input_data()
data = list(data)

match choise:
    case 'Вакансии':
        report.generate_excel(data[:5], data[5:])
    case 'Статистика':
        report.generate_image(data)
