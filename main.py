import csv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.styles.borders import Border, Side
import matplotlib.pyplot as plt
import numpy as np
import pdfkit
from jinja2 import Environment, FileSystemLoader

class Salary:
    def __init__(self, salary_from, salary_to, salary_gross, salary_currency):
        self.salary_from = int(float(salary_from))
        self.salary_to = int(float(salary_to))
        self.salary_gross = salary_gross
        self.salary_currency = salary_currency

    def convert_to_rub(self):
        return (self.salary_from + self.salary_to) / 2 * self.__currency_to_rub[self.salary_currency]  

    __currency_to_rub = {  
        "AZN": 35.68,  
        "BYR": 23.91,  
        "EUR": 59.90,  
        "GEL": 21.74,  
        "KGS": 0.76,  
        "KZT": 0.13,  
        "RUR": 1,  
        "UAH": 1.64,  
        "USD": 60.66,  
        "UZS": 0.0055,  
    }       

class DataVacancy:
    def __init__(self, name, salary_from, salary_to, salary_currency, area_name, published_at):
        self.name = name
        self.salary = Salary(salary_from, salary_to, False, salary_currency)
        self.area_name = area_name
        self.published_at = published_at

class InputConect:
    def input_data(self):
        f = input('Введите название файла: ')

        job = input('Введите название профессии: ')

        salary_rub = self.__get_dict()
        salary_count = self.__get_dict()

        job_rub = self.__get_dict()
        job_count = self.__get_dict()

        data_objs = []

        with open(f, encoding='utf-8-sig') as file:
            reader = csv.reader(file)
            head = []

            is_first = True
            for row in reader:
                if is_first:  
                    is_first = False
                    head = row
                else:
                    if not "" in row and len(row) == len(head):

                        obj = DataVacancy(
                        row[head.index('name')], 
                        row[head.index('salary_from')], 
                        row[head.index('salary_to')], 
                        row[head.index('salary_currency')], 
                        row[head.index('area_name')], 
                        row[head.index('published_at')]
                        )
                        data_objs.append(obj)

                        year = int(obj.published_at[:4])
                        salary_rub[year] = self.__medium(salary_rub[year], obj.salary.convert_to_rub(), salary_count[year]) 
                        salary_count[year] += 1

                        if(obj.name.find(job) != -1):
                            job_rub[year] = self.__medium(job_rub[year], obj.salary.convert_to_rub(), job_count[year]) 
                            job_count[year] += 1

        salary_rub = self.__erase_empty(self.__round_values(salary_rub))
        job_rub = self.__erase_empty(self.__round_values(job_rub))
        salary_count = self.__erase_empty(salary_count)
        job_count = self.__erase_empty(job_count)

        print('Динамика уровня зарплат по годам:', salary_rub)
        print('Динамика количества вакансий по годам:', salary_count)
        print('Динамика уровня зарплат по годам для выбранной профессии:', job_rub)
        print('Динамика количества вакансий по годам для выбранной профессии:', job_count)
        
        city_salary = {}
        city_count = {}
        city_frac = {} 

        for it in data_objs:
            city = it.area_name
            if city not in city_salary.keys():
                if len([x for x in data_objs if x.area_name == city]) >= int(len(data_objs) / 100):
                    city_salary[city] = it.salary.convert_to_rub()
                    city_count[city] = 1
            else:
                city_salary[city] = self.__medium(city_salary[city], it.salary.convert_to_rub(), city_count[city])
                city_count[city] += 1

        all = len(data_objs)
        for key, value in city_count.items():
            city_frac[key] = round(value / (all / 100) / 100, 4)

        city_salary = self.__round_values(self.__erase_empty(self.__sort_city(city_salary)))
        city_frac = self.__erase_empty(self.__sort_city(city_frac))
        
        print('Уровень зарплат по городам (в порядке убывания):', city_salary)
        print('Доля вакансий по городам (в порядке убывания):', city_frac)

        return job, salary_rub, salary_count, job_rub, job_count, city_salary, city_frac

    def __get_dict(self):
        return {x: 0 for x in range(2007, 2023)}

    def __medium(self, m, x, n):
        return (m * n + x) / (n + 1)  

    def __sort_city(self, d):
        return dict(sorted(d.items(), key=lambda x: x[1], reverse=True)[:10])

    def __round_values(self, d):
        return dict(map(lambda x: (x[0], int(x[1])), d.items()))

    def __erase_empty(self, d):
        cd = dict(filter(lambda x:x[1], d.items()))
        if len(cd.keys()) == 0:
            cd[2022] = 0
        return cd
        
class Report:
    def __init__(self):
        self.wb = Workbook()

        for sheet_name in self.wb.sheetnames:
            sheet = self.wb[sheet_name]
            self.wb.remove(sheet)

        self.wb.create_sheet('Статистика по годам')
        self.wb.create_sheet('Статистика по городам')

    __first_headers = [
        'Год', 
        'Средняя зарплата', 
        'Средняя зарплата - ', 
        'Количество вакансий',
        'Количество вакансий - '
        ]

    def __as_text(self, value):
        if value is None:
            return ""
        return str(value)

    def __set_size(self):
        for column_cells in self.wb.active.columns:
            length = max(len(self.__as_text(cell.value)) for cell in column_cells)
            self.wb.active.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 2

    def __make_border(self):
        for row in self.wb.active.rows:
            for cell in row:
                cell.border = Border(
                    left=Side(style='thin'), 
                    right=Side(style='thin'), 
                    top=Side(style='thin'), 
                    bottom=Side(style='thin'))

    def __make_first_sheet(self, data):
        self.wb.active = self.wb['Статистика по годам']
        ws = self.wb.active
        
        self.__first_headers[2] = self.__first_headers[2] + data[0]
        self.__first_headers[4] = self.__first_headers[4] + data[0]
        ws.append(it for it in self.__first_headers)
        for row in ws.rows:
            for cell in row:
                cell.font = Font(bold=True)
        
        for year in data[1].keys():
            row = [year, data[1][year], data[3][year], data[2][year], data[4][year]]
            ws.append(row)

        self.__set_size()
        self.__make_border()
        
    __second_headers = [
        'Город',
        'Уровень зарплат',
        '',
        'Город',
        'Доля вакансий'
    ]

    def __make_second_sheet(self, data):
        self.wb.active = self.wb['Статистика по городам']
        ws = self.wb.active

        ws.append(it for it in self.__second_headers)
        for row in ws.rows:
            for cell in row:
                cell.font = Font(bold=True)
        
        info1 = list(data[0].keys())
        info2 = list(data[0].values())
        info3 = list(data[1].keys())
        info4 = list(data[1].values())

        for i in range(len(data[0])):
            row = [info1[i], info2[i], '', info3[i], info4[i]]
            ws.append(row)

        self.__set_size()
        self.__make_border()

        for i in range(1, 12):
            ws[f"C{i}"].border = Border()
        
        for i in range(1, 12):
            ws[f"E{i}"].number_format = '0.00%'

        self.wb.active = self.wb['Статистика по годам']

    def generate_excel(self, data1, data2):
        self.__make_first_sheet(data1)
        self.__make_second_sheet(data2)

        self.wb.save('report.xlsx')

    def __make_salary_year(self, job, data1, data2, ax):
        labels = list(data1.keys())
        average = list(data1.values())
        jobs = list(data2.values())

        x = np.arange(len(labels))
        width = 0.35
        
        ax.bar(x - width/2, average, width, label='средняя з/п')
        ax.bar(x + width/2, jobs, width, label=f"з/п {job}")    

        ax.set_title('Уровень зарплат по годам')
        ax.set_xticks(x, labels, rotation=90)
        ax.legend(prop={"size":8})
        ax.grid(axis='y')
        ax.tick_params(axis='both', labelsize=8)

    def __make_counts_year(self, job, data1, data2, ax):
        labels = list(data1.keys())
        counts = list(data1.values())
        jobs = list(data2.values())

        x = np.arange(len(labels))
        width = 0.35
        
        ax.bar(x - width/2, counts, width, label='Количество вакансий')
        ax.bar(x + width/2, jobs, width, label=f"Количество вакансий\n{job}")    

        ax.set_title('Количество вакансий по годам')
        ax.set_xticks(x, labels, rotation=90)
        ax.legend(prop={"size":8})
        ax.grid(axis='y')
        ax.tick_params(axis='both', labelsize=8)

    def __make_salary_city(self, data, ax):
        sep = lambda x: x.replace(' ', '\n').replace('-', '\n')

        cities = list(map(sep, data.keys()))[::-1]
        values = list(data.values())[::-1]
        y_pos = np.arange(len(cities))

        ax.barh(y_pos, values)
        ax.set_yticks(y_pos, labels=cities, fontsize=6)
        ax.set_title('Уровень зарплат по городам')
        ax.tick_params(axis='x', labelsize=8)

    def __make_jobs_count(self, data, ax):
        x = list(data.values())
        x.append(1 - sum(x))
        cities = list(data.keys()) + ['Другие']

        ax.set_title('Доля вакансий по городам')
        ax.pie(x, labels = cities, textprops={'fontsize': 6}, startangle=90)

    def generate_image(self, data):
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(nrows=2, ncols=2)

        self.__make_salary_year(data[0], data[1], data[3], ax1)
        self.__make_counts_year(data[0], data[2], data[4], ax2)
        self.__make_salary_city(data[5], ax3)
        self.__make_jobs_count(data[6], ax4)
        
        fig.tight_layout()
        fig.savefig('graph.png')
        
    def generate_pdf(self, data):
        job = data[0]
        image_file = "graph.png"

        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template("pdf_template.html")

        pdf_template = template.render({'job': job, 'image_file': image_file})

        config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')
        pdfkit.from_string(pdf_template, 'report.pdf', configuration=config)

ic = InputConect()
data = ic.input_data()

data = list(data)

data[1] = {2007: 38916, 2008: 43646, 2009: 42492, 2010: 43846, 2011: 47451, 2012: 48243, 2013: 51510, 2014: 50658}
data[2] = {2007: 2196, 2008: 17549, 2009: 17709, 2010: 29093, 2011: 36700, 2012: 44153, 2013: 59954, 2014: 66837}
data[3] = {2007: 43770, 2008: 50412, 2009: 46699, 2010: 50570, 2011: 55770, 2012: 57960, 2013: 58804, 2014: 62384}
data[4] = {2007: 317, 2008: 2460, 2009: 2066, 2010: 3614, 2011: 4422, 2012: 4966, 2013: 5990, 2014: 5492}
data[5] = {"Москва": 57354, "Санкт-Петербург": 46291, "Новосибирск": 41580, "Екатеринбург": 41091, "Казань": 37587, "Самара": 34091, "Нижний Новгород": 33637, "Ярославль": 32744, "Краснодар": 32542, "Воронеж": 29725}
data[6] = {"Москва": 0.4581, "Санкт-Петербург": 0.1415, "Нижний Новгород": 0.0269, "Казань": 0.0266, "Ростов-на-Дону": 0.0234, "Новосибирск": 0.0202, "Екатеринбург": 0.0143, "Воронеж": 0.014, "Самара": 0.0133, "Краснодар": 0.0131}
