import unittest
from main import Salary, DataVacancy, InputConect, Report

class SalaryTests(unittest.TestCase):
    def test_salary_type(self):
        self.assertEqual(type(Salary(10.0, 20.4, True, 'RUR')).__name__, 'Salary')

    def test_salary_from(self):
        self.assertEqual(Salary(10.0, 20.4, True, 'RUR').salary_from, 10)

    def test_salary_to(self):
        self.assertEqual(Salary(10.9, 20.4, True, 'RUR').salary_to, 20)

    def test_salary_currency(self):
        self.assertEqual(Salary(10.0, 20.4, True, 'RUR').salary_currency, 'RUR')

    def test_int_get_salary(self):
        self.assertEqual(Salary(10, 20, True, 'RUR').convert_to_rub(), 15.0)

    def test_float_salary_from_in_get_salary(self):
        self.assertEqual(Salary(10.0, 20, True, 'RUR').convert_to_rub(), 15.0)

    def test_float_salary_to_in_get_salary(self):
        self.assertEqual(Salary(10, 30.0, True, 'RUR').convert_to_rub(), 20.0)

    def test_currency_in_get_salary(self):
        self.assertEqual(Salary(10, 30.0, True, 'EUR').convert_to_rub(), 1198.0)

class DataVacancyTests(unittest.TestCase):
    def test_data_vacancy_type(self):
        self.assertEqual(type(DataVacancy(
            'Программист',
            10,
            20,
            'RUR',
            'Москва',
            '2022-07-06T04:11:17+0300'
        )).__name__, 'DataVacancy')

class InputConnectTests(unittest.TestCase):
    def test_input_connect_type(self):
        ic = InputConect()
        self.assertEqual(type(InputConect()).__name__, 'InputConect')

class ReportTests(unittest.TestCase):
    def test_report_type(self):
        self.assertEqual(type(ReportTests()).__name__, 'Report')
