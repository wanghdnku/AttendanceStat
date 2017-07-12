import unittest
from utilities import *


class StringToTime(unittest.TestCase):

    def test_case_1(self):
        result = string_to_time('9:30:59')
        self.assertAlmostEqual(result, 9.5)

    def test_case_2(self):
        result = string_to_time('10:00:00')
        self.assertAlmostEqual(result, 10.0)

    def test_case_3(self):
        result = string_to_time('23:59:59')
        self.assertAlmostEqual(result, 23.9833333)

    def test_case_4(self):
        result = string_to_time('12:00:00')
        self.assertAlmostEqual(result, 12.0)

    def test_case_5(self):
        result = string_to_time('10:00:01')
        self.assertAlmostEqual(result, 10.0)


class ScheduleForDepartment(unittest.TestCase):

    def test_case_1(self):
        result = schedule_for_department('技术部', 9.5)
        self.assertAlmostEqual(result, (10.0, 18.5))

    def test_case_2(self):
        result = schedule_for_department('技术部', 9.6)
        self.assertAlmostEqual(result, (10.0, 18.6))

    def test_case_3(self):
        result = schedule_for_department('技术部', 9.7777)
        self.assertAlmostEqual(result, (10.0, 18.7777))

    def test_case_4(self):
        result = schedule_for_department('技术部', 10.0)
        self.assertAlmostEqual(result, (10.0, 19.0))

    def test_case_5(self):
        result = schedule_for_department('技术部', 10.5)
        self.assertAlmostEqual(result, (10.0, 19.0))

    def test_case_6(self):
        result = schedule_for_department('其他部', 8.9)
        self.assertAlmostEqual(result, (9.0, 18.0))

    def test_case_7(self):
        result = schedule_for_department('其他部', 9.0)
        self.assertAlmostEqual(result, (9.0, 18.0))

    def test_case_8(self):
        result = schedule_for_department('其他部', 9.1)
        self.assertAlmostEqual(result, (9.0, 18.0))


class TurnoutChecking(unittest.TestCase):

    def test_case_1(self):
        result = turnout_checking('其他部', '9:00:00', '18:00:00')
        self.assertAlmostEqual(result, '正常')

    def test_case_2(self):
        result = turnout_checking('其他部', '9:01:00', '18:00:00')
        self.assertAlmostEqual(result, '迟到')

    def test_case_3(self):
        result = turnout_checking('其他部', '9:30:00', '18:00:00')
        self.assertAlmostEqual(result, '迟到')

    def test_case_4(self):
        result = turnout_checking('其他部', '9:31:00', '18:00:00')
        self.assertAlmostEqual(result, '旷工')

    def test_case_5(self):
        result = turnout_checking('其他部', '9:00:00', '17:29:00')
        self.assertAlmostEqual(result, '旷工')

    def test_case_6(self):
        result = turnout_checking('其他部', '9:00:00', '17:30:00')
        self.assertAlmostEqual(result, '早退')

    def test_case_7(self):
        result = turnout_checking('其他部', '9:00:00', '17:31:00')
        self.assertAlmostEqual(result, '早退')

    def test_case_8(self):
        result = turnout_checking('其他部', '9:00:00', '17:59:00')
        self.assertAlmostEqual(result, '早退')

    def test_case_9(self):
        result = turnout_checking('其他部', '9:01:00', '17:59:00')
        self.assertAlmostEqual(result, '迟到早退')

    def test_case_10(self):
        result = turnout_checking('其他部', '9:31:00', '17:29:00')
        self.assertAlmostEqual(result, '旷工旷工')


if __name__ == '__main__':
    unittest.main()
