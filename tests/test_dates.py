import unittest
from datetime import date
from decimal import Decimal

from dates import DateError, convert_serial


class DateEpochTests(unittest.TestCase):
    def test_1904_epoch(self):
        self.assertEqual(convert_serial(Decimal('0'), True), date(1904, 1, 1))
        self.assertEqual(convert_serial(Decimal('60'), True), date(1904, 3, 1))
        self.assertEqual(convert_serial(Decimal('61'), True), date(1904, 3, 2))

    def test_1900_epoch_real_dates(self):
        self.assertEqual(convert_serial(Decimal('1'), False), date(1900, 1, 1))
        self.assertEqual(convert_serial(Decimal('59'), False), date(1900, 2, 28))
        self.assertEqual(convert_serial(Decimal('61'), False), date(1900, 3, 1))

    def test_1900_fictional_leap_day_is_not_fabricated(self):
        with self.assertRaises(DateError) as context:
            convert_serial(Decimal('60'), False)
        self.assertIn('虚构闰日', str(context.exception))

    def test_same_serial_different_epoch(self):
        # 同一序列 60：1904 系统是 3 月 1 日，1900 系统是必须拒绝的虚构日。
        self.assertEqual(convert_serial(Decimal('60'), True), date(1904, 3, 1))
        with self.assertRaises(DateError):
            convert_serial(Decimal('60'), False)

    def test_non_integer_and_non_positive(self):
        with self.assertRaises(DateError):
            convert_serial(Decimal('60.5'), False)
        with self.assertRaises(DateError):
            convert_serial(Decimal('0'), False)
        with self.assertRaises(DateError):
            convert_serial(Decimal('-1'), True)
