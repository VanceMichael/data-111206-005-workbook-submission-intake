"""Excel 序列日期转换，严格区分 1900 与 1904 日期制。

1900 系统的序列 60 是 Excel 为兼容 1900 年虚构闰日保留的占位，不是任何真实日期，
必须显式报告为非法，而不是伪造出 1900-02-29。1904 系统以 1904-01-01 为零点，
且不存在该虚构日。
"""
from __future__ import annotations

from datetime import date, timedelta
from decimal import Decimal

EPOCH_1900 = date(1899, 12, 31)
EPOCH_1904 = date(1904, 1, 1)
EXCEL_LEAP_BUG_SERIAL = 60


class DateError(ValueError):
    """序列值无法按工作簿日期制解释。"""


def convert_serial(serial: Decimal, date1904: bool) -> date:
    if serial != serial.to_integral_value():
        raise DateError(f'日期列出现非整数序列值: {serial}')
    value = int(serial)
    if date1904:
        if value < 0:
            raise DateError(f'1904 日期制不支持负序列值: {value}')
        return EPOCH_1904 + timedelta(days=value)
    if value <= 0:
        raise DateError(f'1900 日期制序列值必须为正: {value}')
    if value == EXCEL_LEAP_BUG_SERIAL:
        raise DateError('序列 60 是 1900 系统的虚构闰日占位，不是真实日期')
    # 序列 1..59 与真实日期一一对应；61 起 Excel 比真实计数多一天，需要扣回。
    if value < EXCEL_LEAP_BUG_SERIAL:
        return EPOCH_1900 + timedelta(days=value)
    return EPOCH_1900 + timedelta(days=value - 1)
