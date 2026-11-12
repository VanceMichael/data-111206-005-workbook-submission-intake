"""工作簿容器入口：校验 OOXML 并列出工作表关系。

结构化采集的解析能力在 ``workbook_parser`` / ``collector`` 中实现，这里保留原有
``inventory`` 接口，供只需要工作表清单与日期制的调用方使用。
"""
from workbook_parser import parse_workbook


def inventory(path):
    workbook = parse_workbook(path)
    return {
        'sheets': [
            {'name': sheet.name, 'relationship': sheet.rel_id,
             'target': sheet.target, 'external': sheet.external}
            for sheet in workbook.sheets
        ],
        'members': workbook.members,
        'date1904': workbook.date1904,
    }
