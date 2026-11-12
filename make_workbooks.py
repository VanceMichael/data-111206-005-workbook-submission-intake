from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def cell(reference, value, kind=None, formula=None, formula_type=None, no_cache=False):
    attribute = f' t="{kind}"' if kind else ''
    if formula_type:
        expression = f'<f t="{formula_type}">{formula}</f>'
    elif formula:
        expression = f'<f>{formula}</f>'
    else:
        expression = ''
    if formula and no_cache:
        return f'<c r="{reference}"{attribute}>{expression}</c>'
    stored = f'<is><t>{value}</t></is>' if kind == 'inlineStr' else f'<v>{value}</v>'
    return f'<c r="{reference}"{attribute}>{expression}{stored}</c>'


def empty_string_cell(reference):
    # 显式空字符串：t="inlineStr" 且 <is> 内没有文本节点，区别于缺失的空单元格。
    return f'<c r="{reference}" t="inlineStr"><is></is></c>'


def error_cell(reference, code):
    return f'<c r="{reference}" t="e"><v>{code}</v></c>'


def blank_row(number, body):
    # 行存在但不包含全部列，缺失单元格即空单元格。
    return f'<row r="{number}">{body}</row>'


def sheet(rows):
    return f'<worksheet xmlns="{MAIN}"><sheetData>' + ''.join(f'<row r="{number}">{body}</row>' for number, body in enumerate(rows, 1)) + '</sheetData></worksheet>'


def org_sheet():
    return sheet([cell('A1', '企业编号', 'inlineStr') + cell('B1', '企业名称', 'inlineStr'),
                  cell('A2', '0007', 'inlineStr') + cell('B2', '合成甲', 'inlineStr'),
                  cell('A3', '0008', 'inlineStr') + cell('B3', '合成乙', 'inlineStr')])


def cost_sheet():
    return sheet([cell('A1', '企业编号', 'inlineStr') + cell('B1', '费用项', 'inlineStr') + cell('C1', '金额', 'inlineStr'),
                  cell('A2', '0007', 'inlineStr') + cell('B2', '场地', 'inlineStr') + cell('C2', '120.10'),
                  cell('A3', '0007', 'inlineStr') + cell('B3', '交通', 'inlineStr') + cell('C3', '20.20'),
                  cell('A4', '0008', 'inlineStr') + cell('B4', '交通', 'inlineStr') + cell('C4', '15.00')])


HEADERS = cell('A1', '企业编号', 'inlineStr') + cell('B1', '合计', 'inlineStr') + \
          cell('C1', '日期', 'inlineStr') + cell('D1', '校验', 'inlineStr')

NORMAL_APPLY = sheet([
    HEADERS,
    cell('A2', '0007', 'inlineStr') + cell('B2', '999.00', formula='SUM(甲费用)') + cell('C2', '60') + cell('D2', '0', formula='B2+1'),
    cell('A3', '0008', 'inlineStr') + cell('B3', '15', formula='费用明细!C4') + cell('C3', '61'),
])
CYCLE_APPLY = sheet([
    HEADERS,
    cell('A2', '0007', 'inlineStr') + cell('B2', '0', formula='D2') + cell('C2', '60') + cell('D2', '0', formula='B2'),
    cell('A3', '0008', 'inlineStr') + cell('B3', '15', formula='费用明细!C4'),
])
EXTERNAL_APPLY = NORMAL_APPLY.replace('SUM(甲费用)', "'[remote.xlsx]Sheet1'!A1")

# 空值四态：零（B2 录入零）、空单元格（0008 行缺 C）、空字符串（D2）、错误值（D3）。
# 合计列对 0007 故意写成无缓存公式（no_cache），核对“无缓存现场重算”路径。
BLANKS_APPLY = (
    f'<worksheet xmlns="{MAIN}"><sheetData>'
    + blank_row(1, HEADERS)
    + blank_row(2, cell('A2', '0007', 'inlineStr')
                + cell('B2', '140.30', formula='SUM(甲费用)', no_cache=True)
                + cell('C2', '61') + empty_string_cell('D2'))
    + blank_row(3, cell('A3', '0008', 'inlineStr')
                + cell('B3', '0') + error_cell('D3', '#DIV/0!'))
    + '</sheetData></worksheet>'
)

# 跨企业引用：0008 的申报合计引用了归属 0007 的费用行 C2，必须隔离 0008、保留 0007。
CROSS_APPLY = sheet([
    HEADERS,
    cell('A2', '0007', 'inlineStr') + cell('B2', '140.30', formula='SUM(甲费用)') + cell('C2', '61'),
    cell('A3', '0008', 'inlineStr') + cell('B3', '120.10', formula='费用明细!C2') + cell('C3', '61'),
])


def workbook_xml(sheet_names, epoch='0'):
    return (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<workbookPr date1904="{epoch}"/><sheets>'
            f'<sheet name="{sheet_names[0]}" sheetId="1" r:id="rOrg"/>'
            f'<sheet name="{sheet_names[1]}" sheetId="2" r:id="rCost"/>'
            f'<sheet name="{sheet_names[2]}" sheetId="3" r:id="rApply"/>'
            '</sheets><definedNames>'
            f'<definedName name="企业表">{sheet_names[0]}!$A$1:$B$3</definedName>'
            f'<definedName name="费用表">{sheet_names[1]}!$A$1:$C$4</definedName>'
            f'<definedName name="申报表">{sheet_names[2]}!$A$1:$D$3</definedName>'
            f'<definedName name="甲费用">{sheet_names[1]}!$C$2:$C$3</definedName>'
            '</definedNames></workbook>')


def renamed_apply(cost_sheet_name):
    # 公式中的表名与改名后的工作表一致，验证定位依赖关系文件/定义名称而非固定名称。
    return (
        f'<worksheet xmlns="{MAIN}"><sheetData>'
        + blank_row(1, HEADERS)
        + blank_row(2, cell('A2', '0007', 'inlineStr')
                    + cell('B2', '999.00', formula='SUM(甲费用)') + cell('C2', '60')
                    + cell('D2', '0', formula='B2+1'))
        + blank_row(3, cell('A3', '0008', 'inlineStr')
                    + cell('B3', '15', formula=f'{cost_sheet_name}!C4') + cell('C3', '61'))
        + '</sheetData></worksheet>'
    )


def relationships():
    return ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + ''.join(f'<Relationship Id="{identity}" Type="{REL}/worksheet" Target="worksheets/{name}.xml"/>'
                      for identity, name in [('rOrg', 'org'), ('rCost', 'cost'), ('rApply', 'apply')])
            + '</Relationships>')


def content_types():
    parts = [('xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
             ('xl/worksheets/org.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
             ('xl/worksheets/cost.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
             ('xl/worksheets/apply.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')]
    return ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            + ''.join(f'<Override PartName="/{part}" ContentType="{content}"/>' for part, content in parts)
            + '</Types>')


def write_book(target, name, apply_xml, epoch='0', sheet_names=('企业资料', '费用明细', '申请登记')):
    contents = {
        '[Content_Types].xml': content_types(),
        '_rels/.rels': f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rBook" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': workbook_xml(sheet_names, epoch),
        'xl/_rels/workbook.xml.rels': relationships(),
        'xl/worksheets/org.xml': org_sheet(),
        'xl/worksheets/cost.xml': cost_sheet(),
        'xl/worksheets/apply.xml': apply_xml,
    }
    with ZipFile(target / name, 'w', compression=ZIP_DEFLATED) as archive:
        for path, body in contents.items():
            info = ZipInfo(path, (2026, 11, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, body.encode())


def generate(directory='fixtures'):
    target = Path(directory)
    target.mkdir(parents=True, exist_ok=True)
    write_book(target, 'cached.xlsx', NORMAL_APPLY)
    write_book(target, 'epoch1904.xlsx', NORMAL_APPLY, epoch='1')
    write_book(target, 'cycle.xlsx', CYCLE_APPLY)
    write_book(target, 'external.xlsx', EXTERNAL_APPLY)
    write_book(target, 'blanks.xlsx', BLANKS_APPLY)
    write_book(target, 'cross.xlsx', CROSS_APPLY)
    # 工作表改名但定义名称与关系文件仍能定位，验证不依赖固定 sheet 名。
    renamed_names = ('org_2026', 'cost_2026', 'apply_2026')
    write_book(target, 'renamed.xlsx', renamed_apply(renamed_names[1]),
               sheet_names=renamed_names)


if __name__ == '__main__':
    generate()
