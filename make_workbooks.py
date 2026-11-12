from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def cell(reference, value, kind=None, formula=None):
    attribute = f' t="{kind}"' if kind else ''
    expression = f'<f>{formula}</f>' if formula else ''
    stored = f'<is><t>{value}</t></is>' if kind == 'inlineStr' else f'<v>{value}</v>'
    return f'<c r="{reference}"{attribute}>{expression}{stored}</c>'


def sheet(rows):
    return f'<worksheet xmlns="{MAIN}"><sheetData>' + ''.join(f'<row r="{number}">{body}</row>' for number, body in enumerate(rows, 1)) + '</sheetData></worksheet>'


def generate(directory='fixtures'):
    target = Path(directory)
    target.mkdir(parents=True, exist_ok=True)
    workbook = f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><workbookPr date1904="EPOCH"/><sheets><sheet name="企业资料" sheetId="1" r:id="rOrg"/><sheet name="费用明细" sheetId="2" r:id="rCost"/><sheet name="申请登记" sheetId="3" r:id="rApply"/></sheets><definedNames><definedName name="企业表">企业资料!$A$1:$B$3</definedName><definedName name="费用表">费用明细!$A$1:$C$4</definedName><definedName name="申报表">申请登记!$A$1:$D$3</definedName><definedName name="甲费用">费用明细!$C$2:$C$3</definedName></definedNames></workbook>'
    relationships = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(f'<Relationship Id="{identity}" Type="{REL}/worksheet" Target="worksheets/{name}.xml"/>' for identity, name in [('rOrg','org'),('rCost','cost'),('rApply','apply')]) + '</Relationships>'
    org = sheet([cell('A1','企业编号','inlineStr')+cell('B1','企业名称','inlineStr'),cell('A2','0007','inlineStr')+cell('B2','合成甲','inlineStr'),cell('A3','0008','inlineStr')+cell('B3','合成乙','inlineStr')])
    cost = sheet([cell('A1','企业编号','inlineStr')+cell('B1','费用项','inlineStr')+cell('C1','金额','inlineStr'),cell('A2','0007','inlineStr')+cell('B2','场地','inlineStr')+cell('C2','120.10'),cell('A3','0007','inlineStr')+cell('B3','交通','inlineStr')+cell('C3','20.20'),cell('A4','0008','inlineStr')+cell('B4','交通','inlineStr')+cell('C4','15.00')])
    headers = cell('A1','企业编号','inlineStr')+cell('B1','合计','inlineStr')+cell('C1','日期','inlineStr')+cell('D1','校验','inlineStr')
    normal = sheet([headers,cell('A2','0007','inlineStr')+cell('B2','999.00',formula='SUM(甲费用)')+cell('C2','60')+cell('D2','0',formula='B2+1'),cell('A3','0008','inlineStr')+cell('B3','15',formula='费用明细!C4')+cell('C3','61')])
    cycle = sheet([headers,cell('A2','0007','inlineStr')+cell('B2','0',formula='D2')+cell('C2','60')+cell('D2','0',formula='B2'),cell('A3','0008','inlineStr')+cell('B3','15',formula='费用明细!C4')])
    external = normal.replace('SUM(甲费用)', "'[remote.xlsx]Sheet1'!A1")
    for name, epoch, application in [('cached.xlsx','0',normal),('epoch1904.xlsx','1',normal),('cycle.xlsx','0',cycle),('external.xlsx','0',external)]:
        contents = {'[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+''.join(f'<Override PartName="/{part}" ContentType="{content}"/>' for part,content in [('xl/workbook.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),('xl/worksheets/org.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),('xl/worksheets/cost.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),('xl/worksheets/apply.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')])+'</Types>', '_rels/.rels':f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rBook" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>', 'xl/workbook.xml':workbook.replace('EPOCH',epoch), 'xl/_rels/workbook.xml.rels':relationships, 'xl/worksheets/org.xml':org, 'xl/worksheets/cost.xml':cost, 'xl/worksheets/apply.xml':application}
        with ZipFile(target/name,'w',compression=ZIP_DEFLATED) as archive:
            for path, body in contents.items():
                info=ZipInfo(path, (2026,11,1,0,0,0))
                info.compress_type=ZIP_DEFLATED
                archive.writestr(info,body.encode())


if __name__ == '__main__':
    generate()
