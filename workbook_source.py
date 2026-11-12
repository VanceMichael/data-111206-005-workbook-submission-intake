from pathlib import PurePosixPath
from zipfile import ZipFile
from xml.etree import ElementTree

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def inventory(path):
    with ZipFile(path) as archive:
        entries = archive.infolist()
        if len({entry.filename for entry in entries}) != len(entries):
            raise ValueError('容器出现重名成员')
        if sum(entry.file_size for entry in entries) > 8 * 1024 * 1024:
            raise ValueError('工作簿展开尺寸超限')
        for entry in entries:
            name = PurePosixPath(entry.filename)
            if name.is_absolute() or '..' in name.parts or '\\' in entry.filename:
                raise ValueError('容器路径不安全')
        workbook = ElementTree.fromstring(archive.read('xl/workbook.xml'))
        relationships = ElementTree.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        links = {entry.attrib['Id']: entry.attrib for entry in relationships}
        sheets = []
        for sheet in workbook.findall(f'.//{{{MAIN}}}sheet'):
            relation = links[sheet.attrib[f'{{{REL}}}id']]
            sheets.append({'name': sheet.attrib['name'], 'relationship': relation['Id'],
                           'target': relation['Target'], 'external': relation.get('TargetMode') == 'External'})
        return {'sheets': sheets, 'members': len(entries),
                'date1904': workbook.find(f'{{{MAIN}}}workbookPr').get('date1904', '0') == '1'}
