import zipfile
import xmltodict
from itertools import chain

def getTables(filepath):

    archive = zipfile.ZipFile(filepath)

    namelist = archive.namelist()


    def getDict(xml_file):
        parsed_dict = xmltodict.parse(archive.read(xml_file))
        return(parsed_dict)


    def getSheetName(sheet_no):
        for sheet in sheets:
            if sheet_no in  sheet.values():
                return(sheet['@name'])


    def getTableInfo(table_xml_name):
        table = getDict(table_xml_name)
        table_range = table['table']['@ref']
        table_name = table['table']['@displayName']
        return(table_range, table_name)


    ##get workbook and sheets
    workbook = getDict('xl/workbook.xml')
    sheets = [od for od in workbook['workbook']['sheets']['sheet']]

    ##parse a sheetname that corresponds with other references
    for od in sheets:
        od['sheetname'] = 'sheet' + od['@sheetId']

    ##find which tables are in which sheets
    rels = {s.split('/')[-1].replace('.xml.rels', ''): getDict(s) for s in namelist if 'xl/worksheets/_rels' in s}

    rels = {k: v['Relationships']['Relationship'] for k,v in rels.items()}

    rels = {k: [v] if type(v) != list else v for k, v in rels.items()}


    for sheet, rellist in rels.items():
        for rel in rellist:
            search = sheet
            name = getSheetName(search)
            rel['sheet_ui_name'] = name

            table_xml = rel['@Target'].replace('..', 'xl')
            
            table_info = getTableInfo(table_xml)

            rel['table_range'] = table_info[0]
            rel['table_ui_name'] = table_info[1]

    table_data = [{'sheet_ui_name': d['sheet_ui_name'],
                   'table_ui_name': d['table_ui_name'],
                   'table_range': d['table_range']} for d in chain(*rels.values())]

    return(table_data)
    
