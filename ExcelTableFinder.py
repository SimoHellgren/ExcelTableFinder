import zipfile
import xmltodict
from itertools import chain



def findTables(filepath):

    archive = zipfile.ZipFile(filepath)

    namelist = archive.namelist()


    def getDict(s):
        res = xmltodict.parse(archive.read(s))
        return(res)


    def getTableInfo(table_xml_filename):
        soup = xmltodict.parse(archive.read(table_xml_filename))
        reference = soup['table']['@ref']
        uiname = soup['table']['@displayName']
        return(reference, uiname)


    def getSheetDataByID(sheetID):
        for sheet in sheets:
            if sheetID in sheet.values():
                sheet_name = sheet['@name']
                sheet_no = sheet['@sheetId']

                return(sheet_name, sheet_no)

    ##get sheet names, ID:s and relationshipID:s
    workbook = getDict('xl/workbook.xml')
    sheets = [od for od in workbook['workbook']['sheets']['sheet']]

    ##find all relationships in document (connect relationshipID:s with table names)
    rels = [getDict(s) for s in namelist if 'xl/worksheets/_rels/' in s]
    rels = [r['Relationships']['Relationship'] for r in rels]

    rels = [[r] if type(r) != list else r for r in rels]
    rels = list(chain(*rels))

    ##get tables and their respective sheet ID:s
    tables = [{'table_path': s['@Target'],
               'sheet_ID': s['@Id']} for s in rels]

    ##create names and get tables':
    ##1) cell references
    ##2) sheetnames
    ##3) sheetnumbers

    for d in tables:
        d['name'] = d['table_path'].split('/')[-1][:-4]

        table_info = getTableInfo(d['table_path'].replace('..', 'xl'))

        d['cells'] = table_info[0]
        d['ui_name'] = table_info[1]

        sheet_data = getSheetDataByID(d['sheet_ID'])

        d['sheet_name'] = sheet_data[0]
        d['sheet_number'] = sheet_data[1]


    return(tables)
