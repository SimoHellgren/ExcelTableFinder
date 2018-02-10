# ExcelTableFinder
A function for extracting information on all tables in an Excel file. Main use case is to get the sheet and cell range of each table, so they can be utilized in openpyxl/pandas. More precicely, this function returns a list of dictionaries, each corresponding to one table in the specified Excel file, containing the table's name, cell range and the name of the worksheet the table resides in.

All credit to Martin Blech & other contributors for xmltodict
