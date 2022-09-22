# ChrCodes Microsoft Access VBA.  
Find chr() characters in any dataset.  
The application contains a table, ChrCodes, with 255 specific Chr() codes with a "1" or "0" status.  
The code creates new fields in the "selectedTable" called [YourFiledName]_Error looping through both ChrCodes and selectedTable.  
Special character names with status "0" will display in the new [YourFiledName]_Error field.  
