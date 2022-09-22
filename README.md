# ASCII Chr() Codes Microsoft Access VBA.  
Find ASCII chr() characters in any dataset.  
Table ChrCodes contains 255 ASCII character codes with status of "1" or "0".  
The code creates new fields in the "selectedTable" called [YourFiledName]_Error looping through both ChrCodes and selectedTable.  
Special character names with status "0" will display in the new field [YourFiledName]_Error.  
