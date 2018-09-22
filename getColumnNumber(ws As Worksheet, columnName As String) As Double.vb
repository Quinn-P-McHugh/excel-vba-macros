''=======================================================
'' Program:   getColumnNumber
'' Desc:      Searches the first row of a worksheet and returns the column
''            number of the column who's title matches a search string
'' Arguments: ws         -- The worksheet to be searched
''            columnName -- The name of the column who's column number will be returned
'' Comments:  This subroutine only searches the first row of the spreadsheet, so column titles must be in the first row
'' Changes----------------------------------------------
'' Date        Programmer     Change
'' <Date>      <Name>         Written
''=======================================================
Public Function getColumnNumber(ws As Worksheet, columnName As String) As Double
    getColumnNumber = WorksheetFunction.Match(columnName, ws.Rows(1).EntireRow, 0)
End Function
