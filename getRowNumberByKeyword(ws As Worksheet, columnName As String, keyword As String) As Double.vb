''=======================================================
'' Program:   getRowNumberByKeyword
'' Desc:      Searches down a user-specified column and returns the row of the first item
''            that matches a keyword
'' Arguments: ws         -- The worksheet the macro should target
''            columnName -- The name of the column to search through
''	          keyword    -- The keyword to be used as a search string
'' Comments:
'' Changes----------------------------------------------
'' Date       Programmer     Change
'' 5/29/2018  Quinn McHugh   Written
''=======================================================
Public Function getRowNumberByKeyword(ws As Worksheet, columnName As String, keyword As String) As Double
        Set column = ws.Cells(1, getColumnNumber(ws, columnName)).EntireColumn
        getRowNumberByKeyword = WorksheetFunction.Match(keyword, column, 0)
End Function