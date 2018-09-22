''=======================================================
'' Program:   findNextRowInColumn
'' Desc:      Returns the next empty row in a user-specified
''            column on a user-specified worksheet
'' Arguments: ws -- The worksheet to be searched
''            columnLetter -- The letter of the column to be searched
'' Comments: 
'' Changes----------------------------------------------
'' Date        Programmer     Change
'' 7/12/2018   Quinn McHugh   Written
''=======================================================
Public Function findLastRowInColumn(ws As Worksheet, columnLetter As String) As Integer
    findLastRowInColumn = ws.Range(columnLetter & ws.Rows.Count).End(xlUp).Offset(1, 0).row
End Function