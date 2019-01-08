''=======================================================
'' Program:   DynamicVLookup
'' Desc:      Looks up a user-specified value in a user-specified column in a user-specified
''            worksheet and returns another value that is in the same row but is in a different
''            user-specified column.
'' Calls:     getColumnNumber, colNum2Letter
'' Arguments: wsName        -- The name of the worksheet where the lookup should be performed.
''            lookupValue   -- The value to be looked up.
''            lookupColName -- The name of the column that contains the lookup value.
''            returnColName -- the name of the column that contains the desired return value.
'' Comments:  This function should only be used in spreadsheets where all column titles are on
''            the first row and all data is listed underneath. Like so:
''            First Name | Last Name | Favorite Color
''            John       | Williams  | Blue
''            Jane       | Smith     | Orange
'' Changes----------------------------------------------
'' Date        Programmer     Change
'' <Date>      <Name>         Written
''=======================================================
Public Function DynamicVLookup(wsName As String, lookupValue As String, _
    lookupColName As String, returnColName As String) As String
    
    Dim ws As Worksheet
    Set ws = Sheets(wsName)
    
    Dim lookupColNum As Integer
    lookupColNum = getColumnNumber(ws, lookupColName)
    
    Dim lookupColLetter As String
    lookupColLetter = colNum2Letter(CDbl(lookupColNum))
    
    Dim rowNum As Integer
    lookupCol = ws.Range(lookupColLetter & ":" & lookupColLetter)
    rowNum = WorksheetFunction.Match(lookupValue, lookupCol, 0)
    
    Dim returnColNum As Integer
    returnColNum = getColumnNumber(ws, returnColName)
    
    DynamicVLookup = ws.Cells(rowNum, returnColNum).Value
End Function
