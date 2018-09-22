''=======================================================
'' Program:   findSheetByName
'' Desc:      Searches through all sheets and returns the first sheet who's name exactly matches
''            the search string
'' Arguments: wsName -- The name of the worksheet to be used as a search string
'' Comments:
'' Changes----------------------------------------------
'' Date       Programmer     Change
'' <Date>     <Name>         Written
''=======================================================
Public Function findSheetByName(wsName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name = wsName Then
            Set findSheetByName = ws
            Exit Function
        End If
    Next
End Function
