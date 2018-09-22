''=======================================================
'' Program:   findSheetByKeyword
'' Desc:      Searches through all sheets and returns the first sheet who's name matches
''            a keyword
'' Arguments: keyword -- The keyword to be used as a search string
'' Comments:
'' Changes----------------------------------------------
'' Date       Programmer     Change
'' 5/30/2018  Quinn McHugh   Written
''=======================================================
Public Function findSheetByKeyword(keyword As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name Like "*" & keyword & "*" Then
            Set findSheetByName = ws
            Exit Function
        End If
    Next
End Function