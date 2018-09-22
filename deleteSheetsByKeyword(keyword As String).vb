''=======================================================
'' Program:   deleteSheetsByKeyword
'' Desc:      Deletes all worksheets who's name matches a user-specified keyword
'' Arguments: keyword -- The keyword to be used as a search string
'' Comments:
'' Changes----------------------------------------------
'' Date       Programmer     Change
'' <Date>     <Name>         Written
''=======================================================
Public Sub deleteSheetsByKeyword(keyword As String)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.name Like ("*" & keyword & "*") Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub
