''=======================================================
'' Program:   colNum2Letter
'' Desc:      Takes in a column number and returns its equivalent letter.
'' Arguments: columnNumber -- The column number to be converted.
'' Comments:
'' Changes----------------------------------------------
'' Date        Programmer     Change
'' <Date>      <Name>         Written
''=======================================================
Public Function colNum2Letter(colNumber As Double) As String
    Dim strArray() As String
    strArray = Split(Cells(1, colNumber).Address(True, False), "$")
    colNum2Letter = strArray(0)
End Function
