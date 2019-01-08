''=======================================================
'' Program:   isConnectedToInternet
'' Desc:      Checks if the computer is connected to the internet
'' Arguments: 
'' Comments:
'' Changes----------------------------------------------
'' Date       Programmer     Change
'' <Date>     <Name>         Written
''=======================================================
Public Declare Function InternetGetConnectedState _
                         Lib "wininet.dll" (lpdwFlags As Long, _
                                            ByVal dwReserved As Long) As Boolean

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" ( _
ByRef lpdwFlags As Long, _
ByVal lpszConnectionName As String, _
ByVal dwNameLen As Long, _
ByVal dwReserved As Long) As Long

' Checks if the computer is connected to the Internet
Function isConnectedToInternet() As Boolean
    isConnectedToInternet = CBool(InternetGetConnectedStateEx(0, vbNullString, 512, 0&))
End Function
