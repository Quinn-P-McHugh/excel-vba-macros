''=======================================================
'' Program:   exportAsPostScript
'' Desc:      Exports a user-specified worksheet as a PostScript file
'' Arguments: ws       -- The worksheet to be exported
''            fileName -- The name to be given to the exported file (i.e. "Exported_Spreadsheet")
''            filePath -- The path that the file should be exported to (i.e. "C:\Documents and Settings\User\Desktop")
'' Comments:
'' Changes----------------------------------------------
'' Date       Programmer     Change
'' <Date>     <Name>         Written
''=======================================================
Public Sub exportAsPostScript(ws As Worksheet, fileName As String, filePath As String)
	' NOTE: PrintOut command w/ "Adobe PDF" as ActivePrinter outputs a PostScript (.ps) file, NOT a PDF

	Dim completeFilePath As String
	completeFilePath = filePath & "\" & fileName & ".ps"

	' ActivePrinter:="Adobe PDF on Ne02:", <--- Alternative name for Adobe PDF printer, depending on your settings
	ws.PrintOut ActivePrinter:="Adobe PDF", PrintToFile:=True, _
		PrToFileName:=completeFilePath
	
	MsgBox "Spreadsheet was successfully exported." & vbNewLine & vbNewLine & "File Path - " & filePath, vbOKOnly, "Success!"
End Sub
