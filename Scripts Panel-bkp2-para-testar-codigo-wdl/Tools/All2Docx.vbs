Set myApp = CreateObject("InDesign.Application")

Set objFSO = CreateObject( "Scripting.FileSystemObject" )
Set objWord = CreateObject( "Word.Application" )

objWord.Visible = False
objWord.DisplayAlerts = 0
objWord.FeatureInstall = 0

objWord.Documents.Open arguments(0)
objWord.ActiveDocument.ConvertNumbersToText
objWord.ActiveDocument.SaveAs arguments(1), 12
objWord.ActiveDocument.Close

objWord.Quit



