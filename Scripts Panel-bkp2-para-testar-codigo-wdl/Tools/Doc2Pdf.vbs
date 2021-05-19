Set myApp = CreateObject("InDesign.Application")

Set objFSO = CreateObject( "Scripting.FileSystemObject" )
Set objWord = CreateObject( "Word.Application" )

objWord.Visible = False
 objWord.Documents.Open arguments(0)
 objWord.ActiveDocument.ExportAsFixedFormat arguments(1), 17, False
 'objWord.ActiveDocument.SaveAs arguments(1), 17, False
 objWord.ActiveDocument.Close
objWord.Quit
