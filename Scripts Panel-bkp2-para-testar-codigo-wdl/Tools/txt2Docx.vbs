Set myApp = CreateObject("InDesign.Application")

Set objFSO = CreateObject( "Scripting.FileSystemObject" )
Set WordApp = CreateObject("Word.Application")
WordApp.Visible = FALSE

Set WordDoc = WordApp.Documents.Open(WScript.Arguments.Item(0),false)



WordDoc.SaveAs arguments(1), 12
WordDoc.Close()
WScript.Quit

