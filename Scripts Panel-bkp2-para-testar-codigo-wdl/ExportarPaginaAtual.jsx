

var myDoc = app.activeDocument;
var myFolderName = myDoc.filePath;
var myDocumentName = myDoc.name;
var myPages = myDoc.pages;
var myPageName = myPages[0].name;
var myFilePath = "/Users/ciga1/Desktop/" + myDoc.name.slice (0, -5) + "page " + myPageName + ".pdf";
var myFile = new File(myFilePath);
app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
//Do not open the PDF Export dialog box. Set “false” to “true” if you want the dialog box.
myDoc.exportFile(ExportFormat.pdfType, myFile, false, "Smallest File Size");
