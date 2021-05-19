var doc = app.activeDocument;

var arquivo = new File(String(doc.fullName).replace('.indd','.txt'));
arquivo.open("w", "ANSI");

app.findTextPreferences.findWhat = "";
app.changeTextPreferences.changeTo = "";
app.findChangeTextOptions.caseSensitive = false;
app.findChangeTextOptions.includeFootnotes = false;
app.findChangeTextOptions.includeHiddenLayers = false;
app.findChangeTextOptions.includeLockedLayersForFind = false;
app.findChangeTextOptions.includeLockedStoriesForFind = false;
app.findChangeTextOptions.includeMasterPages = false;
app.findChangeTextOptions.wholeWord = false;
app.findTextPreferences.appliedParagraphStyle = "codAto";
var myFoundItems = doc.findText();

for(var j=0;j<myFoundItems.length;j++){
    var codAto = String(myFoundItems[j].contents).replace(/[^\d]/g,'');
    var pagina = String(myFoundItems[j].parentTextFrames[0].parentPage.name).replace(/[^\d]/g,'');
    arquivo.writeln(codAto+";"+pagina);
}


arquivo.close();