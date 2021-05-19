var doc = app.activeDocument;
var rootElement = doc.xmlElements.item(0);

function abrirDoc() {
    
    if (app.selection.length != 0) {
        app.findTextPreferences.appliedParagraphStyle = doc.paragraphStyles.item("codAto");
        app.findTextPreferences.findWhat = '';
        var codAtos;
        try {
            codAtos = app.selection[0].findText();
        } catch (e) {
            return false;
        }
        if (codAtos.length > 0) {
            var codato = codAtos[0].contents;
            var folder = new Folder(doc.fullName.path + '/arquivos/');
            var files = folder.getFiles(codato.replace(/[^\d]/g,'') + '.*');
            for(var i=0; i < files.length; i++) {
                var file = files[i];
                if (Window.confirm('Abrir '+file+' ?')) {
                    file.execute();
                }
            }
        }
    }

}

abrirDoc();