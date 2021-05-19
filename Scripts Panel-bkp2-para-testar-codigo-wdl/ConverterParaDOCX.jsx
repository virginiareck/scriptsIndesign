var doc = app.activeDocument;
var arquivos = doc.fullName.path + '/arquivos/';
var folder = new Folder(arquivos);

function selecionarAtos() {   
    try {
         if (app.selection.length != 0) {
            app.findTextPreferences.appliedParagraphStyle = myDoc.paragraphStyles.item("codAto");
            app.findTextPreferences.findWhat = '';
            var atos = app.selection[0].findText();
            for (var i=0; i<atos.length;i++ ) {
                atos[i] = atos[i].associatedXMLElements[0];
            }
            return atos;
        } 
        else if (Window.confirm('Deseja converter TODOS os atos agora?')) 
        {
            return myDoc.xmlElements.item(0).evaluateXPathExpression("//ato").reverse();
            
        }
    } catch (e) {}
    
    return new Array();
    
}
var atos = selecionarAtos();
for(var ato=0; ato < atos.length; ato++) {
    converterArquivo(ato);
}


function converterArquivo(ato) {

    var codigo = ato.evaluateXPathExpression("//codAto")[0].xmlContent.contents;
    try {
        var file = folder.getFiles(codigo + ".docx");
        if (file.length == 0) {
            file = folder.getFiles(codigo + ".*");
        }
        file = file[0] instanceof File ? file[0] : file[1];
        
      
        var doc = file.fsName.replace(/\.[a-z]+$/,'.docx');
               
        if (file.fileExt() == 'odt' && converterODT) {
            
            var args = '--norestore --invisible --nologo --nodefault --headless --convert-to docx --outdir "' + file.parent.fsName + '" "' + file.fsName + '"';
            var cmd = '"' + soffice + " " + args + '"';
            app.doScript(runvbs, ScriptLanguage.VISUAL_BASIC, [ cmd ]); 
            
            $.writeln(''+file.fsName); 
            return folder.getFiles(codigo + ".docx")[0];
        } else if (file.fileExt() == 'doc' && converterDOC) {
                        
            app.doScript(all2docx, ScriptLanguage.VISUAL_BASIC, [ file.fsName, doc ] ); 
            
            $.writeln(''+file.fsName); 
            return folder.getFiles(codigo + ".docx")[0];
        }else {
            return file;
        }
    } catch (e) {
        Window.alert('Não foi possível encontrar ou processar o arquivo do ato '+codigo);    
    }
}

////MUDANÇAS
//2017-11-30 COLOCADO COMO COMENTARIO pdf.execute();
// 2018-11-09 $.writeln(''+pdf.fsName); logo apos app.doScript para conseguirmos ver pelo console se app executou
//2018-11-09   if (!pdf.exists) { Window.alert else  $.writeln   ###   se deu errado, um msgbox de alerta, se deu certo, apenas um writeln no console