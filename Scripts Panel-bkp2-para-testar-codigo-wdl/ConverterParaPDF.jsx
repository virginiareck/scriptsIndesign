var doc = app.activeDocument;
var arquivos = doc.fullName.path + '/arquivos/';
var folder = new Folder(arquivos);

function selecionarAtos() {   
    if (app.selection.length != 0) {
        var p = app.findTextPreferences.appliedParagraphStyle;
        var w = app.findTextPreferences.findWhat;
        app.findTextPreferences.appliedParagraphStyle = doc.paragraphStyles.item("codAto");
        app.findTextPreferences.findWhat = '';
        var atos = app.selection[0].findText();
        for (var i=0; i<atos.length;i++ ) {
            atos[i] = atos[i].associatedXMLElements[0];
        }
        app.findTextPreferences.appliedParagraphStyle = p;
        app.findTextPreferences.findWhat = w;
        return atos;
    } 

    Window.alert('Nenhum ato selecionado.');
    return Array();

}
var atos = selecionarAtos();
for(var ato=0; ato < atos.length; ato++) {
    var c = atos[ato].evaluateXPathExpression("//codAto")[0].xmlContent.contents;
    var pdf = new File(arquivos + 'refeito_' + c + '.pdf');
    if (!pdf.exists) {
        pdf = new File(arquivos + c + '.pdf');
    }
    var file = folder.getFiles(c + ".docx");
    if (file.length == 0) {
        file = folder.getFiles(c + ".*");
    }
    file = file[0];

    $.writeln(''+file.fsName);
    $.writeln(''+pdf.fsName);
    var doc2pdf =  new File((new File($.fileName)).parent + "/Tools/Doc2Pdf.vbs"); 
    app.doScript (doc2pdf, ScriptLanguage.VISUAL_BASIC, [ file.fsName,  pdf.fsName ] );  
    
    //if (Window.confirm('Abrir o pdf gerado? '+pdf.name)) {  //2017-11-30
    //    pdf.execute();
    //}
    if (!pdf.exists) {
        Window.alert('PDF não pode ser gerado.');
        }
    else{
        $.writeln(''+pdf.name + "  criado com sucesso.");
    }

}

////MUDANÇAS
//2017-11-30 COLOCADO COMO COMENTARIO pdf.execute();
// 2018-11-09 $.writeln(''+pdf.fsName); logo apos app.doScript para conseguirmos ver pelo console se app executou
//2018-11-09   if (!pdf.exists) { Window.alert else  $.writeln   ###   se deu errado, um msgbox de alerta, se deu certo, apenas um writeln no console