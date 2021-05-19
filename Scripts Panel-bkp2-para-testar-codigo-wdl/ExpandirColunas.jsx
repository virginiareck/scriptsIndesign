
function transporColunas(ato) {
    var w = 195;
    with (ato.xmlContent) {
        //if (spanColumnType != SpanColumnTypeOptions.SPAN_COLUMNS) {
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
        //} else {
         //   spanColumnType = SpanColumnTypeOptions.SINGLE_COLUMN;
         //  w = 95;
        //}
    }
    for(var t = 0; t < ato.tables.length; t++) {
        growColumns(ato.tables[t], w / ato.tables[t].width);
    }
}

function growColumns(table, factor) {
    for(var c=0;c<table.columns.length;c++) {
        var w = Math.floor(table.columns[c].width * factor);
        table.columns[c].width = w > 1.5 ? w : "1.5mm";
    }    
}


function selecionarAtos() {   
    try {
         if (app.selection.length != 0) {
            app.findTextPreferences.appliedParagraphStyle = app.activeDocument.paragraphStyles.item("codAto");
            app.findTextPreferences.findWhat = '';
            var atos = app.selection[0].findText();
            for (var i=0; i<atos.length;i++ ) {
                atos[i] = atos[i].associatedXMLElements[0];
            }
            return atos;
        } 
    } catch (e) {}
    
    return new Array();
    
}

function main() {
    var atos = selecionarAtos();
    for(var a=0; a<atos.length; a++) {
        transporColunas(atos[a]);
    }
}

main();