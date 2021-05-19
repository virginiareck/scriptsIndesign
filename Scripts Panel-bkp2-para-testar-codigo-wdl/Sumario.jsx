var doc = app.activeDocument;

function criarSumario(pagina) {   
    app.scriptPreferences.enableRedraw = false;   
    
    var style = doc.tocStyles[0];
    doc.select(pagina);
    
    var story = doc.createTOC(style, true, undefined, [100, 200])[0];
    
    var capa = doc.masterSpreads.itemByName('B-Capa');
    
    for (var f=0; f < capa.textFrames.length; f++) {
        
        var frameCapa = capa.textFrames[f];
   
        if (frameCapa.contents.length == 0) { // localiza o textframe vazio na master como sendo o lugar do sumário
            pagina.textFrames.lastItem().geometricBounds = frameCapa.geometricBounds;
            break;
        }
    }

	app.scriptPreferences.enableRedraw = true; 
    return;
}


function main(){
        criarSumario(doc.pages[0]);
}

app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.fastEntireScript, "Gerando sumário ..."); 
 