/*
*	07/07/2011
*	Alterado para trabalhar com o novo XML e o padrão de referências
*	Adicionado suporte para ler apenas a seleção
*/
var myDocument = app.activeDocument;

var nomeEntidadeGrande = new Array;
var titulos     = encontrar("titulo");
var entidades   = encontrar("entidade");
var consorcios  = encontrar("entidade - consorcio", nomeEntidadeGrande);
var associacoes = encontrar("entidade - associacao", nomeEntidadeGrande);
var multas = encontrar("entidade - multas");
var referencias = encontrar("codAto");
var municipios  = encontrar("municipio");
var msgErro     = localizaErroCodato();
var msgAtosImportar  = "Possivelmente há atos por importar na pág.:;";
var msgAtosFallha  = "Há com falha na importação na pág.:;";
main();

function main(){
	var alerta = '';
	var icon = false;
	
	if(titulos !=referencias){
		alerta = 'ALERTA!!!\r\rNúmero dos Titulos DIFERENTE do Numero dos CodAtos\r-------------------------------------------------------\r\r';
		icon = true;
	}

    var time = startTime = new Date().getHours();
    if(time >= 10){
        localizaFalhaAoImportar();
        localizaAtoporImportar();
    }
    if (nomeEntidadeGrande.length > 0) {
        msgErro += '\r\r-------------------------------------------------------\r\rAtenção! Foram encontradas entidades com nome muito extenso:\r\r' + nomeEntidadeGrande.reverse().join('\r');
    }

   //if(msgAtosImportar.length > 44)    {
      //  msgErro += msgAtosFallha + ':(';
   //}
    if (associacoes > 0 || consorcios > 0 || multas > 0 ){
       $.writeln('P1: ' + municipios + ', '+entidades +', '+referencias+  ', A: ' +associacoes+';\   C: '+consorcios+', \   M :' +multas+";");
    }
    else{  
         $.writeln('P1: '+ municipios + ', '+entidades +', '+referencias+";");
         $.writeln('P1:	'+ municipios + '	'+entidades +'	'+referencias);
      }
	alert(alerta + 'Municipios :\t' + municipios + '\t\tTitulos :\t\t'+titulos + '\rEntidades :\t'+entidades+ ' \t\tCodAto :\t\t'+referencias+ '\r\rAssociações :\t' +associacoes+'\t\rConsórcios :\t'+consorcios+' \r\rEntidades Multas :\t' +multas+ ' \r\r--------------------------------------------------\r'  + msgErro   , 'Conferencia de titulos x atos', icon);
      	  Clipboard.SetText('P1:	'+ municipios + '	'+entidades +'	'+referencias);

}

function encontrar(nome, nomesGrandes){
	//Search the document for the string "Text".
	app.findTextPreferences.findWhat = "";
	app.changeTextPreferences.changeTo = "";
	var contador = 0;
	//Set the find options.
	app.findChangeTextOptions.caseSensitive = false;
	app.findChangeTextOptions.includeFootnotes = false;
	app.findChangeTextOptions.includeHiddenLayers = false;
	app.findChangeTextOptions.includeLockedLayersForFind = true;
	app.findChangeTextOptions.includeLockedStoriesForFind = true;
	app.findChangeTextOptions.includeMasterPages = false;
	app.findChangeTextOptions.wholeWord = false;
	var stilo = myDocument.paragraphStyles.item(nome);
    if(!stilo.isValid){
            return 0;
    }

	app.findTextPreferences.appliedParagraphStyle = stilo;
	var items;
	
	if(app.selection.length != 0){
		items = app.selection[0].findText();
	}else{
		items = myDocument.findText();
	}
	//alert("Found " + myFoundItems[0].texts[0].contents + " instances of the search string.");
	contador = contador + items.length;

    if (nomesGrandes) {
        for(var i=0; i < items.length; i++) {
                if (items[i].length > 20) {
                        nomesGrandes.unshift(''+items[i].contents);
                }
        }
    }
    
	return contador;
}
function localizaErroCodato(){

var doc = app.activeDocument;

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

var contador = new Array;

for(var j=0;j<myFoundItems.length;j++){
	try {
		var pagina;
		if ('parentPage' in myFoundItems[j].parentTextFrames[0]) {
			pagina = myFoundItems[j].parentTextFrames[0].parentPage.name;
		} else {
			pagina = myFoundItems[j].parentTextFrames[0].parent.name;
		}
		pagina = String(pagina).replace(/[^\d]/g,'');
		
		if (contador[pagina] == undefined) {
			contador[pagina] = 0;
		}

		contador[pagina]++;
	} catch (e) {}
}


app.findTextPreferences.findWhat = "";
app.changeTextPreferences.changeTo = "";
app.findChangeTextOptions.caseSensitive = false;
app.findChangeTextOptions.includeFootnotes = false;
app.findChangeTextOptions.includeHiddenLayers = false;
app.findChangeTextOptions.includeLockedLayersForFind = false;
app.findChangeTextOptions.includeLockedStoriesForFind = false;
app.findChangeTextOptions.includeMasterPages = false;
app.findChangeTextOptions.wholeWord = false;
app.findTextPreferences.appliedParagraphStyle = "titulo";
var myFoundItems = doc.findText();

for(var j=0;j<myFoundItems.length;j++){
	try {
		var pagina;
		if ('parentPage' in myFoundItems[j].parentTextFrames[0]) {
			pagina = myFoundItems[j].parentTextFrames[0].parentPage.name;
		} else {
			pagina = myFoundItems[j].parentTextFrames[0].parent.name;
		}
		pagina = String(pagina).replace(/[^\d]/g,'');

		if (contador[pagina] == undefined) {
			contador[pagina] = 0;
		}

		contador[pagina]--;
	} catch (e) {}
}

var msgFct ='';

for(page=0; page < contador.length; page++) {
		if (contador[page] !== undefined && contador[page] !== 0) {
			msgFct += page + ' ,   ';
		}
}
if (msgFct != '') {
	return "Falha na(s) página(s):  "+msgFct;
} else {
	return "Nenhum problema de Titulo x codAto encontrado!";
}

} 


function localizaFalhaAoImportar(){

var doc = app.activeDocument;

    app.findTextPreferences.findWhat = "##### Falha na importação ###";
    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = true;
    app.findChangeTextOptions.wholeWord = true;
    app.findTextPreferences.appliedParagraphStyle = "texto";
    var myFoundItems = doc.findText();

    var contador = new Array;

    for(var j=0;j<myFoundItems.length;j++){
        try {
            var pagina;
            if ('parentPage' in  myFoundItems[j].parentTextFrames[0]) {
                pagina = myFoundItems[j].parentTextFrames[0].parentPage.name;
            } else {
                pagina = myFoundItems[j].parentTextFrames[0].parent.name;
            }
            if (contador[pagina] == undefined) {
                contador[pagina] = 0;
            }

            contador[pagina]++;
        } catch (e) {}
        $.writeln(pagina);
     }
}


function localizaAtoporImportar(){

var doc = app.activeDocument;
    app.findTextPreferences.findWhat = "==============^pImportar";
    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.includeLockedLayersForFind = false;
    app.findChangeTextOptions.includeLockedStoriesForFind = false;
    app.findChangeTextOptions.includeMasterPages = true;
    app.findChangeTextOptions.wholeWord = true;
    app.findTextPreferences.appliedParagraphStyle = "texto";
    var myFoundItems = doc.findText();

    var contador = new Array;

    for(var j=0;j<myFoundItems.length;j++){
        try {
            var pagina;
            if ('parentPage' in myFoundItems[j].parentTextFrames[0]) {
                pagina = myFoundItems[j].parentTextFrames[0].parentPage.name;
            } else {
                pagina = myFoundItems[j].parentTextFrames[0].parent.name;
            }
            if (contador[pagina] == undefined) {
                contador[pagina] = 0;
            }

            contador[pagina]++;
        } catch (e) {}
        
        $.writeln(pagina);
     }
  pagina = 0;
}
