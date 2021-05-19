/*
*	07/07/2011
*	Alterado para trabalhar com o novo XML e o padrão de referências
*	Adicionado suporte para ler apenas a seleção
*/
var myDocument = app.activeDocument;

var titulos = encontrar("titulo");
var entidades = encontrar("entidade");
entidades += encontrar("entidade - consorcio");
entidades += encontrar("entidade - associacao");
var consorcios = encontrar("entidade - consorcio");
var associacoes = encontrar("entidade - associacao");
var referencias = encontrar("codAto");
var municipios = encontrar("municipio");
var msgErro = localizaErroCodato();

main();

function main(){
	var alerta = '';
	var icon = false;
	
	if(titulos !=referencias){
		alerta = 'ALERTA!!!\r\rNúmero dos Titulos DIFERENTE do Numero dos CodAtos\r--------------------------------------------------\r\r';
		icon = true;
	}

	//alert(alerta + 'Municipios:'+municipios+' \tEntidades : '+entidades+' \rTitulos: '+titulos + ' \t\tCodAto: '+referencias , 'Titulo', icon); 07/05/2014
	//alert(alerta + 'Municipios :\t' + municipios + '\tTitulos :\t'+titulos + '\rEntidades :\t'+entidades+ ' \tCodAto :\t'+referencias+ '\r\r--------------------------------------------------\r'  + msgErro   , 'Conferencia de titulos x atos', icon); 01/12/2015
    alert(alerta + 'Municipios :\t' + municipios + '\t\tTitulos :\t\t'+titulos + '\rEntidades :\t'+entidades+ ' \t\tCodAto :\t\t'+referencias+ '\r\rAssociações :\t' +associacoes+'\t\tConsórcios :\t'+consorcios+' \r\r--------------------------------------------------\r'  + msgErro   , 'Conferencia de titulos x atos', icon);
}

function encontrar(nome){
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
	app.findTextPreferences.appliedParagraphStyle = stilo;
	var myFoundItems;
	
	if(app.selection.length != 0){
		myFoundItems = app.selection[0].findText();
	}else{
		myFoundItems = myDocument.findText();
	}
	//alert("Found " + myFoundItems[0].texts[0].contents + " instances of the search string.");
	contador = contador + myFoundItems.length;
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


} //localizaErroCodato