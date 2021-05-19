/* // ATUALIZAÇÕES
   //2019-09-02 - Opçao  Cabeçalho / Rodapé -  myCabRod - myCabRod - paginaCabRodape
  //2019-09-02 - paginaCabecalho(myDocument, myFrame)  -  [45.2,5.2,289.8,204.8]; // y1, x1, y2, x2
   //Cabeçalho (200x244mm)" > 200x245
 //2019-08-26 -  staticLabel:"Cabeçalho (200x245mm)",   -     staticLabel:"Inteira (297x210mm)", 
//2016-01-07  - ImportarAniversarios recebe nova versao : importarDestaques
//2017-09-26  - https://ciga.sc.gov.br/diario-oficial-dos-municipios-qr-code/


//  melhorias futuras   //
Orientacao horizontal Radios - Flip?? - http://jongware.mit.edu/idcs6js/pc_RadioButton.html
Imagem thumnail para raio cabecalho / inteira
Tratativa para sem link
*/

var myNumero, myURL,myURLField, myPgQR,myPgCabecalho,myPgInteira, myVersao, myFile, myPage;

//main();


function main(){
    myVersao = "2019.09.02"
	var myDocument = app.documents.item(0);
	//Set the measurement units to points.
	myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
	myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
	//Get the current page.
	
	pegarPagina();
	myNumero--;
    myPage = myDocument.pages.item(myNumero);
	myFile = File.openDialog ("Selecione o Arquivo PDF Destaque");
	
	if(myPgQR == true){
		paginaQR(myDocument,myPage,myFile);
	}else{
		myPlacePDF(myDocument,myPage,myFile);
	}
}
function pegarPagina(){
	var myDialog = app.dialogs.add({name:"Página para Adicionar PDF Destaque               vr : " + myVersao, canCancel:true});
	with(myDialog){
		//Add a dialog column.
		with(dialogColumns.add()){
			//Create a border panel.
			with(borderPanels.add()){
				with(dialogColumns.add()){
					//The following line shows how to set a property as you create an object.
                        alignment:'left';
                        justify:'left';
                        ////         "     "
					staticTexts.add({staticLabel:"ANTES da Página:                                        ",minWidth:260,alignment:'left',justify:'left'});
				}
				with(dialogColumns.add()){
					//The following line shows how to set multiple properties
					//as you create an object.
					var myNumeroEditField = textEditboxes.add({editContents:"1", minWidth:60});
				}
 
			}
        
        	with(borderPanels.add()){
				with(dialogColumns.add()){
					//The following line shows how to set a property as you create an object.
					staticTexts.add({staticLabel:"URL"});
				}
				with(dialogColumns.add()){
					//The following line shows how to set multiple properties
					//as you create an object.
					//myURLField = textEditboxes.add({editContents:"https://diariomunicipal.sc.gov.br/site/?r=site%2Fedicao&edicao=9999", minWidth:400});
					myURLField = textEditboxes.add({editContents:"https://ciga.sc.gov.br/diario-oficial-dos-municipios-qr-code/", minWidth:400});
					
				}
 
			}
		
            
			with(borderPanels.add()){
                with(dialogColumns.add()){
                    staticTexts.add({staticLabel:"Espaço da Pagina:                                                       ",alignment:'left'});
                    var myRBG1 = radiobuttonGroups.add();
						with(myRBG1){
                                var myVD1QRCode = radiobuttonControls.add({staticLabel:"QR Code", checkedState:true, minWidth:280});
                                var myVD1CabRod = radiobuttonControls.add({staticLabel:"Cab e Rodapé (200x223mm)", checkedState:false, minWidth:280});
							var myVD1Cabecalho = radiobuttonControls.add({staticLabel:"Cabeçalho (200x245mm)",checkedState:false, minWidth:280});
							var myVD1Inteira = radiobuttonControls.add({staticLabel:"Inteira (297x210mm)", checkedState:false,alignment:'left'});
							//radiobuttonControls.orientation='row';						
						}//with
					
				}
				with(dialogColumns.add()){
					//The following line shows how to set multiple properties
					//as you create an object.
				
				}		
			}
		}
				
	}

            
             
	//Display the dialog box.
	if(myDialog.show() == true){
		myNumero = myNumeroEditField.editContents;
        myURL = myURLField.editContents;
        
        if(myRBG1.selectedButton == 0){
                myPgQR = true;
                myCabRod =  false;
                myPgCabecalho = false;
                myPgInteira = false;
                //alert("myPgQR : " + myPgQR);
        } else  if(myRBG1.selectedButton == 1){
                myPgQR = false;
                myCabRod =  true;
                myPgCabecalho = false;
                myPgInteira = false;
              //  alert("myCabRod : " + myCabRod);
        }else if(myRBG1.selectedButton == 2){
                myCabRod =  false;
                myPgQR = false;
                myPgCabecalho = true;
                myPgInteira = false;
                //alert("myPgCabecalho " + myPgCabecalho);
         }else if(myRBG1.selectedButton ==3){
                myCabRod =  false;
                myPgQR = false;
                myPgCabecalho = false;
                myPgInteira = true;
                //alert("myPgInteira " + myPgInteira);
         }else{
             myCabRod =  false;
             myPgQR = false;
             myPgCabecalho = false;
             myPgInteira = false;
             alert ("Erro na selecao dos botoes radio true - false", "Erro Radio true - false", true);
        }//ifelse

        //$.writeln ("myPgCabecalho:  " + myPgCabecalho.valueOf() + "\t\t\tmyPgInteira : " + myPgInteira.valueOf()  + "   \t\t\t selected : " + myRBG1.selectedButton);
		myDialog.destroy();
		
		
	}else{
		myDialog.destroy()
	}
}
///FIM function pegarPagina(){

function myPlacePDF(myDocument,myPage, myPDFFile){
	var myPDFPage;
	app.pdfPlacePreferences.pdfCrop = PDFCrop.cropMedia;
	var myCounter = 1;
	var myBreak = false;
	var ehPaisagem = false;
	while(myBreak == false){
		myPage = myDocument.pages.add(LocationOptions.before, myPage);
		
		app.pdfPlacePreferences.pageNumber = myCounter;
		myPDFPage = myPage.place(File(myPDFFile), [0,0])[0];
         
        //$.writeln ("myFrame B-Capa - URL:" + myURL);
              
		var myFrame = myPDFPage.parent;
        //alert("myPgCabecalho: " + myPgCabecalho + "\t\tmyPgInteira: " + myPgInteira);
        
         if(myCabRod == true){
			paginaCabRodape(myDocument, myFrame);
		 }else if(myPgQR == true){
			//paginaQR(myDocument, myFrame);
		 }else if(myPgCabecalho == true){
			paginaCabecalho(myDocument, myFrame);
		 }else if(myPgInteira == true){
			 paginaInteira(myDocument, myFrame);
		 }else{
			 alert ("Não selecionadoo tipo de Pagina : Com cabeçalho ou Pagina Inteira.", "Erro tipo de página", true);
			 break;
		  }

		inserirLink(myDocument, myFrame);
           
		
		if(myCounter == 1){
			var myFirstPage = myPDFPage.pdfAttributes.pageNumber;
		}else{
			if(myPDFPage.pdfAttributes.pageNumber == myFirstPage){
				myPDFPage.parent.parent.remove();
				myBreak = true;
			}
		}
		myCounter = myCounter + 1;
	}
}

function inserirLink(myDocument, myFrame){
    var myHyperlinkURL = myDocument.hyperlinkURLDestinations.add(myURL);
    var myHyperlinkSource = myDocument.hyperlinkPageItemSources.add(myFrame);
    var  myHyperlink=myDocument.hyperlinks.add(myHyperlinkSource,myHyperlinkURL );
    myHyperlink.visible=false;    //borda do hyperlink
    // $.writeln ("function: function (myFrame) : " + myFrame);

}

function inserirImagem(alvo, posicao, arquivo) {
	var grafico = alvo.place(File(arquivo));
    //grafico[0].enableStroke = false;
	var frame = grafico[0].parent;
	frame.geometricBounds = posicao;
	frame.fit(FitOptions.PROPORTIONALLY);
    return frame;
}



function paginaQR(myDocument, myFrame){
		var doc = app.activeDocument;
		var master_spread1_pg1 = doc.masterSpreads.item(1).pages.item(0);
    
    if (myFile.exists) {
        var imagem = inserirImagem (master_spread1_pg1, [269,181,288.8,201.5], myFile);
		inserirLink(doc, imagem);
    }
    
	$.writeln ("myPage : " + myPage + "\rmyPgQR : " + myPgQR);
   /* var myFrameBounds = [279,191,298.8,211.5];
    myFrame.geometricBounds = myFrameBounds;
    myFrame.fit(FitOptions.proportionally);
    myFrame.fit(FitOptions.frameToContent);
    myFrame.parent.appliedMaster = myDocument.masterSpreads.item("B-Capa");
     $.writeln ("function paginaQR(my..., myFrame)" );*/
}

function paginaCabRodape(myDocument, myFrame){
    var myFrameBounds = [45.2,5.2,267.8,204.8]; // y1, x1, y2, x2
    myFrame.geometricBounds = myFrameBounds;
    myFrame.fit(FitOptions.proportionally);
    myFrame.fit(FitOptions.frameToContent);
    myFrame.parent.appliedMaster = myDocument.masterSpreads.item("B-Capa");
     $.writeln ("function paginaCabRodape(myDocument, myFrame)" );
}

function paginaCabecalho(myDocument, myFrame){
    var myFrameBounds = [45.2,5.2,289.8,204.8]; // y1, x1, y2, x2
    myFrame.geometricBounds = myFrameBounds;
    myFrame.fit(FitOptions.proportionally);
    myFrame.fit(FitOptions.frameToContent);
    myFrame.parent.appliedMaster = myDocument.masterSpreads.item("B-Capa");
     $.writeln ("function paginaCabecalho(myDocument, myFrame)" );
}

function paginaInteira(myDocument, myFrame){
    //16-12-01//var myFrameBounds = [-5,-5,302,215];
    var myFrameBounds = [0,0,297,210];
    myFrame.geometricBounds = myFrameBounds;
    myFrame.fit(FitOptions.proportionally);
    myFrame.fit(FitOptions.frameToContent);
    myFrame.parent.appliedMaster = myDocument.masterSpreads.item(0);
     $.writeln ("function paginaInteira" );
}

 app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.fastEntireScript, "Importando Destaque ...");