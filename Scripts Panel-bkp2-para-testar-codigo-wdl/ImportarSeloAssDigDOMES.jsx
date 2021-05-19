/* 2021-05-02 Inserir imagem PDF que informa assinatura digital no QR Code (no canto inferior direito do sumario)
Uma vez que nao estamos mais assinando pelo token e assim nao inserindo a imagem de assinatura digital, esse campo ficava somente com o QR Code
inspirado no Renato que coloca a imagem de forma manual, esse scriipt insere a imagem assinatura-digital-dom.pdf no selo
No momento somente na pagina 01 , discutir se nao deve ser a mesma filosofia do QR Code (ai tem que inserir via Pagina mestre */
var printDebug = false;
var myVersao, myFile, myPage;
var rootFolder = (new File($.fileName)).parent;
if(printDebug) $.writeln("ImportarSeloAssDigDOMES.jsx : var rootFolder: " + rootFolder);
var img_selo_ass =  '/Logos/assinatura-digital-domes.pdf';


function main(){
    myVersao = "2021.05.02" // 2021.05.02 - 0.1
	var myDocument = app.documents.item(0);
	//Set the measurement units to points.
	myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
	myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
    myPage = myDocument.pages.item(0);
	myFile = rootFolder + img_selo_ass;
	
    myPlacePDF(myDocument,myPage,myFile);
}

function myPlacePDF(myDocument,myPage, myPDFFile){
   if(printDebug)  $.writeln("ImportarSeloAssDigDOMES.jsx : myPlacePDF.myDocument: " + myDocument);
   if(printDebug)  $.writeln("ImportarSeloAssDigDOMES.jsx : myPlacePDF.myPDFFile: " + myPDFFile);
    
	var myPDFPage;
    var myPDFPageReserva = rootFolder + '/Logos/qr_cinza.pdf';  //var myPDFPageReserva = rootFolder + "/Logos/extra-manus.pdf";
    app.pdfPlacePreferences.pdfCrop = PDFCrop.cropMedia;
	var myCounter = 1;
	var myBreak = false;
	var ehPaisagem = false;
    if(printDebug) $.writeln("myPlacePDF.myPage: " + myPage);
    app.pdfPlacePreferences.pageNumber = myCounter;
    
    if(printDebug)  $.writeln("(new File (myPDFPageReserva)).exists) :" + (new File (myPDFPageReserva)).exists);
    if(printDebug)  $.writeln("(new File (myPDFFile)).exists) : " + (new File (myPDFFile)).exists );

    if ((new File (myPDFFile)).exists ) {
       myPDFPage = myPage.place(File(myPDFFile), [0,0])[0];
     } else {
        myPDFPage = myPage.place(File(myPDFPageReserva), [0,0])[0];
    }
   
    var myFrame = myPDFPage.parent;
    paginaSeloSumario(myDocument, myFrame);

}//myPlacePDF


function paginaSeloSumario(myDocument, myFrame){
    var myFrameBounds = [268.3,137.4 ,289.6 ,203.236]; // y1, x1, y2, x2
    myFrame.geometricBounds = myFrameBounds;
    myFrame.fit(FitOptions.proportionally);
    myFrame.fit(FitOptions.frameToContent);
    //myFrame.parent.appliedMaster = myDocument.masterSpreads.item("B-Capa");
     if(printDebug)  $.writeln ("function paginaSeloSumario(myDocument, myFrame)" );
}//paginaSeloSumario

main();
 //app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.fastEntireScript, "Importando Destaque ...");