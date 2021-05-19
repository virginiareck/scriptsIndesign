var printDebug = true;
var myVersao, myFile, myPage;
var rootFolder = (new File($.fileName)).parent;
if(printDebug) $.writeln("ImportarTextoEdicaoExtra : var rootFolder: " + rootFolder);
var img_extra =  '/Logos/extra.pdf';


function main(){
    myVersao = "2020.03.27"
	var myDocument = app.documents.item(0);
	//Set the measurement units to points.
	myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.millimeters;
	myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.millimeters;
    myPage = myDocument.pages.item(0);
	myFile = rootFolder + img_extra;
	
    myPlacePDF(myDocument,myPage,myFile);
}

function myPlacePDF(myDocument,myPage, myPDFFile){
   //if(printDebug)  $.writeln("ImportarTextoEdicaoExtra : myPlacePDF.myDocument: " + myDocument);
   //if(printDebug)  $.writeln("ImportarTextoEdicaoExtra : myPlacePDF.myPDFFile: " + myPDFFile);
    
	var myPDFPage;
    var myPDFPageReserva = rootFolder + '/Logos/qr_cinza.pdf';  //var myPDFPageReserva = rootFolder + "/Logos/extra-manus.pdf";
    app.pdfPlacePreferences.pdfCrop = PDFCrop.cropMedia;
	var myCounter = 1;
	var myBreak = false;
	var ehPaisagem = false;
     //$.writeln("myPlacePDF.myPage: " + myPage);
    app.pdfPlacePreferences.pageNumber = myCounter;
    
    if(printDebug)  $.writeln("(new File (myPDFPageReserva)).exists) :" + (new File (myPDFPageReserva)).exists);
    if(printDebug)  $.writeln("(new File (myPDFFile)).exists) : " + (new File (myPDFFile)).exists );

    if ((new File (myPDFFile)).exists ) {
       myPDFPage = myPage.place(File(myPDFFile), [0,0])[0];
     } else {
        myPDFPage = myPage.place(File(myPDFPageReserva), [0,0])[0];
    }
   
    var myFrame = myPDFPage.parent;
    paginaTextoExtraSumario(myDocument, myFrame);

}//myPlacePDF


function paginaTextoExtraSumario(myDocument, myFrame){
    var myFrameBounds = [64.636,7.5,118.903,202.5]; // y1, x1, y2, x2
    myFrame.geometricBounds = myFrameBounds;
    myFrame.fit(FitOptions.proportionally);
    myFrame.fit(FitOptions.frameToContent);
    //myFrame.parent.appliedMaster = myDocument.masterSpreads.item("B-Capa");
     if(printDebug)  $.writeln ("function paginaTextoExtraSumario(myDocument, myFrame)" );
}//paginaSumarioTextoExtra

main();
 //app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.fastEntireScript, "Importando Destaque ...");