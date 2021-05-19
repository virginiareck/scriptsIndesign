var doc = app.activeDocument;
var rootElement = doc.xmlElements.item(0);

function abrirInfo() {
    
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
            
            var codato = codAtos[0].contents.replace(/[^\d]/g,'') ;
           
            // var  linkJumper = File(url+codato);  
             var linkJumper = File("/C/Windows/Temp/In-Tools.html"); 
                linkJumper.open("w");  
                $.writeln( linkJumper.write( '<html><head><META HTTP-EQUIV=Refresh CONTENT="0; URL=https://diariomunicipal.sc.gov.br/site/?r=ato/view&id=', codato, '"></head><body> <p></body></html>') );  
                linkJumper.close();  
                linkJumper.execute(); 
            
           /// alert("urlato : " + urlato + "\r\rcodato : " + codato + "\r\rjumper :  +linkJumper ");


        }
    }

}

abrirInfo();