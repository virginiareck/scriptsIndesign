function novo() {
     app.scriptPreferences.enableRedraw = false;

var doc,dir_xml,arq_xml;
var dia,mes,ano,diaSemana,edicao;

var semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo'];
var meses  = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

var rootFolder = (new File($.fileName)).parent;
var dtd = new File(rootFolder + '/Tools/DOM.dtd');
var script_file = new File(rootFolder + '/Tools/qrGeneratorwin64DOMSC.exe');

var img_assinatura = '/Logos/logo_cinza_assinatura_digital.pdf';
var img_logo = '/Logos/logo_diario_color.pdf';
var img_qr = '/Logos/qr_cinza.pdf';
var img_carimbo = '/Logos/assinatura-digital-dom.pdf';

var url = "www.diariomunicipal.sc.gov.br";

var sigla = 'DOM/SC';
var local = "Florianópolis/SC";

var estilo = new File(rootFolder + "/Estilos/XML_ESTILOS.idml");



var rodape = [
       [[270,5,274, 137], "CIGA - Consórcio de Informática na Gestão Pública Municipal"],
       [[274,5,278, 137], "Rua Gen. Liberato Bittencourt, n.º 1885 - Sala 102, Canto - CEP 88070-800 - Florianópolis / SC"],
       [[278,5,282, 137], "http://diariomunicipal.sc.gov.br"],
       [[282,5,286, 137], "Atendimento: Fone/Fax (48) 3321 5300"],
       //[[286,5,290, 137], "diagramador@diariomunicipal.sc.gov.br"],
	   [[286,5,290, 137], "dom@ciga.sc.gov.br"],
   ];

var uma_coluna = true; //Window.confirm('Criar edição com apenas uma coluna?');
var domes = false; //var domes = Window.confirm('Criar edição DOM/ES?');  //  foi absrvido no configuraVersao(sigla)  //else if (sigla == "DOM/ES") 
var styles;
var offset_y = 0;    //var offset_y = domes ? 20 : 0;   //  foi absrvido no configuraVersao(sigla)  //else if (sigla == "DOM/ES") 

/*if (domes) {          //colocado no configuraVersao(sigla)
   img_assinatura = '/Logos/logo_cinza_assinatura_digital_domes.pdf';
   img_logo = '/Logos/Logo DOMES.pdf';

   url = "www.diariomunicipal.es.gov.br";

   sigla = 'DOM/ES';
   local = "Vitória/ES";
   
   rodape = [
       [[270,5,274, 137], "AMUNES - Associação dos Municípios do Estado do Espírito Santo"],
       [[274,5,278, 137], "Avenida Princesa Isabel, 629 - Ed. Vitória Center - Sala 401 - Centro - Vitória/ES"],        
       [[278,5,282, 137], "CEP: 29010-904 - Telefones: (27) 3227-3077 | (27) 3222-4537"],
       [[286,5,290, 137], "contato@diariomunicipal.es.gov.br                                     www.diariomunicipal.es.gov.br"],
       [[282,5,286, 137], "Execução: CIGA - Consórcio de Informática na Gestão Pública Municipal"],
       
       
   ];
   
    estilo = new File(rootFolder + "/Estilos/Estilo_DOMES.indd");
}//if (domes)*/

function main() {
	app.documentPreferences.facingPages = false;    
    
    identificaVersao ();
    criarArquivo();
    estilos();
    importarXML();
    mestres();
    popular();
    // vincula os pdfs dos logotipos no documento principal
    importarLogos();
    importarDestaque();
}

//busca pelo atributo propriedades->sigla no xml de atos
function identificaVersao(){
   arq_xml = File.openDialog("Selecione o Arquivo XML");
   arq_xml.open("r");
   var tmp_obj = new XML ( arq_xml.read() );
   arq_xml.close();
   
   var prop = tmp_obj.children ()[0];
   if (prop.child("sigla")){
       sigla = prop.child("sigla");
       configuraVersao(sigla);
   }
}

function configuraVersao(sigla){   
   
  /* if (sigla == "DOM/ITAJAI"){
       //usa estilos do domes
       domes = true;
       img_assinatura = '/Logos/domitajai/logo_cinza_assinatura_digital_domitajai.pdf';
       img_logo = '/Logos/domitajai/logo_domitajai.pdf';

       url = "www.domi.ciga.sc.gov.br";

       local = "Itajaí/SC";
       
       rodape = [
           [[270,5,274, 137], "Prefeitura de Itajaí"],
           [[274,5,278, 137], "Rua Alberto Werner, nº 100 - Itajaí-SC"],        
           [[278,5,282, 137], "CEP: 88304-900 - Telefone: (47) 3341 6000"],
           [[286,5,290, 137], "contato@domi.ciga.sc.gov.br                                     www.domi.ciga.sc.gov.br"],
           [[282,5,286, 137], "Execução: CIGA - Consórcio de Informática na Gestão Pública Municipal"],
       ];        
        estilo = new File(rootFolder + "/Estilos/domitajai/Estilo_DOMITAJAI.indd");
    }
    else if (sigla == "DOE-TCE/SC"){
        //usa estilos do domes
        domes = true;
        img_assinatura = '/Logos/domtcesc/logo_cinza_assinatura_digital_domtcesc.pdf';
        img_logo = '/Logos/domtcesc/logo_domtcesc.pdf';

        url = "www.domtcesc.ciga.sc.gov.br";

        local = "TCE/SC";
        
        rodape = [
            [[270,5,274, 137], "Diário Oficial Eletrônico - Coordenação: Secretaria-Geral"],
            [[274,5,278, 137], "Rua Bulcão Vianna, nº 90, Florianópolis-SC."],        
            [[278,5,282, 137], "CEP: 88020-160 - Telefone:  (48) 3221-3648"],
            [[286,5,290, 137], "diario@tce.sc.gov.br                                     www.domtcesc.ciga.sc.gov.br"],
            [[282,5,286, 137], "Execução: CIGA - Consórcio de Informática na Gestão Pública Municipal"],
        ];        
         estilo = new File(rootFolder + "/Estilos/domtcesc/Estilo_domtcesc.indd");
     }*/
   
 //  else if (sigla == "DOM/ES") {
	   if (sigla == "DOM/ES") {
       domes = true;
       offset_y =20;
       img_assinatura = '/Logos/logo_cinza_assinatura_digital_domes.pdf';
       img_logo = '/Logos/Logo DOMES.pdf';

       url = "www.diariomunicipal.es.gov.br";

       sigla = 'DOM/ES';
       local = "Vitória/ES";
       
       rodape = [
           [[270,5,274, 137], "AMUNES - Associação dos Municípios do Estado do Espírito Santo"],
           [[274,5,278, 137], "Avenida Princesa Isabel, 629 - Ed. Vitória Center - Sala 401 - Centro - Vitória/ES"],        
           [[278,5,282, 137], "CEP: 29010-904 - Telefones: (27) 3227-3077 | (27) 3222-4537"],
           [[286,5,290, 137], "contato@diariomunicipal.es.gov.br                                     www.diariomunicipal.es.gov.br"],
           [[282,5,286, 137], "Execução: CIGA - Consórcio de Informática na Gestão Pública Municipal"],
           
           
       ];
       
        estilo = new File(rootFolder + "/Estilos/Estilo_DOMES.indd");
    }//elseif (domes)

	else{
		sigla = 'DOM/SC'
	}

}

function importarDestaque(){
   var destaque = dir_xml + '/destaque.pdf';
   
   if ((new File (destaque)).exists){
       app.scriptArgs.set('DestaqueAutomatico', destaque);
       app.doScript(new File(rootFolder + '/ImportarDestaquesAutomatico.jsx'), ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Executando...");
       app.scriptArgs.clear();
      // $.writeln ("function importarDestaque : Automatico = true" );
   } 

}

function importarLogos() {
   for(var n = 0; n < doc.links.length; n++) {
       var link = doc.links[n];
       if (link.status == LinkStatus.NORMAL) {
           link.unlink();
       }
   }    
}

function estilos() {
   doc.importStyles(ImportFormat.textStylesFormat,  estilo, GlobalClashResolutionStrategy.loadAllWithOverwrite);
   doc.importStyles(ImportFormat.tableAndCellStylesFormat, estilo, GlobalClashResolutionStrategy.loadAllWithOverwrite);
   doc.importStyles(ImportFormat.TOC_STYLES_FORMAT,  estilo, GlobalClashResolutionStrategy.loadAllWithOverwrite);
   doc.importStyles(ImportFormat.OBJECT_STYLES_FORMAT, estilo, GlobalClashResolutionStrategy.loadAllWithOverwrite);

   styles = doc.paragraphStyles;

   with (styles.item("entidade")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
   }

   with (styles.item("entidade - associacao")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
   }

   with (styles.item("entidade - consorcio")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
   }



   if (!domes) {
       
       with (styles.item("entidade - multas")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
       
       }//with

      with (styles.item("municipio - multas")) {

           keepWithNext = 5;
           keepAllLinesTogether = true;
       
           startParagraph = StartParagraph.NEXT_PAGE;
       
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
               
       }//with
	   

   }//if nao domes



   with (styles.item("municipio")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
           startParagraph = StartParagraph.NEXT_PAGE;
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
           
   }

   with (styles.item("municipio - consorcio")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
           startParagraph = StartParagraph.NEXT_PAGE;
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
   }



   with (styles.item("municipio - associacao")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
           startParagraph = StartParagraph.NEXT_PAGE;
           spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
           spanSplitColumnCount = SpanColumnCountOptions.ALL;
   }

   with (styles.item("titulo")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
   }

   with (styles.item("codAto")) {
           keepWithNext = 5;
           keepAllLinesTogether = true;
   }
}

function importarXML() {
	with (doc.xmlViewPreferences) {
        showAttributes = true;
        showStructure = true;
        showTaggedFrames = true;
        showTagMarkers = true;
        showTextSnippets = true;	
    }
	with(doc.xmlImportPreferences) {
        allowTransform = false;
        createLinkToXML = false;
        ignoreUnmatchedIncoming = false;
        ignoreWhitespace = false;
        importCALSTables = true;
        importStyle = XMLImportStyles.mergeImport;
        importTextIntoTables = false;
        importToSelected = false;
        removeUnmatchedExisting = false;
        repeatTextElements = true;
	}
	
    //arq_xml = File.openDialog("Selecione o Arquivo XML");
    dir_xml = arq_xml.path;    
    doc.importXML(arq_xml);
    
    if (dtd.exists) {
        doc.importDtd(dtd);
    }
    
    var rootElement = doc.xmlElements.item(0);
    var prop = rootElement.evaluateXPathExpression("//propriedades/dia");
    //$.writeln(prop[0]);
    dia = prop[0].contents;
    prop = rootElement.evaluateXPathExpression("//propriedades/mes");
    mes = prop[0].contents;
    prop = rootElement.evaluateXPathExpression("//propriedades/ano");
    ano = prop[0].contents;
    prop = rootElement.evaluateXPathExpression("//propriedades/diaSemana");
    diaSemana = prop[0].contents;
    prop = rootElement.evaluateXPathExpression("//propriedades/edicao");
    edicao = prop[0].contents;
    
    prop = rootElement.evaluateXPathExpression("//propriedades");
    prop[0].remove();
	
	doc.xmlImportMaps.add(doc.xmlTags.item("nomeEntidade"), styles.item("entidade"));
	doc.xmlImportMaps.add(doc.xmlTags.item("nomeGrupo"),    styles.item("municipio"));
    try { doc.xmlTags.add("integra"); } catch (e) {}
    try { doc.xmlTags.add("integraMono"); } catch (e) {}
    try { doc.xmlTags.add("grupoConsorcio"); } catch (e) {}
    try { doc.xmlTags.add("entidadeConsorcio"); } catch (e) {}
    try { doc.xmlTags.add("grupoAssociacao"); } catch (e) {}
    try { doc.xmlTags.add("entidadeAssociacao"); } catch (e) {}
    try { doc.xmlTags.add("grupoEditais"); } catch (e) {}
    try { doc.xmlTags.add("entidadeEditais"); } catch (e) {}
	try { doc.xmlTags.add("texto_cod_tcesc"); } catch (e) {}
	try { doc.xmlTags.add("cod_tcesc"); } catch (e) {}
    try { doc.xmlImportMaps.add(doc.xmlTags.item("integra"),      styles.item("texto")); } catch (e) {}
    try { doc.xmlImportMaps.add(doc.xmlTags.item("integraMono"),  styles.item("texto - mono")); } catch (e) {}
	doc.xmlImportMaps.add(doc.xmlTags.item("titulo"),       styles.item("titulo"));
	doc.xmlImportMaps.add(doc.xmlTags.item("codAto"),       styles.item("codAto"));
    
    doc.xmlImportMaps.add(doc.xmlTags.item("grupoConsorcio"), styles.item("municipio - consorcio"));
    doc.xmlImportMaps.add(doc.xmlTags.item("grupoAssociacao"), styles.item("municipio - associacao"));

    doc.xmlImportMaps.add(doc.xmlTags.item("entidadeConsorcio"), styles.item("entidade - consorcio"));
    doc.xmlImportMaps.add(doc.xmlTags.item("entidadeAssociacao"), styles.item("entidade - associacao"));

    // cria o estilo para os editais de notificacao
    //doc.paragraphStyles.item("municipio - consorcio").duplicate().name = "grupo - editais";
    //doc.paragraphStyles.item("entidade - consorcio").duplicate().name = "entidade - editais";

      
       if (!domes) {
          doc.xmlImportMaps.add(doc.xmlTags.item("entidadeEditais"), styles.item("entidade - multas"));
          doc.xmlImportMaps.add(doc.xmlTags.item("grupoEditais"),    styles.item("municipio - multas"));
		  doc.xmlImportMaps.add(doc.xmlTags.item("cod_tcesc"),styles.item("cod_tcesc"));
		  doc.xmlImportMaps.add(doc.xmlTags.item("texto_cod_tcesc"),styles.item("cod_tcesc"));
		}//if

    // substitui os grupos de associacao e consorcio
    var grupos = rootElement.evaluateXPathExpression("//folha/grupo/nomeGrupo");
    
    for(var g = 0; g < grupos.length; g++) {
        var nomeGrupo = grupos[g].xmlContent.contents;
        if (nomeGrupo == "AMUNES" || nomeGrupo == "Associações") {
            grupos[g].markupTag = "grupoAssociacao";
            var entidades = grupos[g].parent.evaluateXPathExpression("//entidade/nomeEntidade");
            for(var e = 0;e < entidades.length; e++) {
                entidades[e].markupTag = "entidadeAssociacao";
            }
        }
        if (nomeGrupo == "Consórcios Intermunicipais" || nomeGrupo == "Consórcios") {
            grupos[g].markupTag = "grupoConsorcio";
            var entidades = grupos[g].parent.evaluateXPathExpression("//entidade/nomeEntidade");
            for(var e = 0;e < entidades.length; e++) {
                entidades[e].markupTag = "entidadeConsorcio";
            }
        }
        if (nomeGrupo == "Editais de Notificação") {
            grupos[g].markupTag = "grupoEditais";
            var entidades = grupos[g].parent.evaluateXPathExpression("//entidade/nomeEntidade");
            for(var e = 0;e < entidades.length; e++) {
                entidades[e].markupTag = "entidadeEditais";
            }
        }
    }
    
	doc.mapXMLTagsToStyles();
	
}

function criarArquivo() {
   doc = app.documents.add();
   
   with(doc.viewPreferences) {
       horizontalMeasurementUnits = MeasurementUnits.millimeters;
       verticalMeasurementUnits = MeasurementUnits.millimeters;
       rulerOrigin = RulerOrigin.pageOrigin;
   }

   doc.colors.add({name:"Cinza", model:ColorModel.process, colorValue:[20,15,15,0]});

}

function mestres() {
   if (uma_coluna == false) {
       addCMaster();
   }
   addAMaster();
   addBCapa();
   addFPaisagem();    
}

function popular() {
   // Pagina 1 = CAPA
   doc.pages[0].appliedMaster = doc.masterSpreads.item("B-Capa");
   
   // Insere XML a partir da segunda Pagina
   var page = doc.pages.add();
   page.appliedMaster = doc.masterSpreads.item("A-Mestre");
   
   var frame = page.textFrames.add();
   
   // cria e popula a primeira página a partir do XML
   if (uma_coluna == false) {
       frame.geometricBounds = [15,8, 282, 100]; 
       doc.xmlElements[0].placeXML(frame);
       
   } else {
       frame.geometricBounds = [15,8, 282, 202];
       frame.textFramePreferences.textColumnCount = 1;
       frame.textFramePreferences.textColumnGutter = 8;
       
       doc.xmlElements[0].placeXML(frame);        
       var story = frame.parentStory;
       
       page = doc.pages.add();
       page.appliedMaster = doc.masterSpreads.item("A-Mestre");
   
       story.storyTitle = domes ? "DOM/ES" : "DOM/SC"; 
   }
   
       
}


// funcoes relacionadas as paginas mestre

function inserirFrameTexto(pagina, posicao, conteudo, tamanho, alinhamento) {
   var frame = pagina.textFrames.add();
   frame.geometricBounds = posicao;
   frame.contents = conteudo;
   frame.paragraphs[0].appliedFont = "Tahoma";
   frame.paragraphs[0].pointSize = tamanho;
   frame.paragraphs[0].justification = alinhamento;
   //frame.contents = conteudo;
   return frame;
}

function inserirImagem(alvo, posicao, arquivo) {
   var grafico = alvo.place(File(arquivo));
   //grafico[0].enableStroke = false;
   var frame = grafico[0].parent;
   frame.geometricBounds = posicao;
   frame.fit(FitOptions.PROPORTIONALLY);
   return frame;
}

function addFPaisagem(){
   var mestre = doc.masterSpreads.add({baseName: "Paisagem", namePrefix: "F"});
   var pagina = mestre.pages[0];
   
   //Retangulo Geral
   pagina.rectangles.add({ geometricBounds: [5,5, 290, 195], strokeWeight: 1 });
   
   //Linha Superior
   pagina.graphicLines.add({ geometricBounds: [5,12, 290, 12], strokeWeight: 1 });

   //Linha Cinza
   pagina.graphicLines.add({ geometricBounds: [10,10.5, 285, 10.5],
                             strokeColor: doc.colors.item("Cinza"),
                             strokeWeight: 4});
   
   //Logo CINZA
   if (img_assinatura != false) {
       var logo = inserirImagem (pagina, [286,5,296,60], rootFolder + img_assinatura);
       logo.rotationAngle = 90;
       logo.move([198, 290]);
   }
   
   // Numero da Pagina
   var numero = inserirFrameTexto (pagina, [6.5, 155, 12, 200], SpecialCharacters.autoPageNumber, 8, Justification.rightAlign);
   numero.insertionPoints[0].contents = "Página ";
   numero.rotationAngle = 90;
   numero.move([6.5, 55]);
   
   // Add Site   
   var site = inserirFrameTexto (pagina, [288, 140, 294, 206], url, 8, Justification.rightAlign );
   site.textFramePreferences.firstBaselineOffset = FirstBaseline.leadingOffset;
   site.rotationAngle = 90;
   site.move([198, 75]);
   
   //Add Data   
   var data = inserirFrameTexto (pagina, [6.5, 8, 12, 50], dia+"/"+mes+"/"+ano+" ("+semana[diaSemana-1]+")", 8, Justification.leftAlign);
   data.rotationAngle = 90;
   data.move([6.5, 285]);

   // frame unico
   pagina.textFrames.add({geometricBounds: [12,14, 285, 193] }); //pagina.textFrames.add({geometricBounds: [12,5, 285, 205] }); 2020-03-27
   
   // Add Edição
   var ed = inserirFrameTexto (pagina, [6.5, 65, 12, 140], sigla+" - Edição N° "+edicao, 8, Justification.centerAlign);
   ed.rotationAngle = 90;
   ed.move([6.5, 170]);
   
}
function addAMaster(){
   var mestre = doc.masterSpreads.item(0);
   mestre.baseName = "Mestre";
   mestre.namePrefix = "A";

   var pagina = mestre.pages[0];

   helperA_C(pagina);
   
   if (uma_coluna == false) {
       //Linha Central
       pagina.graphicLines.add({geometricBounds: [12,105, 285, 105], strokeWeight: 1});
       // Adiciona 2 quadros de texto
       var esquerda = pagina.textFrames.add({geometricBounds: [15,8, 282, 100]});
       var direita  = pagina.textFrames.add({geometricBounds: [15,110, 282, 202]});
       esquerda.nextTextFrame = direita;
       //esquerda.textWrapPreferences.textWrapMode = TextWrapModes.JUMP_OBJECT_TEXT_WRAP;
       //direita.textWrapPreferences.textWrapMode = TextWrapModes.JUMP_OBJECT_TEXT_WRAP;
   } else {
       var frame = pagina.textFrames.add({geometricBounds: [15,8, 282, 202] });
       
       // configuração das colunas
       frame.textFramePreferences.textColumnCount = 1;
       frame.textFramePreferences.textColumnGutter = 8;
       //frame.textFramePreferences.verticalBalanceColumns = true;
       //frame.textFramePreferences.verticalJustification = VerticalJustification.JUSTIFY_ALIGN;
       //frame.textFramePreferences.verticalThreshold = 3;
       
   }
}


// Pagina mestre normal sem linha no meio (uma unica coluna)
function addCMaster(){
   var mestre = doc.masterSpreads.add({baseName: "Mestre - sem linha",namePrefix: "C"});	
   var pagina = mestre.pages[0];
   
   helperA_C(pagina);
       
   // frame unico
   pagina.textFrames.add({geometricBounds: [15,8, 282, 202] });
   
}

// Constroi a pagina mestre da capa
function addBCapa(){
   var capa = doc.masterSpreads.add({baseName: "Capa", namePrefix: "B"});
   var pagina = capa.pages[0];
   
   // Retangulo Geral
   pagina.rectangles.add({geometricBounds: [5,5, 290, 205], strokeWeight: 1});
   
   //Rodape
   
   for(var i=0;i<rodape.length;i++) {
       inserirFrameTexto (pagina, rodape[i][0], rodape[i][1], 8, Justification.centerAlign);
   }
   
   // Logo Colorido
   if (domes) {
       inserirImagem (capa, [5,5,50,205], rootFolder + img_logo);
   } else {
       inserirImagem (capa, [10,8,35,201], rootFolder + img_logo);
   }
   
   // Linha Superior
   pagina.graphicLines.add({ geometricBounds: [45+offset_y,5, 45+offset_y, 205], strokeWeight: 1 });
   
   // Linha Inferior
   pagina.graphicLines.add({ geometricBounds: [268,   5, 268, 205], strokeWeight: 1 });
   pagina.graphicLines.add({ geometricBounds: [268, 137, 290, 137], strokeWeight: 1 });
  //  script_file.execute(edicao);
   

   //Linha QR Code
   qrFile = new File(dir_xml + "/qr.png");
   if (qrFile.exists ) {
       inserirImagem (capa, [269,181,288.8,201.5], qrFile);
   } else {
       inserirImagem (capa, [269,181,288.8,201.5], rootFolder + img_qr);
   }
   //Linha CARIMBO
   carimboFile = new File(rootFolder + img_carimbo);
   if (carimboFile.exists ) {
       inserirImagem (capa, [268.3,137.4 ,289.6 ,203.236], carimboFile);
   } else {
       inserirImagem (capa, [268.3,137.4 ,289.6 ,203.236], dir_xml + "/assinatura-digital-dom.pdf");
   }
   
   //Linha Vertical
   if (domes) {
       //pagina.graphicLines.add({ geometricBounds: [45+offset_y,105, 268, 105]});                    
   }
   // Linha cinza
   pagina.graphicLines.add({ geometricBounds:  [40+offset_y, 8, 40+offset_y, 202],
                            strokeColor: doc.colors.item("Cinza"),
                            strokeWeight: 22});
   // Data
   inserirFrameTexto (pagina, [38+offset_y, 61, 45+offset_y, 148], semana[diaSemana-1]+" - "+dia+" de "+meses[mes-1]+" de "+ano, 12, Justification.centerAlign);
   
   // Local (Cidade)
   inserirFrameTexto (pagina, [38+offset_y, 11, 45+offset_y, 198], local, 12, Justification.rightAlign);

   // Edição
   inserirFrameTexto (pagina, [38+offset_y, 11, 45+offset_y, 198], "Edição N° "+edicao, 12, Justification.leftAlign);
   
   //Quadro para o sumário
   pagina.textFrames.add({ geometricBounds: [54+offset_y, 8, 265, 200] });
   
   var sumario = inserirFrameTexto (pagina, [47+offset_y, 8, 52+offset_y, 202], "Sumário", 10, Justification.centerAlign);
   sumario.texts[0].fontStyle = "Bold";
   //sumario.texts[0].underline = true;

   // Margens
   with (pagina.marginPreferences) {
       left = 12,7;
       top = 12,7;
       right = 12,7;
       bottom = 12,7;
   }
}


// controi quase tudo que tem na A-Mestre e C-Mestre sem Linha
function helperA_C(pagina) {
   // Retangulo Geral
   pagina.rectangles.add({ geometricBounds: [5,5, 285, 205], strokeWeight: 1 });
   
   // Linha Superior
   pagina.graphicLines.add({ geometricBounds: [12,5, 12, 205], strokeWeight: 1 });

   // Linha cinza
   pagina.graphicLines.add({ geometricBounds: [10.5,8, 10.5, 202],
                             strokeColor: doc.colors.item("Cinza"),
                             strokeWeight: 4});
   // Data
   inserirFrameTexto (pagina, [6.5, 8, 12, 50], dia+"/"+mes+"/"+ano+" ("+semana[diaSemana-1]+")", 8, Justification.leftAlign);
   
   // Edição
   inserirFrameTexto (pagina, [6.5, 55, 12, 155], sigla+" - Edição N° "+edicao, 8, Justification.centerAlign);
   
   // Site
   inserirFrameTexto (pagina, [288, 160, 294, 202], url, 8, Justification.rightAlign);
   
   // Margens
   with (pagina.marginPreferences) {
       left = 12,7;
       top = 12,7;
       right = 12,7;
       bottom = 12,7;
   }

   // Logo pequeno inferior
   if (img_assinatura != false) {
       inserirImagem (pagina, [286,5,296,60], rootFolder + img_assinatura);
   }

   // Número de Página
   var numero = inserirFrameTexto (pagina, [6.5, 160, 12, 202], SpecialCharacters.autoPageNumber, 8, Justification.rightAlign);
   numero.insertionPoints[0].contents = "Página ";
}

main();

app.scriptPreferences.enableRedraw = true;

} // novo()

novo();

