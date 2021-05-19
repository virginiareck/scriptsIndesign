#target InDesign
#targetengine "ImportarAtos"
/*    2019-12-02 tentando resolver esse problema PDF
    http://edicao.dom.sc.gov.br/1575309526_edicao_2993_assinada.pdf#page=763
    752 function importarPaginaPDF(target, file, pageNum) {
    800     if (pageNum == 1) {    mudado de 230 para 220.
 */
/*  2019-12-02 Inserido o GREP para Multas
     executarGrep(integra, regrasMultas);
    var regrasMultas = [
    Para arrumar a assinatura das multas, remover o excesso de espaços
 var myDate em importando atos, importando PDF*/
File.prototype.fileExt = function(){
    return this.name.replace(/^.*\./,'');
}

Array.prototype.contains = function(obj) {
    var i = this.length;
    while (i--) {
        if (this[i] == obj) {
            return true;
        }
    }
    return false;
}
var myDoc = app.activeDocument;

var folder = new Folder(myDoc.fullName.path + '/arquivos/');

var status;

var atos = selecionarAtos(); 

if (atos.length == 0) exit();

var soffice;

// #TODO Remmover suporte às versões antigas e de 32 bits.
// LibreOffice 6 e/ou mais recentes.
if (Folder('/c/Program Files/LibreOffice').exists) {
    soffice = '"c:\\Program\ Files\\LibreOffice\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files/LibreOffice 4').exists) {
    soffice = '"c:\\Program\ Files\\LibreOffice\ 4\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files (x86)/LibreOffice 5').exists) {
    soffice = '"c:\\Program Files\ (x86)\\LibreOffice 5\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files (x86)/LibreOffice 4').exists) {
    soffice = '"c:\\Program Files\ (x86)\\LibreOffice 4\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files (x86)/LibreOffice 4.0').exists) {
    soffice = '"c:\\Program Files\ (x86)\\LibreOffice 4.0\\program\\soffice.exe"'; 
} else {
        Window.alert('Não foi possível encontrar o LibreOffice em seu computador. A conversão ODT não irá funcionar.');
}

var scriptFolder = (new File($.fileName)).parent;
var all2docx = 
"Set myApp = CreateObject(\"InDesign.Application\")\r\n"+
"Set objFSO = CreateObject(\"Scripting.FileSystemObject\")\r\n"+
"Set objWord = CreateObject(\"Word.Application\")\r\n"+
"objWord.Visible = False\r\n"+
" objWord.Documents.Open arguments(0)\r\n"+
" objWord.ActiveDocument.ConvertNumbersToText\r\n"+
" objWord.ActiveDocument.SaveAs arguments(1), 12\r\n"+
" objWord.ActiveDocument.Close\r\n"+
"objWord.Quit\r\n"
;
var runvbs =  new File(scriptFolder + "/Tools/Run.vbs");
var mergeZIP = '"' + (new File(scriptFolder + "/Tools/merge.bat")).fsName + '"';

// preferencias
var versaoDOM = myDoc.xmlElements[0].parentStory.storyTitle;

var formatacao = app.scriptArgs.get('formatacao') != 'false'; 
var imagens = app.scriptArgs.get('importIMG') != 'false'; 
var pdfs = app.scriptArgs.get('importPDF') != 'false';
var converterODT = app.scriptArgs.get('convertODT') != 'false'; 
var converterDOC = app.scriptArgs.get('convertDOC') != 'false';
var converterPDF = app.scriptArgs.get('convertZIP') != 'false';
var redistribuirColunas = app.scriptArgs.get('redisColunas') != 'true';
var converterTabelas = app.scriptArgs.get('convertTables') != 'false';

if (versaoDOM == '') {
    var formatacao = Window.confirm("Preservar formatação?", true);
    var imagens    = Window.confirm("Preservar imagens?", true);
    var converter  = Window.confirm("Converter atos?", true);    
}

// estilos

var estiloTexto = myDoc.paragraphStyles.item("texto");
var estiloTextoMono = myDoc.paragraphStyles.item("texto - mono");
var estiloNenhum = myDoc.characterStyles.item(0);
var estiloTabela = myDoc.paragraphStyles.itemByName("texto tabela");


with(app.wordRTFImportPreferences){
    convertBulletsAndNumbersToText = true;
    convertPageBreaks = ConvertPageBreaks.none;
    convertTablesTo = ConvertTablesOptions.unformattedTable;
    importEndnotes = false;
    importFootnotes = false;
    importIndex = false;
    importTOC = false;
    importUnusedStyles = false;
    preserveGraphics = imagens;
    preserveLocalOverrides = true;
    preserveTrackChanges = false;
    removeFormatting = false; //!formatacao;
    resolveCharacterStyleClash = ResolveStyleClash.resolveClashAutoRename;
    resolveParagraphStyleClash = ResolveStyleClash.resolveClashAutoRename;
    useTypographersQuotes = false;
}

function limparEstilosImportados() {
    var estiloTexto = myDoc.paragraphStyles.item('texto');
    var estilos = myDoc.paragraphStyles;
    for(var i=0; i < estilos.length; ) {
        if (estilos[i].imported && estilos[i] != estiloTexto) {
            estilos[i].remove(estiloTexto);
        } else {
            i++;
        }
    }
}

function limparFormatacao(integra, fonte, tamanho) {

    for(var p=0; p < integra.texts.count(); p++) {
        var k = integra.texts[p];
       
        k.rightIndent = 0;
        k.leftIndent = 0;
        
        k.firstLineIndent = 0;
        k.lastLineIndent = 0;
        
        k.spaceBefore = 0;
        if (k.spaceAfter > 10) {
            k.spaceAfter = 4;
        } else {
            k.spaceAfter = 2;
        }
        
        k.keepAllLinesTogether = false;
        //k.keepLinesTogether = 2;
        
        k.appliedFont = fonte;
        k.pointSize = tamanho;
        k.leading = Leading.AUTO;
        k.noBreak = false;
        
        //formatarGraficos(k);
    }   
    // corrige o alinhamento da primeira linha do ato quando este difere da segunda,
    // algo que acontece bastante no DOM/ES
    try {
        if (integra.paragraphs[1].justification == Justification.CENTER_ALIGN &&
            integra.paragraphs[0].justification != Justification.CENTER_ALIGN) 
        {
            integra.paragraphs[0].justification = Justification.CENTER_ALIGN;
        }
    } catch (e) {}
}

function manterInicioJunto(integra, linhas) {
    for(var r=0; r < linhas && r < integra.lines.count(); r++) {
        integra.lines[r].keepWithPrevious = true;    
    }
}

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

    return; 
}

function growColumns(table, factor) {
    for(var c=0;c<table.columns.length;c++) {
        var w = Math.floor(table.columns[c].width * factor);
        table.columns[c].width = w > 1.5 ? w : "1.5mm";
    }    
}


function formataTabelaRecursiva(table, tamanhoFonte) {
    var niveis = 0;
    for(var i=0; i<table.cells.count();i++) {
        var cell = table.cells[i];
            
        // remove formatação de texto
        if (!formatacao) {
            cell.texts.everyItem().applyParagraphStyle(estiloTabela);
            cell.clearCellStyleOverrides();
        } else {
            with (cell.texts.everyItem()) {
                allowArbitraryHyphenation  = false;
                appliedFont = "Verdana";
                rightIndent = 0;
                leftIndent = 0;
                
                firstLineIndent = 0;
                lastLineIndent = 0;
                
                spaceBefore = 0;
                spaceAfter = 2;
                
                keepAllLinesTogether = false;
                leading = Leading.AUTO;
                noBreak = false;
                
                pointSize = tamanhoFonte;
            }    
        }
    
        // configura algumas margens internas mínimas para a célula
        with (cell) {
            autoGrow = true;
            verticalJustification = VerticalJustification.CENTER_ALIGN;
            topInset = "1mm";
            bottomInset = "1mm";
            leftInset = "1mm";
            rightInset = "1mm";
            topLeftDiagonalLine = false;
            topRightDiagonalLine = false;
        }
    
        if (cell.tables.count()) {
            for(var t=0; t < cell.tables.count(); ) {
                var c = cell.tables[t];
                niveis = 1 + formataTabelaRecursiva(c, tamanhoFonte); // mais fundo
                
                if (c.cells.count() == 1) {
                    c.convertToText("\r","\r");
                } else {
                    t++;
                }
            }
        }
    }
    return niveis;
}

function formatarTabela(table, w, ato) {

    var profundidade = formataTabelaRecursiva(table, table.width > 200 ? 7 : 8);
    
    try {       
        
        if (!formatacao && profundidade == 0) {
            table.appliedTableStyle = "tabela geral";
            table.clearTableStyleOverrides();
        }
        
        table.spaceBefore = 2;
        table.spaceAfter = 2; 
    
        if (table.columnCount == 1 && table.rows.length == 1) {
            table.convertToText("\t","\r");
            return false; //res;
        }
        if (!redistribuirColunas) {
            var dif = w / table.width;
            growColumns(table, dif);
        } else {
            table.width = w;
        }
    } catch (e) { 
        $.writeln('formatarTabela: '+e);
        //return false; 
    }
    return true;
}

function transporAto(ato) {
    // logica para transpor colunas quando o ato transposto é o primeiro da entidade 
    // e pra quando a entidade é a primeira do grupo
    var anterior = ato.parent.xmlElements.previousItem(ato);
    var primeiroEntidade = anterior.markupTag.name != "ato";
    if (primeiroEntidade) {
        //$.writeln("primeiro ato da entidade transposto = "+anterior.markupTag.name);
        with (anterior.texts[0]) {
            spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
            spanSplitColumnCount = SpanColumnCountOptions.ALL;
            //spanColumnMinSpaceAfter = 20;
            spanColumnMinSpaceBefore = 20;
        }
        
        var entidade = anterior.parent;
        var anteriorEntidade = entidade.parent.xmlElements.previousItem(entidade);
        var primeiroGrupo = anteriorEntidade.markupTag.name != "entidade";
        
        if (primeiroGrupo) {
            //$.writeln("primeira entidade transposto = "+anteriorEntidade.markupTag.name);
            
            with (anteriorEntidade.texts[0]) {
                spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
                spanSplitColumnCount = SpanColumnCountOptions.ALL;
                ///spanColumnMinSpaceAfter = 20;
                spanColumnMinSpaceBefore = 20;
            }
        }
    }
    transporColunas(ato);
}
/*
function verificaOverflow(table) {
    for(var c=0; c<table.columns.length; c++) {
        if (table.columns[c].overflows) return true;
    }
    return false;
}*/

function formatarTabelas(ato) {
    //return true;
    // formata tabelas de até 3 colunas em meia página
    do {
        var podeTranspor = true, repete = false;
        var largura = 95;
        for(var p=0; p < ato.tables.count(); p++) {
            var k = ato.tables[p];
            k.recompose();
            if ( k.columnCount > 3 ) {
                largura = 195;
                break;
            }
        }
        
        //$.writeln(ato.tables.count());
        for(var p=0; p < ato.tables.count(); p++) {
            var k = ato.tables[p];

            if (! formatarTabela(k, largura) ) {
                // reinicia processamento
                repete = true;
                break;
            } else {
                podeTranspor = true; 
            }
        }
        if (!repete) {
            return podeTranspor && largura > 95; 
        }
    } while (true);
}

function converterTabelasTexto(integra) {
    
    while (integra.tables.count() > 0) {
        integra.tables[0].convertToText("\r", "\r");
    }
}

function formatarGraficos(integra) {
    var transpor = false;
    for(var p=0; p < integra.pageItems.count(); p++) {
        try {
            var k = integra.pageItems[p];
            if (!imagens) {
                k.remove();
            } else {
                k.clearObjectStyleOverrides();
                k.anchoredObjectSettings.anchoredPosition = AnchorPosition.ABOVE_LINE;
                k.anchoredObjectSettings.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
                var bounds = k.geometricBounds;
                var w = bounds[3] - bounds[1];
                if (w > 90) {
                    transpor = true  
                }
                k.fit(FitOptions.FRAME_TO_CONTENT);
            }
        } catch (e) {
                $.writeln('formatarGraficos: '+e);
        }
    }

    return transpor; 
}


function formatarNroContagemAto(nro){
    var nroForm = 0;
     $.writeln("var nroForm; " + nroForm + " " + nro.typeoff );
    if(nro < 10){
        nroForm = ("00" + myNumber).slice(-3);
        $.writeln("9 nroForm = "+ nroForm );
    }else if( nro < 100){
        nroForm = ("0" + myNumber).slice(-3);
        $.writeln("99 nroForm = "+ nroForm );
    }else{
        nroForm = nro;
        $.writeln("else nroForm = "+ nroForm );
    }
   $.writeln("return "+ nroForm );
    return  nroForm;
    
}//formatarNroContagemAto

function importarAto(ato, file, indAto) {
    var myDate = new Date().toTimeString().replace(/.*(\d{2}:\d{2}:\d{2}).*/, "$1");//2019-12-03 wdl
    //var totalAtosFormatado =  formatarNroContagemAto(atos.length); //201912 NAO FUNCIONOU
    var totalAtosFormatado =  atos.length;
    //var indAtoFormatado =  formatarNroContagemAto(++indAto);  //201912 NAO FUNCIONOU
    var indAtoFormatado =  ++indAto; 
            
    $.writeln(myDate + " - " + indAtoFormatado+ "/" + totalAtosFormatado + " - Importando ato - "+ file.name); //2019-05-28 wdl
    var integra = ato.evaluateXPathExpression("//integra|//integraMono")[0];

    try {
        integra.xmlContent.contents = '';
        integra.insertionPoints[0].place(file, false);
        
    } catch (e) {
        integra.xmlContent.contents = "\r\r##### Falha na importação ####\r\r";
        Window.alert( 'Falha no importarAto: '+e ); 
        return null;
    }
        
    if (!formatacao) {
        limparEstilosImportados();
    }

    return integra;
}

function reconstruirPDF(pdf) {
    var refeito = new File(pdf.path + '/refeito_' + pdf.name);
    var gs = '"' + File((new File($.fileName)).parent + "/Tools/gswin32c.exe").fsName.replace('/[ \]/g','\$1') + '"';

    var args = '-q -DSAFER -dNOPAUSE -sDEVICE=pdfwrite -dBATCH -sstdout=%stderr -sOutputFile="' + refeito.fsName + '" "' + pdf.fsName + '"';
    var cmd = '"' + gs + " " + args + '"';
    app.doScript(runvbs, ScriptLanguage.VISUAL_BASIC, [ cmd ]); 
    return refeito;
}

function importarPDF(ato, file) {
    var myDate = new Date().toTimeString().replace(/.*(\d{2}:\d{2}:\d{2}).*/, "$1");//2019-12-03 wdl
    $.writeln(myDate + " - importando PDF - "+ file.name); // 2019-05-28
    var integra = ato.evaluateXPathExpression("//integra|//integraMono")[0];
    integra.xmlContent.contents = " ";
    var anterior = ato.parent.xmlElements.previousItem(ato);
    var primeiroEntidade = anterior.markupTag.name != "ato";
    if (!primeiroEntidade) {
        ato.xmlContent.paragraphs[0].startParagraph = StartParagraph.NEXT_PAGE;
    }

    var rebuild = false;
    var ok = true;
    var pageNum = 1;
    while (ok) {
        try {
            ok = importarPaginaPDF (integra, file, pageNum);
            pageNum++;
        } catch (e) {
            if (!rebuild) {
                rebuild = true;
                file = reconstruirPDF(file);
            } else {
                $.writeln('Falha importar PDF pagina '+pageNum+' - '+e);
                ok = false;
            }
        }
    }

    return integra;
}

function importarPaginaPDF(target, file, pageNum) {

    app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_CONTENT_VISIBLE_LAYERS;
    app.pdfPlacePreferences.pageNumber = pageNum;
    
    var pdf;
    var ip = target.insertionPoints.lastItem();
    
    ip.startParagraph = StartParagraph.NEXT_PAGE;
    myDoc.recompose();  // força a criacao da caixa de texto
    var textFrame = ip.parentTextFrames[0];
    ip.startParagraph = StartParagraph.ANYWHERE;    

    pdf = ip.place(file, false)[0];
    if(pdf.pdfAttributes.pageNumber != pageNum){
        pdf.parent.remove(); // remove pagina adicional
  
        return false;
    }

    ip.contents = "\r";  
    myDoc.recompose();  // força a criacao da caixa de texto
    
    var rotate = false;
        
    var frame = pdf.parent;
    var textFrame_bounds;
    if (textFrame) { // aumenta o frame para caber o conteudo, depois diminui lá em baixo
        textFrame_bounds = textFrame.geometricBounds.slice(0);
        textFrame.geometricBounds = [0,0, 400, 400];
    }
    frame.fit(FitOptions.FRAME_TO_CONTENT);
        
    var bounds = pdf.parent.geometricBounds;
    var w = bounds[3] - bounds[1];
    var h = bounds[2] - bounds[0];

    if (w > 195) {
        // rotacao
        if (w > h) {
            h = w * 220 / w;
            w = 195;
            rotate = true;
        } else {
            h = h * 220 / h;
            w = 195;
        }
    }
    if (pageNum == 1) {
        if (h > 220) {
                w = w * 220 / h;
                h = 220;
        }
    } else {
        if (h > 245) {
                w = w * 245 / h;
                h = 245;
        }
    }
    
    pdf.parent.geometricBounds = [0,0,h,w];


    pdf.parent.anchoredObjectSettings.anchoredPosition = AnchorPosition.ABOVE_LINE;
    pdf.parent.anchoredObjectSettings.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
    pdf.fit(FitOptions.PROPORTIONALLY);
    if (rotate) {
        pdf.absoluteRotationAngle = 90;        
        pdf.fit(FitOptions.PROPORTIONALLY);
    }

    if (textFrame) {
            textFrame.geometricBounds = textFrame_bounds;
    }
  
    return true;
}

function converterArquivo(ato) {

    var codigo = ato.evaluateXPathExpression("//codAto")[0].xmlContent.contents;
    try {
        if (folder.getFiles(codigo + ".pdf").length > 0 && pdfs) {
            return folder.getFiles(codigo + ".pdf")[0];
        }

        var file = folder.getFiles(codigo + ".docx");
        if (file.length == 0) {
            file = folder.getFiles(codigo + ".*");
        }
        file = file[0] instanceof File ? file[0] : file[1];
        
        if (file.fileExt() == 'txt') {
            return file;
        }
        
        var doc = file.fsName.replace(/\.[a-z]+$/,'.docx');
                
        if (file.fileExt() == 'odt' && converterODT) {
            app.doScript(all2docx, ScriptLanguage.VISUAL_BASIC, [ file.fsName, doc ] ); 
            return folder.getFiles(codigo + ".docx")[0];
        } else if (file.fileExt() == 'zip' && converterPDF) {
            var args = '' + codigo + ' "' + file.fsName + '"';
            var cmd = '"' + mergeZIP + " " + args + '"';
            $.writeln('Combinando PDFs... '+file.fsName);
            app.doScript(runvbs, ScriptLanguage.VISUAL_BASIC, [ cmd ]); 
            
            return folder.getFiles(codigo + ".pdf")[0];
        } else if (file.fileExt() == 'doc' && converterDOC) {
                        
            app.doScript(all2docx, ScriptLanguage.VISUAL_BASIC, [ file.fsName, doc ] ); 
            
            return folder.getFiles(codigo + ".docx")[0];
        }
        else {
            return file;
        }

    } catch (e) {
        Window.alert('Não foi possível encontrar ou processar o arquivo do ato '+codigo);    
    }
}

function criarJanelaStatus() {
    status = new Window ("palette","Importando atos...");
    status.status = status.add("statictext",  undefined, "Preparando...");
    status.status.preferredSize = [450,20];
    status.pbar = status.add ("progressbar", undefined, 0, 0);
    status.pbar.color = [255,0,255];
    status.pbar.preferredSize = [450,20];
    status.pbar.value = 0;
	status.show ();
    
    status.onClose = function() {
        pararImportacao();
    };
}

function selecionarAtos() {   
    try {
         if (app.selection.length != 0) {
            app.findTextPreferences.appliedParagraphStyle = myDoc.paragraphStyles.item("codAto");
            app.findTextPreferences.findWhat = '';
            var atos = app.selection[0].findText();
            for (var i=0; i<atos.length;i++ ) {
                atos[i] = atos[i].associatedXMLElements[0];
            }
            return atos;
        } 
        else if (Window.confirm('Deseja importar TODOS os atos agora?')) 
        {
            return myDoc.xmlElements.item(0).evaluateXPathExpression("//ato").reverse();
            
        }
    } catch (e) {}
    
    return new Array();
    
}

function taskImportarAtos(evented) {
    var task = null;
    if (evented) {
        task = app.idleTasks.add({name:"importarAtos", sleep:500});
    }
    var i = 0, state = 0, integra, file, startTime = new Date().getTime();
    var taskHandler = function (event) {
        //$.writeln('task '+i+' state '+state);
        if (i < atos.length) {
            var ato = atos[i];
            var codigo = ato.evaluateXPathExpression("//codAto")[0].xmlContent.contents;
            var endTimeMS = (atos.length - i) * ((new Date().getTime() - startTime) / i);
            
            status.status.text = "(" + status.pbar.value + "/" + status.pbar.maxvalue + ") codigo: " + codigo + 
            " - etapa: " + state + " - estimado: " + Math.ceil(endTimeMS / 60000) + " minutos";
 
            // Primeira etapa: converte arquivos / descompacta zip
            if (state == 0) {
                app.scriptPreferences.enableRedraw = false;  
                
                if (file) {
                    state = 1;               
                } else {
                    // Se entrou aqui foi porque deu pau
                    // Não foi possível pegar o ato, pula pro próximo e deixa a integra como "=== Importar ato ==="
                    state = 3;
                }
            } 
            // Segunda etapa: Importa o arquivo, com lógica separada caso seja PDF
            else if (state == 1) {
                if (file.fileExt() == 'pdf') {
                    if (pdfs) {
                        transporAto(atos[i]);
                        integra = importarPDF(atos[i], file);
                    } else {
                        integra = null;
                    }
                    // A etapa 2 não é necessária para PDFs
                    state = 3;
                } else {
                    integra = importarAto(atos[i], file, i);
                    state = 2;
                }
            } 
            // Limpa o texto do ato importando e formata as tabelas
            else if (state == 2) {
                if (integra == null) {
                    // Integra vazia, não deve acontecer
                } else {
                    var transpor = false;
                    if (integra.markupTag.name == "integraMono") {
                        transpor = true;
                        integra.applyParagraphStyle( estiloTextoMono );
                        $.writeln("Iniciando GREP Multas");
                        executarGrep(integra, regrasMultas);
                        $.writeln("Fim GREP Multas\r----");
                    } else {
                        if (formatacao) {
                            limparFormatacao(integra, "Verdana", 9);    
                        } else {
                            integra.applyParagraphStyle( estiloTexto );
                        }
                    
                      
                        //manterInicioJunto (integra, 6);
                    
                        if (integra.tables.count() > 0) {
                            if (converterTabelas) {
                                converterTabelasTexto(integra);
                            }
                            else if (formatarTabelas(integra)) {
                                transpor = true;
                            }
                        }
                        if (integra.pageItems.count() > 0) {
                            if (formatarGraficos(integra)) {
                                transpor = true;
                            }
                        }
                    
                        if (formatacao) {
                            limparFormatacao(integra, "Verdana", 9);    
                        } else {
                            integra.applyParagraphStyle( estiloTexto );
                        }
                    }
                    if (transpor) {
                            transporAto(atos[i]);
                    }
                }
                state = 3;
            } else if (state == 3) {
                
                transporColunas(atos[i]);
                app.scriptPreferences.enableRedraw = true; 
                status.pbar.value++;
                state = 0;
                i++;    
            }                
                
        } else {
                //$.writeln('fim');
                status.hide();
                if (task) {
                    event.parent.sleep = 0; // remove a tarefa
                    task.remove();
                } 
                app.scriptPreferences.enableRedraw = true;   
                return false;
        }
        return true;
    };
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    //app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    status.pbar.value = 0;
    status.pbar.maxvalue = atos.length;
    if (task) {
        task.addEventListener(IdleEvent.ON_IDLE, taskHandler);
    } else {
        //app.scriptPreferences.enableRedraw = false;  
        while (taskHandler());
        //app.scriptPreferences.enableRedraw = true;  
    }
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
}

function pararImportacao() {
    var task = app.idleTasks.itemByName("importarAtos");
    if (task != null)
    {
        task.remove();
    }
}

function main() {

    criarJanelaStatus();

    if (atos) {
        pararImportacao ();
        // Deve ser habilitado apenas para depuração, pois deixa o processo muito lento
        //Window.confirm("Executar em paralelo? (Permite acompanhar o progresso visualmente)")
        taskImportarAtos(false);
    } else {
        status.hide();
    }
}

main();
//app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.fastEntireScript, "Importando...");
