﻿#target InDesign
//#targetengine "ImportarAtos"

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
} else if (Folder('/c/Program Files/LibreOffice 6').exists) {
    soffice = '"c:\\Program\ Files\\LibreOffice\ 6\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files/LibreOffice 4').exists) {
    soffice = '"c:\\Program\ Files\\LibreOffice\ 4\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files (x86)/LibreOffice 5').exists) {
    soffice = '"c:\\Program\ Files\ (x86)\\LibreOffice\ 5\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files (x86)/LibreOffice 4').exists) {
    soffice = '"c:\\Program\ Files\ (x86)\\LibreOffice\ 4\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files (x86)/LibreOffice 4.0').exists) {
    soffice = '"c:\\Program\ Files\ (x86)\\LibreOffice\ 4.0\\program\\soffice.exe"'; 
/*} else if (Folder('/c/Program Files (x86)/LibreOffice 6.0').exists) {
    soffice = '"c:\\Program\ Files\ (x86)\\LibreOffice\ 6.0\\program\\soffice.exe"'; 
} else if (Folder('/c/Program Files (x86)/LibreOffice 6').exists) {
    soffice = '"c:\\Program\ Files\ (x86)\\LibreOffice\ 6\\program\\soffice.exe"'; */
} else {
    Window.alert('Não foi possível encontrar o LibreOffice no seu computador. A conversão ODT não irá funcionar.');
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
var redistribuirColunas = app.scriptArgs.get('redisColunas') != 'false';
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
    with (ato.xmlContent) {
        spanColumnType = SpanColumnTypeOptions.SPAN_COLUMNS;
        spanSplitColumnCount = SpanColumnCountOptions.ALL;
        //spanColumnMinSpaceAfter = 20;
        //spanColumnMinSpaceBefore = 20;
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
        var podeTranspor = false, repete = false;
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

var regras = [
    // LIXO
    ["(ï¿½|â\\?\\?)", "", 0],
    ["\\r(-{77})-+", "\\r$1", 0], // traçados gigantescos ------------
    ["^[ _]+$", "", 0], // linhas que possuem apenas traços (geralmente assinaturas ou formulários)
    ["EMBED (Word\\.Picture\\.8|CorelDraw\\.Graphic\\.9)", "", 0],

    
    // Simbolos
    ["ï\\?¨", "— ", 0],
        
    // Espaçamento
    ["\\n", "\\r", 0] ,
    ['(?-m)^\r+', "", 0], // remove quebras de linha no inicio apenas
    //['(?-m)\r+$', "\r", 0], // trava indesign pois remove o </ato>
    ["(~S|\\t)+", " ", 0] ,
    ['','•',0], // marcador
    ["^[ \\t]+" , "",0] , // remove espaços no início da linha
    ["[ \\t]+$" , "",0] , // remove espaços no final da linha
//    ["^(.{10,30})[ \\t]{3,}()$" , "$1\\x08",0] , // troca espaços consecutivos por uma tab de alinhamento a direita (para assinaturas na mesma linha)
    ["[ ]{2,}", " ", 0], // troca genérica de espaços consecutivos
    //["[ \\t]+$" , "",0] , // remove espaços consecutivos no fim do parágrafo
	["\\r\\r[ \\t\\r]+", "\\r\\r",0], 
    ["\\r[ \\t\\r]+\\r", "\\r\\r",0] ,
    //["([ ]+\\r|\\r[\\t]+)", "\\r",0] ,
    ["\\r[ \\t]+","\\r",0],
    ["([^\\. ])(\\.{5,})([^\\. ])?","$1 $2 $3", 0], // Espaça sequencias longas de pontos para evitar overflow na caixa de texto toda 
    ["([^_ ])(_{5,})([^_ ])?","$1 $2 $3", 0], // Espaça sequencias longas de traços para evitar overflow na caixa de texto toda
    
];

// regras específicas pro DOM/SC
var verbos = "(ABRE|Abre|ADMITE|Admite|ADMITIR|Admitir|AUTORIZA|Autoriza|CONCEDE|Concede|CONTRATA|CONVOCA|Convoca|CONVOCAR|Convocar|DECRETA|Decreta|DISPENSA|Dispensa|DISPÕE|DISPOE|Dispõe|Dispoe|EXONERA|HOMOLOGA|HOMOLOGA|INSTITUI|NOMEAR|NOMEIA|PROMOVE|RESOLVE)"
var regrasDOMSC = [
	["([^\\r])[\\r]"+ verbos +"\\r([\\r]+)?" , "$1\\r\\r$2$3\\r",0] ,
	["[\\r]" + verbos + "([ ]+|[\\t]+)?:\\r[\\r]+" , "\\r$1:\\r",0] , //10/04/2014 wdl

    ["C O N C E D E R" , "CONCEDER",0] ,
    ["C O N T R A T A R" , "CONTRATAR",0] ,
    ["C O N V I D A R" , "CONVIDAR",0] ,
	["D E C R E T A R" , "DECRETAR",0] ,
	["N O M E A R" , "NOMEAR",0] ,     		
	["R E S O L V E M" , "RESOLVEM",0] ,
    ["R E T O R N A R" , "RETORNAR",0] ,
    
	
    ["C O N C E D E" , "CONCEDE",0] ,
    ["C O N T R A T A" , "CONTRATA",0] ,
    ["C O N V I D A" , "CONVIDA",0] ,
    ["C O N V O C A" , "CONVOCA",0] ,
	["D E C R E T A" , "DECRETA",0] ,
    ["N O M E I A" , "NOMEIA",0] ,     
	["(R E S O L V E|R e s o l v e)" , "RESOLVE",0] ,

    ["L E I", "LEI",0] ,
    ["P O R T A R I A ", "PORTARIA", 0],
	
	
	//Ementas
	["([0-9]+)(\\.)?\\r\\r"+ verbos  , "$1$2\\r$3",0] ,
	["(ESTADO DE SANTA CATARINA)\\r[\\r+](PREFEITURA MUNICIPAL DE[^\\r]+|MUNICIPIO DE[^\\r]+)", "$1\\r$2", 0],
    ["GABINETE DO PREFEITO\\r\\rPORTARIA", "GABINETE DO PREFEITO\\rPORTARIA", 0],
	["([^\\r]+)\\r(O Município de|O FUNDO MUNICIPAL)", "$1\\r\\r$2", 0],

	["(PREFEITURA MUNICIPAL DE[^\\r]+)\\r[\\r+](ESTADO DE SANTA CATARINA)", "$1\\r$2", 0],
	["(PREFEITURA MUNICIPAL DE[^\\r]+)\\r[\\r+](AVISO DE)", "$1\\r$2", 0],


    // Assinaturas
    
    ["1([0-9])(\\.)?\\r\\r([^\\r]+)\\r(Prefeit|PREFEIT|President|PRESIDENT|Diretor|DIRETOR)", 
     "1$1$2\\r$3\\r$4", 0], //2015-06-12 wdl GREP Pref ,Pres, Dir 

	 ["(Abdon Batista, SC,[^\\r]+)\\r+(Elmar Marino Mecabo)[ ]+([^\\r]+)\\r+(Prefeito Municipal em Exercicio)[ ]+([^\\r]+)", 
     "\\r$1\\r$2\\r$4\\r\\r$3\\r$5", 0],
     
    ["(ADELIANA DAL PONT)\\r(Prefeita Municipal)\\r([^\\r]+)\\r(Secret|Superint)", 
     "$1\\r$2\\r\\r$3\\r$4", 0],
     
    ["(Airton Antonio Reinehr|AIRTON ANTONIO REINEHR)[ \\r]+\\r(Prefeito Municipal)", 
     "$1\\r$2", 0],
     
    ["(ALBINO GONÇALVES PADILHA) (DARIO CESAR DE LINS)\\r\\r(Prefeito Municipal) (Sec\\. Mun\\. de Adm\\. e Fazenda)", 
     "$1\\r$3\\r\\r$2\\r$4", 0],
     
    ["(Alcimar de Oliveira)\\r\\r(Prefeito Municipal)",
     "$1\\r$2", 0],
     
    ["(ALCIONEI FRANÇA DA SILVA)\\r\\r(Secretário Municipal)",
     "$1\\r$2", 0],
     
    ["(Amarildo Paglia|AMARILDO PAGLIA)\\r\\r(Prefeito Municipal)", 
     "$1\\r$2", 0],
     
    ["(Ana Claudia Barizon Fontana da Luz)\\r\\r(Secretária de Administração e Fazenda)",
     "$1\\r$2", 0],
     
    ["(ANTÔNIO PAULO REMOR)[\\s\\r]+(Prefeito Municipal)", 
     "$1\\r$2", 0],
     
    ["(Arline Caon|ARLINE CAON)\\r\\r(Diretor)", 
     "$1\\r$2", 0],
     
    ["(CAROLINA DA COSTA TELMA)\\r\\r(Gestora)",
     "$1\\r$2", 0],
     
    ["(DANIEL BROERING HARGER)\\r\\r(Secretário de Administração)",
     "$1\\r$2", 0],
        
    ["(Enoi Scherer|ENOÍ SCHERER)\\r\\r(Prefeito Municipal)",
     "$1\\r$2", 0],
     
    ["(EVALDO JOSÉ GUERREIRO FILHO)\\r\\r(Prefeito de Porto Belo)", 
     "$1\\r$2", 0],
     
    ["(FLÁVIO ALBANO WENDLING)\\r\\r(Presidente)",
     "$1\\r$2", 0],
     
    ["(FÁTIMA LORETE CLEIN DA SILVA)\\r\\r(Responsável pelas Publicações)", 
     "$1\\r$2", 0],
     	 
	["Gabinete do Prefeito Municipal de([^\r]+), (aos |em )?([0-9])", 
     "Gabinete do Prefeito Municipal de$1,\\r$2$3", 0], //wdl 31/03/14

    ["(GARIBALDI ANTONIO AYROSO)\\r[\\r ]+(Prefeito Municipal)\\r[\\r ]+(GIVANILDO SILVA)\\r[\\r ]+(Secretário Municipal de Administração)", 
     "$1\\r$2\\r\\r$3\\r$4", 0],
     
    ["(GARIBALDI ANTONIO AYROSO)\\r\\r(Prefeito Municipal)", 
     "$1\\r$2", 0],
     
    ["Gian FRANCESCO VOLTOLINI", "GIAN FRANCESCO VOLTOLINI", 0],
	
	["(Gilberto Amaro Comazzetto - PREFEITO MUNICIPAL|GILBERTO AMARO COMAZZETTO - PREFEITO MUNICIPAL)", "GILBERTO AMARO COMAZZETTO\\rPREFEITO MUNICIPAL", 0], //25/4/14
	
	["\\rGilsoni Lunardi Albino\\r", "\\rGILSONI LUNARDI ALBINO\\r", 0], 	//wdl-02/04/14
	    
    ["(GISA APARECIDA GIACOMIN|Gisa Aparecida Giacomin)[\\r ]+(Prefeita Municipal)",
     "GISA APARECIDA GIACOMIN\\r$2", 0],
     
    ["(GIVANILDO SILVA)\\r([^\\r]+)\\r(Secretário Municipal de Administração)\\r(Contratado)",
     "$1\\r$3\\r\\r$2\\r$4", 0],
     
    ["(JUCÉLIO KREMER)\\r\\r(Prefeito Municipal)",
     "$1\\r$2", 0],
     
    ["(Lucilaine Mokfa Schwarz)\\r\\r(Secretária Municipal de Administração)",
     "$1\\r$2", 0],
     
    ["(MAURO JUNES POLETTO)\\r\\r(Prefeito Municipal)",
     "$1\\r$2", 0],
     
    ["(Maykel Roberto Laube)\\r\\r(Secretário de Educação, Cultura, Esporte e Lazer)",
     "$1\\r$2", 0],
     
    ["(MILTON LUIZ ESPINDOLA)\\r\\r(Diretor Geral)",
     "$1\\r$2", 0],
     
    ["(Miguel Edival Melniski) – (Presidente do CMAS)",
     "\\r$1\\r$2", 0],
     
    ["(MUNICIPIO DE MONTE CARLO) ([^\\r]+)\\r(Marcos Nei Correa Siqueira) ([^\\r]+)\\r(Contratante) (Contratada)",
     "$1\\r$3\\r$5\\r\\r$2\\r$4\\r$6", 0],
	 
    ["(NELSON CRUZ|Nelson Cruz)([\\r]+)(Prefeito Municipal)", 
     "NELSON CRUZ\\r$3", 0],     

	["(Novelli Sganzerla) (Alexander de Carvalho Fabro) (Prefeito) (Diretor do Dpto de RH)",
     "$1\\r$3\\r\\r$2\\r$4", 0],

    ["(Osvaldo Jurck)([\\r]+)(Prefeito Municipal)",
     "OSVALDO JURCK\\r$3", 0], //10-04-14 wdl
	 
    ["(OSVALDO JURCK) (MAYKEL ROBERTO LAUBE)\\r\\r(Prefeito Municipal) \\t(Secretário de Educação, Cultura, Esporte e Lazer)",
     "$1\\r$3\\r\\r$2\\r$4", 0],
     
    ["(Patrícia Schmitz)\\r\\r(DG - Expediente)",
     "$1\\r$2", 0],
       
    ["(PAULO ROBERTO ECCEL)\\r(ARNALDO F\\. DA SILVA)\\r(CRISTIANO BITTENCOURT)\\r(ANTÔNIO C\\. TILLMANN)\\r(Prefeito Municipal)\\r(Secretário Orç. E Gestão)\\r(Cont\\. CRC 028895/O-9)\\r(CGM - Controle Interno)",
     "$1\\r$5\\r\\r$2\\r$6\\r\\r$3\\r$7\\r\\r$4\\r$8", 0],
     
    ["(Publicada no Diário Oficial dos)\\r+(Municípios de Santa Catarina)\\r+(www.diariomunicipal.sc.gov.br e)\\r+(Registrada no Livro de Publicações)", 
     "$1 $2 $3 $4", 0],
     
    ["(Publicada a presente Portaria)[\\r ]+(nesta Secretaria em data supra.)", 
     "$1 $2", 0],
	
	["(Publicada e registrada na forma da Lei)([^\\r]+)(em _)", 
     "$1$2\r$3", 0], // 20/08/2014 Atos retroativos campo alegre
	 
	 ["(Registrada e publicada na forma da Lei)([^\\r]+)(em _)", 
     "$1$2\r$3", 0], // 20/08/2014 Atos retroativos campo alegre
	 
    
    ["(RONALDO CARLESSI)\\r\\r(Prefeito Municipal)", 
     "$1\\r$2", 0],

    ["(RUBENS BLASZKOWSKI|Rubens Blaszkowski)([\\r]+)(Prefeito Municipal)", 
     "RUBENS BLASZKOWSKI\\r$3", 0],
	 
     
    ["(SHIRLEY NOBRE SCHARF)\\r\\r(Secretária de Educação)", 
     "$1\\r$2", 0],
     
    ["(TIAGO RAFAEL MUCHALSKI PETRY)\\r\\r(Assessor de Planejamento, Gestão e Finanças)",
     "$1\\r$2", 0],
    
    ["(TIAGO ZILLI)\\r\\r(Prefeito Municipal em Exercício)", 
     "$1\\r$2", 0],
    
    ["(VALMOR LUIZ DALL´AGNOL)\\r\\r(Secretário de Administração)", 
     "$1\\r$2", 0],
     
    ["(WILMAR CARELLI)\\r\\r(Prefeito Municipal)", 
     "$1\\r$2", 0],
     
    ["(WALDIR GIRARDI)\\r\\r(Diretor Presidente)",
     "$1\\r$2", 0],

    // Cabeçalhos
    
    ["(ESTADO DE SANTA CATARINA)\\r\\r(PREFEITURA MUNICIPAL DE SÃO JOSÉ)\\r\\r(SECRETARIA DE ADMINISTRAÇÃO)", 
     "$1\\r$2\\r$3", 0],
     
    ["(ESTADO DE SANTA CATARINA)\\r\\r(FUNDO MUNICIPAL DE SAÚDE DE PALHOÇA)\\r\\r(SECRETARIA MUNICIPAL DE SAÚDE)\\r",
     "$1\\r$2\\r$3", 0],
     
    ["(ESTADO DE SANTA CATARINA)\\r\\r(MUNICIPIO DE) ([^\\r]+)\\r\\r(AVISO)", 
     "$1\\r$2 $3\\r$4", 0],
     
    ["(ESTADO DE SANTA CATARINA)\\r\\r(PREFEITURA DE NOVA TRENTO)\\r\\r(Processo Licitatório)",
     "$1\\r$2\\r$3", 0],
     
    ["(ESTADO DE SANTA CATARINA)\\r\\r(PREFEITURA MUNICIPAL DE) ([^\\r]+)\\r\\r(AVISO)",
     "$1\\r$2 $3\\r$4", 0],
     
];

function executarGrep(target, regras) { 
    
    for(var index = 0; index < regras.length; index++){
        alterarGrep(target,regras[index][0],regras[index][1]);
    }
}

function alterarGrep(doc, original, alterado) {

	app.findGrepPreferences.findWhat = original;
	app.changeGrepPreferences.changeTo = alterado;
    app.findTextPreferences.appliedParagraphStyle = NothingEnum.nothing;
    
    try {
        var grep = doc.findGrep();
       
        for(var i=0;i<grep.length;i++) {
            grep[i].changeGrep();
        }

        return grep.length;
    } catch (e) {
        $.writeln('exception '+e);
        return 0;
    }
}

function importarAto(ato, file) {
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
            h = w * 230 / w;
            w = 195;
            rotate = true;
        } else {
            h = h * 230 / h;
            w = 195;
        }
    }
    if (pageNum == 1) {
        if (h > 230) {
                w = w * 230 / h;
                h = 230;
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
            
            var args = '--norestore --invisible --nologo --nodefault --headless --convert-to docx --outdir "' + file.parent.fsName + '" "' + file.fsName + '"';
            var cmd = '"' + soffice + " " + args + '"';
            app.doScript(runvbs, ScriptLanguage.VISUAL_BASIC, [ cmd ]); 
            
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
                file = converterArquivo(atos[i]);
                
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
                    integra = importarAto(atos[i], file);
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
                    } else {
                        if (formatacao) {
                            limparFormatacao(integra, "Verdana", 9);    
                        } else {
                            integra.applyParagraphStyle( estiloTexto );
                        }
                    
                        executarGrep(integra, regras);
                        if (versaoDOM == 'DOM/SC') {
                            executarGrep(integra, regrasDOMSC);
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
