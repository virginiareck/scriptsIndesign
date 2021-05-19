#target InDesign
#targetengine "interface"

var basePath = (new File($.fileName)).parent + '/';

function runScripts(scripts) {
    for(var s=0; s<scripts.length; s++) {
        if (typeof(scripts[s]) === "function") {
            scripts[s]();
        } else {
            app.doScript(new File(basePath + scripts[s]), ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Executando...");
        }
    }
}

var checkDOC;
var checkODT;
var checkZIP;
var checkPDF;
var checkIMG;
var checkFMT;
var checkCOL;
var checkTBL;


function setArgs() {
    app.scriptArgs.set('convertDOC', ''+checkDOC.value);
    app.scriptArgs.set('convertODT', ''+checkODT.value);
    app.scriptArgs.set('convertZIP', ''+checkZIP.value);
    app.scriptArgs.set('importPDF', ''+checkPDF.value);
    app.scriptArgs.set('importIMG', ''+checkIMG.value);
    app.scriptArgs.set('formatacao', ''+checkFMT.value);
    app.scriptArgs.set('redisColunas', ''+checkCOL.value);
    app.scriptArgs.set('convertTables', ''+checkTBL.value);
}

function criarJanela() {
    var window = new Window("palette", "Interface DOM 2017.08.30 (Home Office)");
    window.alignChildren = 'fill';
    var font = ScriptUI.newFont('Tahoma', 'BOLD', 18);
    
    var tpanel = window.add("tabbedpanel");
    var tab = new Array();
    
    tab[0] = tpanel.add("tab", undefined, "Abre");
    tab[1] = tpanel.add("tab", undefined, "Diagrama");
    tab[2] = tpanel.add("tab", undefined, "Fecha");
   // tab[3] = tpanel.add("tab", undefined, "Todos");
    
    var btns = new Array();
  
    btns[0] = [
            'Abertura',
        ['Criar Edição', [ 'Novo.jsx' ] ],
        ['Importar Atos',   [ setArgs, 'ImportarAtos.jsx' ] ],
        ['Importar Destaques',   [ 'ImportarDestaques.jsx' ] ],
    ];
    
    btns[1] = [
            'Diagramação',
        ['Abrir Documento', [ 'AbrirDOC.jsx'  ] ],
        ['Abrir Informações', [ 'AbrirInfoAto.jsx'  ] ],
        ['DOC --> PDF',    [ 'ConverterParaPDF.jsx'  ] ],
        ['PDF --> DOC',    [ 'ConverterParaDOC.jsx'  ] ],
        ['Alternar Colunas',[ 'AlternarColunas.jsx'  ] ],
        ['Expandir Colunas',[ 'ExpandirColunas.jsx'  ] ],
      ];
      
    btns[2] = [
        'Fechamento',
        ['Conferência',     [ 'Conferencia.jsx'  ] ],
        ['Sumário',         [ 'Paginas.jsx', 'Sumario.jsx'  ] ],
        ['(Re)Vincular PDFs',   [ 'EmbarcarLinks.jsx'  ] ],
        ''
    ];
    /*
    btns[3] = [
        'Abertura',
        ['Criar Edição', [ 'Novo.jsx' ] ],
        ['Importar Atos',   [ setArgs, 'ImportarAtos.jsx' ] ],
        ['Importar Destaques',   [ 'ImportarDestaques.jsx' ] ],
        'Diagramação',
        ['Abrir Documento', [ 'AbrirDOC.jsx'  ] ],
        ['Abrir Informações', [ 'AbrirInfoAto.jsx'  ] ],
        ['DOC --> PDF',    [ 'ConverterParaPDF.jsx'  ] ],
        ['PDF --> DOC',    [ 'ConverterParaDOC.jsx'  ] ],
        ['Alternar Colunas',[ 'AlternarColunas.jsx'  ] ],
         'Fechamento',
        ['Conferência',     [ 'Conferencia.jsx'  ] ],
        ['Sumário',         [ 'Paginas.jsx', 'Sumario.jsx'  ] ],
        ['(Re)Vincular PDFs',   [ 'EmbarcarLinks.jsx'  ] ],
        ''
       
    ];
    */

for(var j=0; j < btns.length; j++) {
    for(var i=0; i < btns[j].length; i++) {
            if (typeof btns[j][i] == 'string') {
                var txt = tab[j].add("StaticText", undefined);
                txt.text = btns[j][i];
                txt.justify = 'center';
            } else {
                //var btn = window.add("Button", undefined);
                var btn = tab[j].add("Button", undefined);
                
                var scripts = btns[j][i][1];
                btn.text = btns[j][i][0];
                btn.onClick = (function(scripts) { 
                    return function() {
                        runScripts(scripts);    
                    };
                }(scripts));
                btn.graphics.font = font;
               // $.writeln(btn.onClick.toSource());
            }
        }//for i
    } // for j

    checkODT = window.add("Checkbox", undefined, 'Converter ODT -> DOCX');
    checkODT.value = true;
    checkDOC = window.add("Checkbox", undefined, 'Converter DOC -> DOCX');
    checkDOC.value = true;
    checkZIP = window.add("Checkbox", undefined, 'Extrair PDFs em ZIP');
    checkZIP.value = true;
    checkPDF = window.add("Checkbox", undefined, 'Importar PDFs');
    checkPDF.value = true;
    checkIMG = window.add("Checkbox", undefined, 'Importar Imagens');
    checkIMG.value = false; //2015-07-23 false
    checkFMT = window.add("Checkbox", undefined, 'Manter Formatação');
    checkFMT.value = false;
    checkCOL = window.add("Checkbox", undefined, 'Redistribuir Colunas');
    checkCOL.value = false;
    checkTBL = window.add("Checkbox", undefined, 'Tabela -> Texto');
    checkTBL.value = false;

    var btn = window.add("Button", undefined);
    btn.text = 'Scripts Atualizados';
    btn.onClick = function () { 
        copiarScripts();
        btn.text = "Scripts Atualizados";
        btn.enabled = false;
    };
    btn.graphics.font = font;
    btn.enabled = false;


    window.show();
    
    
    iniciarTarefa(btn);
}



function sincronizar(src, dst, type) {
    var srcFiles = src.getFiles(type);
    for(var f = 0; f < srcFiles.length; f++) {
        var srcDate = srcFiles[f].modified;
        var dstFile = dst.getFiles(srcFiles[f].name);
        if (dstFile.length > 0) {
            if (dstFile[0].modified < srcDate) {
                if (Window.confirm('Atualizar '+srcFiles[f].name+' ?')) {
                    srcFiles[f].copy(dstFile[0]);
                }
            }
        } else {
            if (Window.confirm('Adicionar '+srcFiles[f].name+' ?')) {
                $.writeln( srcFiles[f].copy(new File(dst + '/' + srcFiles[f].name)) );
            }
        }
    }
}

function copiarScripts() {
    var src = new Folder('/Z/Scripts/2015_Skynet/');
    var dst = (new File($.fileName)).parent;
    
    sincronizar(src, dst, '*.jsx');
    
    src = src.getFiles('Tools')[0];
    dst = dst.getFiles('Tools')[0];
    
    sincronizar(src, dst, '*.*');
}

function verificarScripts() {
    var src = new Folder('/Z/Scripts/2015_Skynet/');
    var dst = (new File($.fileName)).parent;
    
    if (verificar(src, dst, '*.jsx')) 
        return true;
    
    src = src.getFiles('Tools')[0];
    dst = dst.getFiles('Tools')[0];
    
    //$.writeln(src.exists);
    //$.writeln(dst.exists);
    
    if (verificar(src, dst, '*.*')) 
        return true;
    
    return false;
}

function verificar(src, dst, type) {
    var srcFiles = src.getFiles(type);
    for(var f = 0; f < srcFiles.length; f++) {
        var srcDate = srcFiles[f].modified;
        var dstFile = dst.getFiles(srcFiles[f].name);
        if (dstFile.length > 0) {
            if (dstFile[0].modified < srcDate) {
                //$.writeln('Modificado: '+srcFiles[f].fsName);
                return true;
            }
        } else {
            //$.writeln('Novo: '+srcFiles[f].fsName);
            return true;
        }
       //$.writeln(srcFiles[f].fsName);
    }
    return false;
};


function pararTarefa() {
    var task = app.idleTasks.itemByName("verificarAtualizacoes");
    if (task != null)
    {
        task.remove();
    }
}

function iniciarTarefa(btn) {
    var task = app.idleTasks.add({name:"verificarAtualizacoes", sleep:300000}); // 5 minutos

    var checkUpdate = function() {
        
        if (verificarScripts()) {
            btn.enabled = true;
            btn.text = "Atualizar Scripts";
        }
    };

    task.addEventListener(IdleEvent.ON_IDLE, checkUpdate);
}


pararTarefa();
criarJanela();