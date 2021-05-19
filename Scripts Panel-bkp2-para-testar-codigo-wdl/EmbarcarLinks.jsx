var doc = app.activeDocument;

if (Window.confirm("Deseja embarcar os arquivos vinculados?")) {
    var embarcados = 0;

    for(var n = 0; n < doc.links.length; n++) {
        var link = doc.links[n];
        if (link.status == LinkStatus.NORMAL) {
            link.unlink();
            embarcados++;
        }
    }

    alert("Foram embarcados "+embarcados+" arquivos de um total de "+doc.links.length);
}

if (Window.confirm("Deseja corrigir vínculos perdidos?")) {

    var logosFolder = new Folder((new File($.fileName)).parent.fullName + '/Logos');
    var atosFolder = new Folder(doc.fullName.path + '/arquivos/');

    var corrigidos = 0, ausentes = 0;

    for(var n = 0; n < doc.links.length; n++) {
        var link = doc.links[n];
        if (link.status == LinkStatus.LINK_MISSING) {
            ausentes++;
            var f = atosFolder.getFiles(link.name);
            if (f.length != 0) {
                link.relink(f[0]);
                corrigidos++;
            } else {
                    f = logosFolder.getFiles(link.name);
                    if (f.length != 0) {
                        link.relink(f[0]);
                        corrigidos++;
                    }
            }
            
        }
    }
    if (ausentes == 0) {
        alert("Nenhum vínculo perdido.");
    } else {
        alert("Foram corrigidos "+corrigidos+" vínculos perdidos de um total de "+ausentes);
    }
}