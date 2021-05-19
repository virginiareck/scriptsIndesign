//**************************************************************************************************************
//**************************************************************************************************************
//****
//****	Page extractor 1.0 - A script from Loïc Aigon with the great help of Peter Kahrel and other guys
//****	that cooperated on this stuff in tha adobe scripting forum. Feel free to improve. 
//****	a little mail at loic _ aigon@yahoo.fr if you appreciate that script would be nice.
//****
//**************************************************************************************************************
//**************************************************************************************************************

//======================================
// <L10N> :: FRENCH_LOCALE :: RUSSIAN_LOCALE
//======================================
// Please save you file first before processing the script :: Enregistrez votre document avant de lancer le script :: Пожалуйста сохраните документ перед запуском скрипта
// Extract pages... :: Extraire des pages... :: Извлечь страницы...
// to :: a :: по
// Extract as separate pages :: Extraire chaque page comme un fichier :: Извлечь как отдельные страницы
// Remove pages after extraction :: Supprimer les pages après extraction :: Удалить страницы после извлечения
// Choose other extraction folder :: Choisir un autre dossier d'extraction :: Выбрать папку для сохранения
// </L10N> ::


// 2. Add the code's block below in your script:

if ( typeof __ == 'undefined' )
{
__sep = ' :: ';
__beg = ' <L10N>';
__end = ' </L10N>';

/*arr*/File.prototype.getCommentsContaining = function(/*str*/k_)
//--------------------------------------
	{
	var r = [];
	if ( this.open('r') )
		{
		var line;
		while(!this.eof)
			{
			line = this.readln();
			if ( ( line.substr(0,2) == '//' )
				&& ( line.indexOf(k_) >= 0 ) )
				r.push(line.substr(2));
			}
		this.close();
		}
	return(r);
	}
	
/*arr*/Array.prototype.parseL10N = function(/*str*/locale_)
//--------------------------------------
	{
	var r = [];
	var sz = this.length;
	var lm, ss;
	var st = -1
	var rg = 0;
	for ( var i=0 ; i < sz ; i++ )
		{
		ss = this[i].split(__sep, lm);
		if ( ( st == -1 ) && ( ss[0] == __beg ) )
			{
			lm = ss.length;
			for ( var j = 1 ; j < lm ; j++ )
				if ( ss[j] == locale_ ) rg=j;
			st = 0;
			continue;
			}
		if ( st == 0 )
			{
			if ( ( rg == 0 ) || ( ss[0] == __end ) ) break;
			if ( ss.length <= rg ) continue;
			r[ss[0].substr(1)] = ss[rg];
			}
		}
	return(r);
	}

/*str*/Number.prototype.toLocaleName = function()
//--------------------------------------
{
for ( var p in Locale )
	if ( Locale[p] == this ) return(p);
return(null);
}

String.L10N = File(app.activeScript).
	getCommentsContaining(__sep);
	// .
	// parseL10N(app.locale.toLocaleName());
	
function __(s){return(String.L10N[s]||s);}
}

// 1. Control for documents open. If true, the script launches the dialog

if(app.documents.length >0)
{
	var doc = app.activeDocument;
	if(doc.saved==true)
	{
		extractdlg();
	}
	else
	{
		alert(__("Salve seu arquivo antes de utilizar este script!!!"));
	}
}
else
{
	alert("Nenhum documento aberto !"); 
}

// 2. Gathers many infos on the document.

function pageinfos()
{
	var pg = doc.pages;
	var pglg  = pg.length;
	var pFirst = Number(pg[0].name);
	var pLast = Number(pg[pglg-1].name);
	var pgHeigth = doc.documentPreferences.pageHeight;
	var pgWitdh = doc.documentPreferences.pageWidth;
	var docname = String(doc.fullName).replace (/\.indd$/, "");
	var docpath = doc.filePath;
	var docfullname = doc.fullName;

	var infoarr = [pglg, pFirst, pLast, pgHeigth,pgWitdh,docname,docpath,docfullname];
	return infoarr;
}


// 3. Main function. First part is the dialog

function extractdlg()
{
	var docfile = String(pageinfos()[7]);
	var dlg = app.dialogs.add({name : "Extrator de páginas, versão 1.1 - ©www.loicaigon.com"}); 
	with(dlg)
	{
		var firstclmn = dialogColumns.add();
		with(firstclmn)
		{
			var firstrow = dialogRows.add();
			with(firstrow)
			{
				var clmn1 = dialogColumns.add();
				with(clmn1)
				{
					var row1 = dialogRows.add();
					row1.staticTexts.add({staticLabel : __("Extraindo páginas...")});
					var row2 = dialogRows.add();
					with(row2)
					{
						var r2c2 = dialogColumns.add();
						with(r2c2)
						{
							var r2c2r1 = dialogRows.add();
							var pgStart = r2c2r1.realEditboxes.add({editValue:pageinfos()[1], minWidth: 30});
						}
						var r2c3 = dialogColumns.add();
						with(r2c3)
						{
							var r2c3r1 = dialogRows.add();
							 r2c3r1.staticTexts.add({staticLabel : __("até")});
						}
						var r2c4 = dialogColumns.add();
						with(r2c4)
						{
							var r2c4r1 = r2c4.dialogRows.add();
							var pgEnd = r2c4r1.realEditboxes.add({editValue:pageinfos()[2], minWidth: 30});
						}
							
					}
				}
			}
			var secondrow = dialogRows.add();
			with(secondrow)
			{
				var clmn2 = dialogColumns.add();
				with(clmn2)
				{
					var row2 = dialogRows.add();
					with(row2)
					{
						var sepbox = checkboxControls.add({staticLabel:__("Extrair páginas separadas."), checkedState:false});
					}
				}
			}
			var thirdrow = dialogRows.add();
			with(thirdrow)
			{
				var clmn3 = dialogColumns.add();
				with(clmn3)
				{
					var row3 = dialogRows.add();
					with(row3)
					{
						var rembox = checkboxControls.add({staticLabel:__("Remover páginas após a extração"), checkedState:false});
					}
				}
			}
	
			var foutrhrow = dialogRows.add();
			with(foutrhrow)
			{
				var clmn4 = dialogColumns.add();
				with(clmn4)
				{
					var row4 = dialogRows.add();
					with(row4)
					{
						var savebox = checkboxControls.add({staticLabel:__("Escolher outro destino para os arquivos extraidos"), checkedState:false});
					}
				}
			}
			with(borderPanels.add()){
				staticTexts.add({staticLabel:"Município inicial:"});
				var myRadioButtonGroup = radiobuttonGroups.add();
				with(myRadioButtonGroup){
					var myLeftRadioButton = radiobuttonControls.add({staticLabel:"Left", checkedState:true});
					var myCenterRadioButton = radiobuttonControls.add({staticLabel:"Center"});
					var myRightRadioButton = radiobuttonControls.add({staticLabel:"Right"});
				}
			}
			
    var capa = doc.masterSpreads.itemByName('B-Capa');
    var frameCapa = ['Selecione o múnicipio']
    for (var f=0; f < capa.textFrames.length; f++) {
    	frameCapa.push(capa.textFrames[f].contents);
	}
		alert(capa.textFrames[0]);

			with(borderPanels.add()){
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:"Município inicial:"});
				}	
				with(dialogColumns.add()){
					//Create a pop-up menu ("dropdown") control.
					var myVerticalJustificationMenu = dropdowns.add({stringList:frameCapa, selectedIndex:0});
				}
			}
		}
	}

	// If the user made good choices, the script operates.
	
	if(dlg.show()==true)
	{
		if(pgStart.editValue >= pageinfos()[2] || pgEnd.editValue <= pageinfos()[1])
		{
			alert("The pages numbers may be at least "+pageinfos()[1] +" for the first page of the range and "+ pageinfos()[2] + " at maximum for the last page");
		}
		else
		{	
			// If the user choose to pick a different folder, he will be asked for. Otherwise, the dafault folder is the one containing the file.
			if(savebox.checkedState==true)
			{
				var extractfolder = Folder.selectDialog ("Please choose a folder where to save extracted pages...");
				if(!extractfolder)
				{
					exit();
				}
				else
				{
					var saveextractfolder = String(extractfolder.fullName)+"/" +String(doc.name).replace (/\.indd/, "");
				}
			}
			else
			{
				var saveextractfolder = String(pageinfos()[5]);
			}
			var rem0 = pageinfos()[0]-1;
			var rem2 =  (pgStart.editValue-2);
			
			// Variables definition regarding to the choice of the user to separate or not the extracted pages.
			
			if(sepbox.checkedState==true)
			{	
				var W = pgEnd.editValue-pgStart.editValue+1;
				var rem1 = pgStart.editValue;
			}
			else
			{
				var W = 1;
				var rem1 = pgEnd.editValue;
			}
			
			// Extraction loop 
			for(w=0; w<W;w++)
			{
				if(sepbox.checkedState==true)
				{
					var exportdocname = "_Pg" +(pgStart.editValue+w) +".indd";
				}
				else
				{
					var exportdocname = "_Pg"+pgStart.editValue+"_to_Pg_"+pgEnd.editValue +".indd";
				}
				for(var i=rem0; i>=rem1+w;i--)
				{
					doc.pages[i].remove();
				}
				for(var i=rem2+w; i>=0;i--)
				{
					doc.pages[i].remove();
				}
				var exportdoc = doc.save(File(saveextractfolder + exportdocname));
				exportdoc.close(SaveOptions.no);
				if(sepbox.checkedState==true && w<(pgEnd.editValue-pgStart.editValue))
				{
					app.open(File(docfile));
				}
			}
		
			// If the user chose to remove the extracted pages from the original document, it will re open the first document then remove the unuseful pages.
			if(rembox.checkedState == true)
			{
				app.open(File(docfile));
				for(var i=pgEnd.editValue-1; i>=pgStart.editValue-1;i--)
				{	
					doc.pages.everyItem().remove();
					doc.pages[i].remove();
				}
				app.activeDocument.close(SaveOptions.yes);
			}
		}
	}
}

			