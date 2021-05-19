//2012-06-12
//linha 40 - mudado de tabela geral para tabela uma coluna
//
//2012-06-11
//Script Tabela1Coluna.jsx - (21) app.selection[0].appliedTableStyle="tabela uma coluna";
// Deixa o fundo Branco facilitanto a pintura da linha, podendo atravessar a tabela (crtl +  shit + [ )
//
//

//(10)
if(app.documents.length != 0){
	if(app.selection.length != 0){
		switch(app.selection[0].constructor.name){
			//When a row, a column, or a range of cells is selected,
			//the type returned is "Cell"
			case "Cell":
			alert("A cell is selected.");
			break;
			case "Table":
			alert("A table is selected.");
			app.selection[0].appliedTableStyle="tabela uma coluna";
			app.selection[0].width="194mm";
			//app.selection[0].fillColor = "Blue";
			//app.selection[0].fillColor = "Paper";
			//app.selection[0].Contents.appliedParagraphStyle = "texto tabela";
			
			break;
			case "InsertionPoint":
			case "Character":
			case "Word":
			case "TextStyleRange":
			case "Line":
			case "Paragraph":
			case "TextColumn":
			case "Text":
				var total=app.selection[0].texts[0].tables.length;
				var count = 0;
				while(total!=count){
						var table = app.selection[0].texts[0].tables[count];
						table.appliedTableStyle="tabela uma coluna";					
						table.width="194mm";						
						count++;
				}
			
				alert("feito");
			
			if(app.selection[0].parent.constructor.name == "Cell"){
			alert("The selection is inside a table cell.");
			}
			break;
			case "Rectangle":
			case "Oval":
			case "Polygon":
			case "GraphicLine":
			if(app.selection[0].parent.parent.constructor.name == "Cell"){
			alert("The selection is inside a table cell.");
			}
			break;
			case "Image":
			case "PDF":
			case "EPS":
			if(app.selection[0].parent.parent.parent.constructor.name == "Cell"){
			alert("The selection is inside a table cell.");
			}
			break;
			default:
			alert("The selection is not inside a table.");
			break;
		}
	}
}