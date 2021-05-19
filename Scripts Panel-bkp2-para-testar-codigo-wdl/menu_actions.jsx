// Show menu actions and their ids
// Peter Kahrel -- www.kahrel.plus.com
// See also http://kasyan.ho.com.ua/open_menu_item.html

// Example: app.menuActions.item('$ID/Find/Change...').invoke();


	#targetengine session;

(function () {

	function thousand_sep (n) {
		return String (n).replace (/(\d)(?=(\d\d\d)+$)/g, "$1,")
	}


	function ignore (action)
		{
			return false;
//~ 		return (Number(action.id) >= 57603 && Number(action.id) <= 60956)  // fonts and style names
//~ 					|| action.name.indexOf ('.indd') >= 0
//~ 					|| action.area.indexOf ('.js') >= 0
//~ 					|| action.area == 'Text Selection'
//~ 					|| action.area.indexOf ('Menu:Insert') > -1 ;
		}


	function sort_object (list, key)
	{
		var temp = [], 
			temp2 = {};

		for (var i in list) {
			temp.push ({name: i, area: list[i].area, id: list[i].id});
		}
		temp.sort (function (a,b) {return a[key] > b[key]});
		for (var i = 0; i < temp.length; i++) {
			temp2[temp[i].name] = {area: temp[i].area, id: temp[i].id};
		}
		return temp2;
	}


	function sort_multi_column_list (list, key)
	{
		var temp = [], 
			L = list.items.length;
			
		for (var i = 0; i < L; i++) {
			temp.push ({name: list.items[i].text, area: list.items[i].subItems[0].text, id: list.items[i].subItems[1].text});
		}
		if (key == 'id') {
			temp.sort (function (a,b) {return Number (a[key]) > Number (b[key])});
		} else {
			temp.sort (function (a,b) {return a[key] > b[key]});
		}

		list.removeAll();
		for (var i = 0; i < L; i++)
		{
			with (list.add ('item', temp[i].name))
			{
				subItems[0].text = temp[i].area;
				subItems[1].text = temp[i].id;
			}
		}
	} // sort_multi_column_list

	//-------------------------------------------------------------------------------------
	// We create two objects here: a key-value object for the dialog's listbox
	// and an array for the Area dropdown.

	function create_object ()
	{
		var list = {}, 
			areas = [], 
			known = {}, 
			actions = app.menuActions.everyItem().getElements();
			
		for(var i = 0; i < actions.length; i++)
		{
			if (!ignore (actions[i]))
			{
				list[actions[i].name] = {area: actions[i].area, id: actions[i].id};
				if (!known[actions[i].area])
				{
					areas.push (actions[i].area);
					known[actions[i].area] = true;
				}
			} // ignore
		}
		// Initially sort by area
		list = sort_object (list, "area");
		areas.sort();
		areas.unshift('[All]'); // add [All] at the beginning of the array
		return {list: list, areas: areas}
	}



	//---------------------------------------------------------------------------------------------------------------
	// The interface
	var action_object = create_object ();
	
	var column_widths = [250, 150, 150];
	var w = new Window ("palette {text: 'Menu actions', orientation: 'row', alignChildren: 'top'}");
	var total = 0;
	var dummy = w.add ('group');
	var list = dummy.add ("listbox", undefined, "", {multiselect: true, 
																	numberOfColumns: 3, 
																	showHeaders: true, 
																	columnTitles: ["Name", "Area", "ID"], 
																	columnWidths: column_widths});
	
	list.maximumSize.height = w.maximumSize.height-100;
//~ 		list.preferredSize.width = 800
//~ 		list.preferredSize.height = w.maximumSize.height-100;
		
		var filter = w.add ('group {orientation: "column", alignChildren: "right"}');
			var name_group = filter.add ('group');
				name_group.add ('statictext {text: "Name:"}');
				var search_name = name_group.add ('edittext {active: true}');
				
			var keystring_group = filter.add ('group');
				keystring_group.add ('statictext {text: "Keystring:"}');
				var keystring_name = keystring_group.add ('edittext');

			var area_group = filter.add ('group');
				area_group.add ('statictext {text: "Area:"}');
				var area_dropdown = area_group.add ('dropdownlist', undefined, action_object.areas);
					area_dropdown.selection = 0;

			var id_group = filter.add ('group');
				id_group.add ('statictext {text: "ID:"}');
				var search_id = id_group.add ('edittext');
			
			var sort_group = filter.add ('group');
				sort_group.add ('statictext {text: "Sort:"}');
				var sort_dropdown = sort_group.add ('dropdownlist', undefined, ['Name', 'Area', 'ID']);
					sort_dropdown.selection = 1;
				
			for (var i = filter.children.length-1; i >= 0; i--) {
				filter.children[i].children[1].preferredSize.width = 150;
			}
		// Populate the list
		
		for (var i in action_object.list)
		{
			with (list.add ('item', i))
			{
				subItems[0].text = action_object.list[i].area;
				subItems[1].text = action_object.list[i].id;
			}
			total++
		}
		w.text += "Menu actions ("+thousand_sep(total)+" items)";


		search_name.onChange = function()
			{
			total = 0;
			var newlist = dummy.add ("listbox", list.bounds, "", {multiselect: true, 
																		numberOfColumns: 3, 
																		showHeaders: true, 
																		columnTitles: ["Name", "Area", "ID"], 
																		columnWidths: column_widths});

			for (var i in action_object.list)
			{
				if (i.search(search_name.text) > -1)
				{
					with (newlist.add ('item', i))
					{
						subItems[0].text = action_object.list[i].area;
						subItems[1].text = action_object.list[i].id;
					}
					total++;
				}
			}
			dummy.remove(list);
			list = newlist;
			w.text = "Menu actions ("+thousand_sep(total)+" items)";
		}

		area_dropdown.onChange = function ()
			{
			total = 0;
			var newlist = dummy.add ("listbox", list.bounds, "", {multiselect: true, 
																		numberOfColumns: 3, 
																		showHeaders: true, 
																		columnTitles: ["Name", "Area", "ID"], 
																		columnWidths: column_widths});

			for (var i in action_object.list)
				{
				if (action_object.list[i].area.indexOf(area_dropdown.selection.text) == 0 || area_dropdown.selection.text == '[All]')
					{
					with (newlist.add ('item', i))
						{
						subItems[0].text = action_object.list[i].area;
						subItems[1].text = action_object.list[i].id;
						}
					total++;
					}
				}
			dummy.remove(list);
			list = newlist;
			w.text = "Menu actions  ("+thousand_sep(total)+" items)";
			}
		
		
		search_id.onChange = function ()
			{
			list.selection = null;
			var L = list.items.length;
			for (var i = 0; i < L; i++)
				{
				if (list.items[i].subItems[1].text == search_id.text)
					{
					list.selection = i;
					break;
					}
				}
			}
		
		list.onDoubleClick = function () {
			//$.bp()
			search_name.text = list.selection[0].text + "|" + action_object.list[list.selection[0].text].id;
			keystring_name.text = app.findKeyStrings(list.selection[0].text).join (' >> ');
		}

		sort_dropdown.onChange = function ()
			{
			var new_list = sort_multi_column_list (list, sort_dropdown.selection.text.toLowerCase());
			}
		
	w.show ();

}());