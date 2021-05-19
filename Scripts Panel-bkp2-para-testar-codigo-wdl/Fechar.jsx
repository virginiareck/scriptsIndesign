// var showMenuItem = app.menus.itemByName("Exibir").submenus.itemByName("Menos zoom");  
//     showMenuItem.associatedMenuAction.invoke();  
    app.menuActions.item("$ID/UpdateTableOfContentsCmd").invoke();
    app.menuActions.item("$ID/UpdateTableOfContentsCmd").invoke("Codatos");
    // Sumário...|
    app.menuActions.item("71426").invoke();

// var hideMenuItem = app.menus.itemByName("$ID/Main").submenus.itemByName("$ID/Object").menuItems.itemByName("$ID/Hide");  
// if(hideMenuItem.enabled)   
//    hideMenuItem.associatedMenuAction.invoke(); 