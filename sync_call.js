function updateOnLoad () {
  
  SerialsUpdate.update('foreign', 'local');
  
  }
  
function updateOnCommand () { 
  
  SerialsUpdate.update('local', 'foreign');

} 

// Add menu item for update function  
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createMenu("Serials Data");//.createAddonMenu(); // Or DocumentApp or FormApp.
      // Add a normal menu item (works in all authorization modes).
  menu.addItem('Update Data Hub', 'updateOnCommand')
       .addItem('Get Latest Data', 'updateOnLoad')
       .addItem('Show Documentation', 'SerialsUpdate.showSidebar')
       .addToUi();
  
  updateOnLoad();
  
  SerialsUpdate.showSidebar();

}  
