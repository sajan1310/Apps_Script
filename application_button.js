function openApplication() {
  // Your deployed Apps Script Web App URL
  var url = "https://script.google.com/macros/s/AKfycbwjIv6koQT8TcQ6vnBhdZPhAmmEkPo65uPnfep4hODKTSLUlIlFJi2IcR3jbimDTvy-/exec";

  // HTML payload to open the URL in a new tab and immediately close the dialog
  var html = "<script>window.open('" + url + "','_blank');google.script.host.close();</script>";

  // Create the modal window
  var userInterface = HtmlService.createHtmlOutput(html)
    .setWidth(90)
    .setHeight(1);

  // Display the dialog in the active spreadsheet
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Opening Application...");
}