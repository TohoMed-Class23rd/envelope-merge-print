function exportPDF() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("差込印刷");
    const ssID = ss.getId();
    const shID = sheet.getSheetId();
    const parentFolders = DriveApp.getFileById(ss.getId()).getParents();
    const folder = parentFolders.next().getFoldersByName('PDF').next();
    let baseUrl = "https://docs.google.com/spreadsheets/d/"
        + ssID
        + "/export?gid="
        + shID;
    let pdfOptions = "&exportFormat=pdf&format=pdf"
        + "&size=A4"
        + '&portrait=false'
        + "&fitw=true"
        + "&top_margin=0.1"
        + "&bottom_margin=0.1"
        + "&left_margin=0.1"
        + "&right_margin=0.1"
        + "&horizontal_alignment=LEFT"
        + "&vertical_alignment=TOP"
        + "&gridlines=false";
    let url = baseUrl + pdfOptions;
    
    for (let groupNum of ([...Array(14)].map((_, i) => i + 1))) {
        let token = ScriptApp.getOAuthToken();
        let options = {
          headers: {
              'Authorization': 'Bearer ' +  token
          },
          muteHttpExceptions : true
        };
        sheet.getRange("A1").setValue(groupNum);
        SpreadsheetApp.flush();
        let blob = UrlFetchApp.fetch(url, options).getBlob().setName(groupNum + '.pdf');
        folder.createFile(blob);
        console.log("Printed page"+ String(groupNum))
        Utilities.sleep(6000);
    }
}
