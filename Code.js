  class reportingObject  {
      constructor (institution,name,room,phone,dateSeen,noVisitPlease,priestComments){
        this.institution = institution;
        this.name = name;
        this.room = room;
        this.phone = phone;
        this.noVisitPlease = noVisitPlease;
        this.dateSeen = dateSeen;
        this.priestComments = priestComments;
      }
    }
//=sort(A2:E159,A2:A159, TRUE, E2:E159, TRUE)
//This function will get the History Table back into sort order
  function _sum(cellToSumIncoming, cellToSet){
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var totalSumOfAllResidentsAtAllFacilities = 0;
    var totalSumOfAllEucharistsGivenAtAllFacilities = 0;
    var totalSumOfAllAnointingsGivenAtAllFacilities = 0;
    var totalSumOfAllApostolicPardonsGivenAtAllFacilities = 0;
    var totalSumOfAllBlessingsAtAllFacilities = 0;
   
    var sumOfResidentsAtFacility = 0;
    var sumOfEucharistAtFacility = 0;
    var sumOfAnointingAtFacility = 0;
    var sumOfApostolicPardonsAtFacility = 0;
    var sumOfBlessingsAtFacility = 0;
    
    var cellToSum = "E1"
    if(cellToSumIncoming != undefined) cellToSum = cellToSumIncoming
    if(cellToSet == undefined) cellToSet = "A1"
    var reportSheet ;
    var historySheet;
    var historySheetByDateSeenOverall;
    // first sheet is the report 
    var facilityCount = 0;
    var facilitiesWithNoResidentsToVisit = 0;
    var reportingObjectArray = [];
    var ifHomeBoundSheet = false;
    const homeBoundOffsetForAS = 3;
    const homeBoundOffsetForNoVisit = 3;
    for(var i = 0; i <= sheets.length-1; i++){
      var sheetName = sheets[i].getName();
      if(sheetName == "Schedule For Pastoral Visits to Nursing Homes and the Homebound"){
        continue;
      }
      //Schedule for Pastoral Visits to Nursing Home
      if(sheetName == "Report"){ 
        reportSheet = sheets[i];
        continue;
      }
      if(sheetName == "History"){ 
        historySheet = sheets[i];
        continue;
      }
      if(sheetName == "History by Last Seen Overall"){ 
        historySheetByDateSeenOverall = sheets[i];
        continue;
      }

      if(sheetName == "Eucharistic Ministers"){ 
        continue;
      }
      if(sheetName == "Facility/EM Visit Schedules"){ 
        continue;
      }
      if(sheetName == "Pastoral Care Visit Report"){ 
        continue;
      }
      if(sheetName == "Homebound"){ 
        ifHomeBoundSheet = true;
      }
      
      try{
        sumOfResidentsAtFacility = 0;
        sumOfEucharistAtFacility = 0;
        sumOfAnointingAtFacility = 0;
        sumOfEucharistAtFacility = 0;
        sumOfBlessingsAtFacility = 0;
        
        // the operation below could be just simply added to the totalSums below without
        // this temporary storage, but it helps debugging and reporting
        sumOfResidentsAtFacility = parseInt(sheets[i].getRange("E1").getValue());
        sumOfEucharistAtFacility = parseInt(sheets[i].getRange("F1").getValue());
        sumOfAnointingAtFacility = parseInt(sheets[i].getRange("G1").getValue());
        sumOfApostolicPardonsAtFacility = parseInt(sheets[i].getRange("H1").getValue());
        sumOfBlessingsAtFacility = sheets[i].getRange("I1").getValue();
  
        
        
        console.log("Sheet name " + sheetName + " sumOfResidentsAtFacility " + sumOfResidentsAtFacility)
        console.log("Sheet name " + sheetName + " sumOfEucharistAtFacility " + sumOfEucharistAtFacility)
        console.log("Sheet name " + sheetName + " sumOfAnointingAtFacility " + sumOfAnointingAtFacility)
        console.log("Sheet name " + sheetName + " sumOfApostolicPardonsAtFacility " + sumOfApostolicPardonsAtFacility)
        console.log("Sheet name " + sheetName + " sumOfBlessingsAtFacility " + sumOfBlessingsAtFacility)
        
        
        if(sumOfResidentsAtFacility == 0){
          facilitiesWithNoResidentsToVisit++;
          continue;
        } 
        facilityCount++;
        totalSumOfAllResidentsAtAllFacilities +=  sumOfResidentsAtFacility
        totalSumOfAllEucharistsGivenAtAllFacilities +=  sumOfEucharistAtFacility;
        totalSumOfAllAnointingsGivenAtAllFacilities +=  sumOfAnointingAtFacility;
        totalSumOfAllApostolicPardonsGivenAtAllFacilities +=  sumOfApostolicPardonsAtFacility;
        totalSumOfAllBlessingsAtAllFacilities +=  sumOfBlessingsAtFacility;
        // find the row at which to start looking at anointing and AP
        // do this by looking for the row with the string "anointing of the Sick (Date only)"
        // in column G and and 1 to the row do one more until the lastrow 
        var foundDesiredRow = false;
        for(row = 1;row <= sheets[i].getLastRow();row++){
   
          var cellValue = !ifHomeBoundSheet? sheets[i].getRange(row,7).getValue(): sheets[i].getRange(row,10).getValue();
          if(cellValue == "Anointing of the Sick (Date only)"){
            foundDesiredRow = true;
            continue;// skip to the next row
          }
          if(foundDesiredRow == true){
            /* this.institution = institution;
                this.name = name;
                this.room = room;
                this.phone = phone;
                this.noVisitPlease = noVisitPlease;
                this.dateSeen = dateSeen;
            */
              var sheetReportingObject = new reportingObject();

              sheetReportingObject.institution = sheets[i].getName();
              
              sheetReportingObject.name = 
              !ifHomeBoundSheet ? sheets[i].getRange(row,1).getValue():
                              sheets[i].getRange(row,1).getValue() + " " + 
                              sheets[i].getRange(row,2).getValue();

              sheetReportingObject.room = 
              !ifHomeBoundSheet ? sheets[i].getRange(row,2).getValue(): 
                                  sheets[i].getRange(row,4).getValue(); 
              
              sheetReportingObject.phone = 
              !ifHomeBoundSheet ? sheets[i].getRange(row,3).getValue(): 
                                  sheets[i].getRange(row,3).getValue();
              
              sheetReportingObject.noVisitPlease  =  
              !ifHomeBoundSheet ? sheets[i].getRange(row,10).getValue(): 
                                  sheets[i].getRange(row,12).getValue();

              sheetReportingObject.priestComments  =  
              !ifHomeBoundSheet ? sheets[i].getRange(row,5).getValue(): 
                                  sheets[i].getRange(row,8).getValue();

              sheetReportingObject.dateSeen  = findLastDateSeen(sheets[i],row,ifHomeBoundSheet) 
              reportingObjectArray.push(sheetReportingObject);
          }
        }
        if(ifHomeBoundSheet)  ifHomeBoundSheet = false; 
      }catch (err) {
          console.log("Error when accessing sheet: " + sheets[i].getName() + " : " + err);
      }
 
    }
  

    var zeroResidents = (facilitiesWithNoResidentsToVisit > 0) ? "plus " + facilitiesWithNoResidentsToVisit + " facility/facilities that have no residents to see at this time":"" ;
    var outputStringToReturn =Utilities.formatString("Count of all Parishioners at %s facilities = %s %s as of %s",sheets.length,totalSumOfAllResidentsAtAllFacilities,zeroResidents, Utilities.formatDate(new Date(),"America/New_York","MM-dd-yyyy  HH:mm"));
    console.log(outputStringToReturn);
    var outputString =Utilities.formatString("Count of all Parishioners at %s facilities",facilityCount)
    
    var row = 1;
    reportSheet.getRange(row,1).setValue(outputString);
    reportSheet.getRange(row,2).setValue(totalSumOfAllResidentsAtAllFacilities);
    reportSheet.getRange(row,3).setValue(Utilities.formatDate(new Date(),"America/New_York","MM-dd-yyyy  HH:mm"));
    row++;
    zeroResidents =  "Facilities that don't have any residents to see at this time";
    reportSheet.getRange(row,1).setValue(zeroResidents);
    reportSheet.getRange(row,2).setValue(facilitiesWithNoResidentsToVisit);
    row++
    reportSheet.getRange(row,1).setValue("Total Sum Of All Eucharists Given At All Facilities");
    reportSheet.getRange(row,2).setValue(totalSumOfAllEucharistsGivenAtAllFacilities);
    row++

    reportSheet.getRange(row,1).setValue("Total Sum Of All Anointings Given At All Facilities");
    reportSheet.getRange(row,2).setValue(totalSumOfAllAnointingsGivenAtAllFacilities);
    row++
    reportSheet.getRange(row,1).setValue("Total Sum Of All Apostolic Pardons Given At All Facilities");
    reportSheet.getRange(row,2).setValue(totalSumOfAllApostolicPardonsGivenAtAllFacilities);
    row++
    reportSheet.getRange(row,1).setValue("Total Sum Of All Blessings Given At All Facilities");
    reportSheet.getRange(row,2).setValue(totalSumOfAllBlessingsAtAllFacilities);
    
    
    reportingObjectArray.sort((a, b) => {
       // Compare institutions first
       if (a.institution < b.institution) return -1;
       if (a.institution > b.institution) return 1;
  
        // Handle blank dates by treating them as earliest or latest
       const dateA = a.dateSeen ? new Date(a.dateSeen) : new Date(0); // Blank -> Epoch (earliest)
       const dateB = b.dateSeen ? new Date(b.dateSeen) : new Date(0); // Blank -> Epoch (earliest)
       // If institutions are the same, compare dates
      return dateA- dateB;
    });
    runHistoryReports(historySheet,reportingObjectArray);
    // console.log(reportingObjectArray);
    reportingObjectArray.sort((a, b) => {
      // Handle missing dateSeen values (treat as highest priority)
      const dateA = a.dateSeen ? new Date(a.dateSeen) : new Date(0); // Epoch date (earliest)
      const dateB = b.dateSeen ? new Date(b.dateSeen) : new Date(0);
    // Sort by dateSeen (earliest date first)
      const dateComparison = dateA - dateB;
      if (dateComparison !== 0) return dateComparison;
    //If dateSeen is the same or missing, sort alphabetically by institution
      return a.institution.localeCompare(b.institution);
    });
    runHistoryReports(historySheetByDateSeenOverall,reportingObjectArray);
    
    return outputStringToReturn;
  }
  function runHistoryReports(sheet,reportingArray){
    sheet.clear();
    var rowToStart = 1; //historySheet.getLastRow()+1;
    sheet.getRange(rowToStart,1).setValue("Institution");
    sheet.getRange(rowToStart,2).setValue("Name");
    sheet.getRange(rowToStart,3).setValue("Room");
    sheet.getRange(rowToStart,4).setValue("Phone");
    sheet.getRange(rowToStart,5).setValue("Date Last Seen");
    sheet.getRange(rowToStart,6).setValue("Priest Comments"); //History from institutional sheets
    
    rowToStart++
    sheet.getRange(1, 3, reportingArray.length).setNumberFormat("@STRING@");
    for(reportingRow = 0; reportingRow < reportingArray.length;reportingRow++){
      // does not wish to be seen 
        if(reportingArray[reportingRow].noVisitPlease == "X")
          continue;
        // no room means that either:
        // its an institution section name 
        // or the resident has no room # 
        // which makes it difficult to see them :)
        if(!(reportingArray[reportingRow].room) || /^\s*$/.test(reportingArray[reportingRow].room)){
          continue;
        }
        sheet.getRange(rowToStart,1).setValue(reportingArray[reportingRow].institution);
        sheet.getRange(rowToStart,2).setValue(reportingArray[reportingRow].name);
        sheet.getRange(rowToStart,3).setValue(reportingArray[reportingRow].room);
        sheet.getRange(rowToStart,4).setValue(reportingArray[reportingRow].phone);
        sheet.getRange(rowToStart,5).setValue(reportingArray[reportingRow].dateSeen);
        sheet.getRange(rowToStart,6).setValue(reportingArray[reportingRow].priestComments);
        
        rowToStart++;
    }

  }
  function findLastDateSeen(sheet,row,ifHomeBoundSheet){
    var dateLastSeenArray = [];
    for(column = 6; column <10;column++)
    {
        dateLastSeenArray.push(!ifHomeBoundSheet ? sheet.getRange(row,column).getValue(): 
                            sheet.getRange(row,column+3).getValue());
    }
    dateLastSeenArray.sort((a,b) => {
       // Handle blank dates by treating them as earliest or latest
      const dateA = a ? new Date(a) : new Date(0); // Blank -> Epoch (earliest)
      const dateB = b ? new Date(b) : new Date(0); // Blank -> Epoch (earliest)
  
      return dateB- dateA;
    });
    return dateLastSeenArray[0]// should have the latest date
  }
  function testOnEdit(){
    var beginRow = 14
    var endrow = 14
    var NUMBER_OF_TESTS = 1;
    "Colonnades"
    var dataRange = SpreadsheetApp.getActiveSheet().getDataRange()
    var sheetName = SpreadsheetApp.getActiveSheet().getName()
    var data = dataRange.getValues();
    var headers = data[0];
    // Start at row 1, skipping headers in row 0
    for (var row=beginRow; row <= endrow; row++) {
      var e = {};
      e.values = data[row]//.filter(Boolean);  
      e.range = dataRange.offset(row,0,1,data[0].length);
      e.namedValues = {};
      // Loop through headers to create namedValues object
      // NOTE: all namedValues are arrays.
      for (var col=0; col<headers.length; col++) {
        e.namedValues[headers[col]] = [data[row][col]];
      }
      // Pass the simulated event to onFormSubmit
      onEdit(e);
    }
  }
  function getDate() {
    var today = new Date();
    today.setDate(today.getDate());
    return Utilities.formatDate(today, 'GMT+03:00', 'dd/MM/yy');
  }
  function printRangeToPrinter() {
  var sSheet =  SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeToPrint = sheet.getRange("A1:E35"); // Change this to the desired range
  
  var printOptions = {
    'landscape': true,   // Set to true for landscape orientation
    'scale': 1,           // Scaling factor (1 = 100%)
    'fitToWidth': false,   // Fit to page width
    'fitToHeight': false   // Fit to page height
  };
  
  var url = sSheet.getUrl();
  var pdf = DriveApp.getFileById(sSheet.getId()).getAs('application/pdf').getBytes();
 
  var attach = { fileName: "PrintedRange.pdf", content: pdf, mimeType: 'application/pdf' };
  
  // Print the PDF to the default printer
  MailApp.sendEmail({
    to: '',  // Leave this empty
    subject: 'Printed Range',
    body: 'Please find the printed range attached.',
    attachments: [attach],
    printOptions: printOptions
  });
  
  // Delete the temporary spreadsheet
  newSpreadsheet.setTrashed(true);
}
function sortSheetsAsc() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNameArray = [];

  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }

  sheetNameArray.sort();

  for( var j = 0; j < sheets.length; j++ ) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j + 1);
  }
}

function sortSheetsDesc() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNameArray = [];

  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }

  sheetNameArray.sort().reverse();

  for( var j = 0; j < sheets.length; j++ ) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j + 1);
  }
}

function sortSheetsRandom() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNameArray = [];

  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }

  sheetNameArray.sort().sort(() => (Math.random() > .5) ? 1 : -1);;

  for( var j = 0; j < sheets.length; j++ ) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j);
  }
}

// Google Apps Script code for building a printable table
// from the user's selection (up to 4 columns), exporting to PDF,
// and emailing it to a recipient.

function buildAndEmailReportFromSelection() {
  const RECIPIENT = "tomh@incarnationparish.org";
  const REPORT_SHEET_NAME = "Pastoral Care Visit Report";
  const ROW_HEIGHT = 40; // slightly taller for handwriting in landscape
  const HEADER_ROW = [
    "Name",
    "Room",
    "Phone",
    "Last\nSeen",
    "Eucharist\nReceived",
    "Anointed",
    "Apostolic\nPardon",
    "Blessing\nOnly",
    "Comments",
  ];
  const INPUT_COLS = 4;   // from selection
  const TOTAL_COLS = 9;  // 4 selected + 5 extra columns (incl. checkboxes)

  const ss = SpreadsheetApp.getActive();
  const userSel = ss.getActiveRange();
  if (!userSel) {
    SpreadsheetApp.getUi().alert("Please select cells first (you can skip rows).");
    return;
  }

  // Gather values from all selected ranges (can be discontiguous)
  const rangeList = SpreadsheetApp.getActiveRangeList();
  const selRanges = rangeList ? rangeList.getRanges() : [userSel];

  const rows = [];
  selRanges.forEach(range => {
    const values = range.getValues(); // 2D array
    for (let r = 0; r < values.length; r++) {
      const rowVals = values[r];
      const take = [];
      for (let c = 0; c < INPUT_COLS; c++) {
        take.push(rowVals[c] !== undefined ? rowVals[c] : "");
      }
      rows.push(take);
    }
  });

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert("No values found in your selection.");
    return;
  }

  // Prepare (replace) the Report sheet
  const existing = ss.getSheetByName(REPORT_SHEET_NAME);
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet(REPORT_SHEET_NAME);

  // Row 1: empty
  // Row 2: header
  sheet.getRange(2, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);

  // Data rows start at row 3: compose 10 columns (4 data + 6 blanks)
  const output = rows.map(r => {
    const padded = r.slice(0, INPUT_COLS);
    while (padded.length < INPUT_COLS) padded.push("");
    while (padded.length < TOTAL_COLS) padded.push("");
    return padded;
  });

  sheet.getRange(3, 1, output.length, TOTAL_COLS).setValues(output);

  // Formatting for readability/printing
  const widths = [
    180, // Name
    120, // Room
    120, // Phone
    120, // Last Seen
    110, // Eucharist Received (checkbox)
    100, // Anointed (checkbox)
    120, // Apostolic Pardon (checkbox)
    130, // Blessing Only (checkbox)
    300 // Comments
  ];
  for (let i = 0; i < Math.min(widths.length, TOTAL_COLS); i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  // Set row height; for print legibility aim for width ≈ 3x height ratio
  // We keep a fixed height and reasonable column widths above.
  sheet.setRowHeightsForced(2, 1, ROW_HEIGHT);         // header row
  sheet.setRowHeightsForced(3, output.length, ROW_HEIGHT*4); // data rows
  sheet.setFrozenRows(2); // keep header visible

  const tableRange = sheet.getRange(2, 1, output.length + 1, TOTAL_COLS);
  tableRange.setBorder(true, true, true, true, true, true);
  // Fonts and alignment
  sheet.getRange(2, 1, 1, TOTAL_COLS)
       .setFontFamily("Arial")
       .setFontSize(14)
       .setFontWeight("bold")
       .setHorizontalAlignment("center");
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
  .setFontFamily("Arial")
  .setFontSize(14)
  .setFontWeight("bold")
  .setVerticalAlignment("middle");

  // Add checkboxes to sacramental columns for data rows
  // Columns: 5 Eucharist, 6 Anointed, 7 Apostolic Pardon, 9 Blessing Only
  if (output.length > 0) {
    sheet.getRange(3, 5, output.length, 1).insertCheckboxes();
    sheet.getRange(3, 6, output.length, 1).insertCheckboxes();
    sheet.getRange(3, 7, output.length, 1).insertCheckboxes();
    sheet.getRange(3, 8, output.length, 1).insertCheckboxes();
  }

  // Export this sheet to PDF and email
  const pdfBlob = exportSheetToPdfBlob_(ss, sheet, { repeatFrozenRows: true});
  MailApp.sendEmail({
    to: RECIPIENT,
    subject: "Pastoral Visits Report",
    body: "Attached is the printable visits report.",
    attachments: [pdfBlob],
  });

  SpreadsheetApp.getUi().alert("Report created and emailed.");
}

// Helper: export only the provided sheet as a PDF blob
function exportSheetToPdfBlob_(ss, sheet, opts) {
  const ssId = ss.getId();
  const gid = sheet.getSheetId();
  const repeatFrozen = opts && opts.repeatFrozenRows === true;

  const exportUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf" +
    "&portrait=false" +
    "&fitw=true" +                 // fit to width
    "&gridlines=true" +            // show gridlines
    "&printtitle=false" +
    "&sheetnames=false" +
    "&pagenum=FOOTER" +
    "&fzr=" + (repeatFrozen ? "true" : "false") + // repeat frozen rows on each page
    "&top_margin=0.5" +
    "&bottom_margin=0.5" +
    "&left_margin=0.5" +
    "&right_margin=0.5" +
    "&horizontal_alignment=CENTER" +
    "&vertical_alignment=TOP" +
    "&gid=" + gid;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true,
  });
  return response.getBlob().setName(sheet.getName() + ".pdf");
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Refresh Residents Count and History', functionName: '_sum'},
    {name: 'Export Selected rows to PDF. Select the history sheet.',functionName:'buildAndEmailReportFromSelection'}
  ];
  spreadsheet.addMenu('Sheet Tools', menuItems);
  if (typeof addPastoralCareMenu_ === "function") {
    addPastoralCareMenu_();
  }
}

  
