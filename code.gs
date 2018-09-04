function OnStart() {
dummy= setupvalidation();
}
function setupvalidation(){
 dummy = getvalidationlists()
 
}
function getvalidationlists() {
//Overview
// read list file and turn it into named ranges in the sheet
// set up the primary named ranges in the name field and all the level 1 "customer" cells
//
// Some delimiters, file names etc.
  var listdelim = "List:";
  var fielddelim = '||';
  var targetfile = 'ListDump.txt'; // file to find and use
  var targetfolder = 'Jim Dev';  // folder to work in
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssid = ss.getId();
  var ssfile= DriveApp.getFileByID(ssid);
  var targetfolders = ssfile.getParents();
  Logger.log (targetfolders[0]); 
  targetfolder = tragetfolders[0];
  var listsheet = ss.getSheetByName('Lists');
  if (listsheet ==null ) {
    listsheet = ss.insertSheet();
    ss.renameActiveSheet('Lists');
  } else {
    ss.setActiveSheet(listsheet);
  }
//
//  get the list data from the List Dump file
//
// for debug use the current folder for list data
// this will have to be replaced by user's folder
//
// clear current LISTS
var namedRanges = ss.getNamedRanges();
  if (namedRanges != null){
    for (var i = 0; i < namedRanges.length; i++) {
    onerange = namedRanges[i].getName();
    //Logger.log(onerange);
    if ( onerange.indexOf("LIST") >-1) {  // what happned to contains
      //Logger.log('Will remove range: ' + onerange); //
      namedRanges[i].remove();
    }
  }
}
// read latest lists from list dump file and set up as named ranges
var myfile = findafileinfolder(targetfile , targetfolder );
    if (myfile != null ){
      var alllines= myfile.getBlob().getDataAsString();
      //Logger.log (alllines);
      var lines = alllines.split('LIST:'); //("\\r?\\n|\\r"); //Zero or one returns followed by new line or just a return
      //Logger.log (lines.length);
      for (var i = 0; i< lines.length; i++)  {
        var mylistline= lines[i];
        if (mylistline.length >0 ) {
          //Logger.log (mylistline);
          var mylist = mylistline.split("||");
          Logger.log ('Range: ' + mylist[0] + ': 1, '+ (i+1) + ', ' +mylist.length);
          range = listsheet.getRange(1,i+1,mylist.length,1) ;  //a range for this list
          ss.setNamedRange( mylist[0]+ 'LIST',range); //mylist[0],range);
          //Logger.log (range.getHeight() + " " + range.getWidth() );
          //Logger.log (mylist);
          for (r=0;r<mylist.length;r++){
            cell = listsheet.getRange(r+1,i+1);
            cell.setValue (mylist[r]);
          }
          cell = listsheet.getRange(1,i+1);
          cell.setValue (mylist[0]);
          //Logger.log("Range Created: " + range.getName());
          }
      }
    }
  
}

function findafileinfolder(targetfile , targetfolder ) {
  ffile=null;
  var files = DriveApp.getFilesByName(targetfile); //There may be mulitple files with the same name
  while (files.hasNext()) {  //find the one that is in the right folder
    onefile=files.next();
//    Logger.log('Got A File: ' + onefile.getName() + onefile.getId());
    // Find the parents of this file
    parents=onefile.getParents();  
    while (parents.hasNext()) {
//      Logger.log("Has parents:-");
      var oneparent= parents.next();
      parentname = oneparent.getName();
//      Logger.log('Parent: ' + parentname);
      if (parentname == targetfolder ) {  //BEWARE double = is essential
        ffile=onefile;
//        Logger.log('Matched To This Parent: ' + parentname + targetfolder);
      }
    }
  }
return ffile;
}

function SendTimeSheet() {      
var ss= SpreadsheetApp.getActiveSpreadsheet();
Logger.log("Spreadsheet: " + ss.getName());
var sh= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Time Recording");  
Logger.log("Sheet: " + sh.getName());
//var franges =sh.getNamedRanges();
//var frange = sh.getRangeByName("PrefFileName");
  // The code below logs the name of the first named range.
var namedRanges = sh.getNamedRanges();
if (namedRanges.length > 1) {
  for (var i = 0; i< namedRanges.length; i++) {
    Logger.log(namedRanges[i].getName());
  }
}
// find completed time sheets folder
var myfolders = DriveApp.getFoldersByName("Jim Dev") ;
var mydir = "";
while ( myfolders.hasNext()) {
  var thisfolder = myfolders.next();
  var foldername =thisfolder.getName();
  if (foldername = "Jim Dev") { 
      myfolder=thisfolder;
      mydir = foldername;
      Logger.log("Found One: " + foldername);
   }
}
  var frange= ss.getRangeByName("PrefFileName"); //Preferred file name is on the sheet
  var PrefFile = frange.getValue()  ;//frange.getValues()
  if (mydir != "" ) {
  var outfile = "PrefFile";
  timesheetrange = ss.getRangeByName("TSData");
  var test = timesheetrange.getValues();
  outfile=convertRangeToCsvFile(timesheetrange);
  Logger.log(outfile)
  var file = myfolder.createFile(PrefFile+".csv", outfile);
  DriveApp.createFile(PrefFile+".csv", outfile);                                
}
}
function convertRangeToCsvFile(rangeToExport) {
  // Get from the spreadsheet the range to be exported 
  try {
    var dataToExport = rangeToExport.getValues();
    var csvFile = undefined;

    // Loop through the data in the range and build a string with the CSV data
    if (dataToExport.length > 1) {
      var csv = "";
      for (var row = 0; row < dataToExport.length; row++) {
        for (var col = 0; col < dataToExport[row].length; col++) {
          if (dataToExport[row][col].toString().indexOf(",") != -1) {
            //dataToExport[row][col] = "\"" + dataToExport[row][col] + "\"";
            dataToExport[row][col] = dataToExport[row][col];
          }
        }

        // Join each row's columns
        // Add a carriage return to end of each row, except for the last one
        if (row < dataToExport.length-1) {
          csv += dataToExport[row].join(",") + "\r\n";
        }
        else {
          csv += dataToExport[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
