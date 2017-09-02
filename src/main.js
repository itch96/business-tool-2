const electron = require("electron");
const { app, BrowserWindow, dialog } = electron;
const xlsx = require('xlsx');

let window = null;
let excelFile = null;
let photosDirectory = null;
let destinationDirectory = null;

app.on("ready", () => {
  console.log("App is ready.");
  window = new BrowserWindow({ width: 650, height: 400, backgroundColor: "#ffffff", show: false });
  window.setPosition(10, 40, false);

  window.once("ready-to-show", () => {
    window.show();
  });

  window.loadURL("http://localhost:3000");

  window.on("closed", () => {
    window = null;
  })
});

function openExcelFile() {
  const files = dialog.showOpenDialog(window, {
    properties: ['openFile'],
    filters: [
      {name: "Excel Files", extensions: ['xlsx', 'xls']}
    ]
  });

  if(!files) {return "No Excel File Chosen";}
  else {
    excelFile = files[0];
    return excelFile;
  }
}

function openPhotosDirectory() {
  const files = dialog.showOpenDialog(window, {
    properties: ['openDirectory'],
    filters: [
      {name: 'Photos Directory'}
    ]
  });

  if(!files) {return "No Photos Directory Chosen.";}
  else {
    photosDirectory = files[0];
    return photosDirectory;
  }
}

function openDestinationDirectory() {
  const files = dialog.showOpenDialog(window, {
    properties: ['openDirectory'],
    filters: [
      {name: 'Destination Directory'}
    ]
  });
  if(!files) {return "No Destination Directory Chosen."}
  else {
    destinationDirectory = files[0];
    return destinationDirectory;
  }
}

function doTheThing() {
  if(excelFile && photosDirectory && destinationDirectory) {
    let data = getExcelData();
  } else {
    return "Please Choose All 3 Sources";
  }
}

function getExcelData() {
  let workbook = xlsx.readFile(excelFile);
  let sheet_name = workbook.SheetNames[0];
  let worksheet = workbook.Sheets[sheet_name];
  let values = xlsx.utils.sheet_to_json(worksheet);
  return values;
}

exports.openExcelFile = openExcelFile;
exports.openPhotosDirectory = openPhotosDirectory;
exports.openDestinationDirectory = openDestinationDirectory;
exports.doTheThing = doTheThing;