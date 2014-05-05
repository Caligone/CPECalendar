var _c = require('./config').config;
var XLSX = require('xlsx');
var _l = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L',]
console.log('===============================');
console.log('You are using CPECalendar v'+_c.version);
console.log('===============================');

console.log('=> Reading the file '+_c.path);

// Get the calendar
var workbook = XLSX.readFile(_c.path);

// Get the sheet
if(workbook.SheetNames.indexOf(_c.worksheetName) === -1) {
  console.error('Sheet "'+_c.worksheetName+'" not found. Change the value in config.js file.');
  process.exit(-1);
}
var sheet = workbook.Sheets[_c.worksheetName];


var x = null, y = null, v = null, lastDate = null, values = [];
var events = {};
for (z in sheet) {
  if(z[0] === '!') continue;
  x = z.match(/\D+/).toString();
  y = z.match(/\d+/);
  v = sheet[z].w;

  if(x.length > 1) continue;

  // Cell is date
  if(v.match(/\d{1,2}\/\d{1,2}/)) {

    // TODO

    lastDate = v;
  }
  // Cell is not date
  else {
    values.push(sheet[z].w);
  }
  //console.log('['+x+','+y+'] '+v);
}

console.log(events);

// Parse the calendar

// Connect to Google Calendar

// Create Calendar if not exists

// Create events


// Applause
