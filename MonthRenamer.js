var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var filename = 'C:\\automate\\monthRenamerTemplate.xlsx'
var i = 1
var monNum
var monName
var monNameNew


workbook.xlsx.readFile(filename)
  .then(function() {
    var worksheet = workbook.getWorksheet(i);
    monNum = worksheet.getColumn('A');

    monNum.eachCell(function(cell, rowNumber) {
      var logMessage = 'Row ' + rowNumber + ' is ' + monName;
      monNameNew = worksheet.getCell('B' + rowNumber);
      if (cell == 1) {
        monName = 'Jan';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 2) {
        monName = 'Feb';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 3) {
        monName = 'Mar';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 4) {
        monName = 'Apr';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 5) {
        monName = 'May';
        monNameNew.value = monName;
        console.log(logMessage);
       }
      else if (cell == 6) {
        monName = 'Jun';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 7) {
        monName = 'Jul';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 8) {
        monName = 'Aug';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 9) {
        monName = 'Sep';
        monNameNew.value = monName;
        console.log(logMessage);
       }
      else if (cell == 10) {
        monName = 'Oct';
        monNameNew.value = monName;
        console.log(logMessage);
      }
      else if (cell == 11) {
        monName = 'Nov';
        monNameNew.value = monName;
        console.log(logMessage);
       }
      else if (cell == 12) {
        monName = 'Dec';
        monNameNew.value = monName;
        console.log(logMessage);
       }
      else {
         console.log('Sorry, not valid data');
       }
    });
    return workbook.xlsx.writeFile('C:\\automate\\monthRenamerResult.xlsx');
  });
