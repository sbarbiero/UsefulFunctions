var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var filename = 'C:\\automate\\script docs\\pivotTemplate.xlsx'
var colName
var spend
var spendCopy



workbook.xlsx.readFile(filename)
  .then(function() {
    var worksheet = workbook.getWorksheet(i);
    // assigns the month column to a variable
    var monCol = worksheet.getColumn('A');
      // iterates over each cell in the month column
      // can I make this skip headers?
      monCol.eachCell(function(cell, rowNumber) {
        if (cell == 1) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('D1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Jan)
            spendCopy = worksheet.getCell('D' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017){
            colName = worksheet.getCell('P1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Jan)
            spendCopy = worksheet.getCell('P' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if (cell == 2) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016){
            colName = worksheet.getCell('E1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Feb)
            spendCopy = worksheet.getCell('E' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('Q1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Feb)
            spendCopy = worksheet.getCell('Q' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 3) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('F1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Mar)
            spendCopy = worksheet.getCell('F' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('R1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Mar)
            spendCopy = worksheet.getCell('R' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }
        } else if(cell == 4) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('G1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Apr)
            spendCopy = worksheet.getCell('G' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('S1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Apr)
            spendCopy = worksheet.getCell('S' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 5) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('H1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (May)
            spendCopy = worksheet.getCell('H' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('T1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (May)
            spendCopy = worksheet.getCell('T' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 6) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('I1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Jun)
            spendCopy = worksheet.getCell('I' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('U1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Jun)
            spendCopy = worksheet.getCell('U' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 7) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('J1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Jul)
            spendCopy = worksheet.getCell('J' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('V1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Jul)
            spendCopy = worksheet.getCell('V' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 8) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('K1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Aug)
            spendCopy = worksheet.getCell('K' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('W1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Aug)
            spendCopy = worksheet.getCell('W' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 9) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('L1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Sep)
            spendCopy = worksheet.getCell('L' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('X1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Sep)
            spendCopy = worksheet.getCell('X' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 10) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('M1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Oct)
            spendCopy = worksheet.getCell('M' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('Y1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Oct)
            spendCopy = worksheet.getCell('Y' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 11) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('N1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Nov)
            spendCopy = worksheet.getCell('N' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('Z1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Nov)
            spendCopy = worksheet.getCell('Z' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else if(cell == 12) {
          year = worksheet.getCell('B' + rowNumber).value
          if (year == 2016) {
            colName = worksheet.getCell('O1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Dec)
            spendCopy = worksheet.getCell('O' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else if (year == 2017) {
            colName = worksheet.getCell('AA1');
          // gets value out of spend for that row
            spend = worksheet.getCell('C' + rowNumber).value;
          // assigns spend value to appropriate column (Dec)
            spendCopy = worksheet.getCell('AA' + rowNumber);
            spendCopy.value = new Number(spend);
            console.log('Row ' + rowNumber + ' has a cell value of ' + cell + ' and has been modified. New value of ' + spend + ' assigned to column ' + colName);
          } else {
              console.log('Sorry, not valid year');
          }

        } else {
            console.log('Sorry, not valid data');
      }
  });
  return workbook.xlsx.writeFile('C:\\automate\\script docs\\pivotTemplateResult.xlsx');
});
