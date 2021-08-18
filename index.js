const fs = require('fs');
const fsPromises = require('fs').promises;
const path = require('path');
const LineByLine = require('n-readlines');
const ExcelJS = require('exceljs');

// read alls files in given folders
async function readAllFiles(path) {
  const res = [];
  const files = await fsPromises.readdir(path);
  files.forEach((file) => {
    if (file.slice(-5) === '.java') {
      res.push(file);
    }
  });
  return res;
}

// read file line-by-line
function readContentFile(path) {
  const regex = /^(private\s\w+\s\w+;)|(public\s\w+\s\w+;)$/;
  const liner = new LineByLine(path);

  let line;

  const res = [];
  let row = [];
  let counter = 1;
  while ((line = liner.next())) {
    try {
      const text = line.toString('utf-8').trim();
      if (regex.test(text)) {
        row = text.slice(0, -1).split(/\s/);
        res.push([counter++, row[2], row[1], '']);
      }
    } catch (e) {}
  }
  return res;
}

function insertSheetAndData(workbook, sheetName, tabularData) {
  // create sheet
  const ws = workbook.addWorksheet(sheetName);
  // insert text with sheetName
  const cell = ws.getCell('A1');
  cell.value = sheetName;
  cell.font = { bold: true };

  // add a table to a sheet
  ws.addTable({
    name: sheetName + ' des',
    ref: 'A3',
    headerRow: true,
    columns: [
      { name: 'No' },
      { name: 'Field Name' },
      { name: 'Field Type' },
      { name: 'Description' },
    ],
    rows: tabularData,
  });
}

async function main() {
  const workbook = new ExcelJS.Workbook();

  try {
    const dirSpace = process.argv[2];
    const fileNames = await readAllFiles(dirSpace);
    for (let i = 0; i < fileNames.length; i++) {
      const sheetData = readContentFile(path.join(dirSpace, fileNames[i]));
      insertSheetAndData(workbook, fileNames[i].slice(0, -5), sheetData);
    }
  } catch (ex) {
    console.error('PK: ', ex);
  } finally {
    const stream = fs.createWriteStream(
      __dirname + '/data/' + (process.argv[3] || `output_${Date.now()}.xlsx`)
    );
    // write to a stream
    await workbook.xlsx.write(stream);
    // close stream
    stream.close();
    console.log('Done!!!');
  }
}

main();
