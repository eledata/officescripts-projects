// Basic Office Scripts
// Office Scripts Github project
// URL Link: https://github.com/eledata/officescripts-projects
// 1. Read the cell value 
function main(workbook: ExcelScript.Workbook) {
  // 1. get the worksheet
  let test_worksheet = workbook.getWorksheet("test");
  if (!test_worksheet) {
    console.log("The script expects a worksheet named \"test\". ");
    return;
  }

  // 2. get the value of cell A1
  let a1_range = test_worksheet.getRange("A1");

  // 3. show the value on console
  console.log(a1_range.getValue());
  
  // set date and datetime value 
  let date = new Date(Date.now());
  let a2_range = test_worksheet.getRange("A2");
  let a3_range = test_worksheet.getRange("A3");
  a2_range.setValue(date.toLocaleDateString());
  a3_range.setValue(date.toLocaleTimeString());
}

// 2. Get the active cell data info
function main(workbook: ExcelScript.Workbook) {
  // 1. . get the active cell 
  let active_cell = workbook.getActiveCell();

  // 2. show the value on console
  console.log(active_cell.getValue());
  console.log(active_cell.getAddress());
}

// 3. Get selected range -- 巧用script录制功能，可以将自动生成的代码嵌入到自己的代码中。
function main(workbook: ExcelScript.Workbook) {
  let range = workbook.getSelectedRange();
  let rows = range.getRowCount();
  let cols = range.getColumnCount();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // 先获取选择的区域的数据，然后在到循环里面去操作数据，这样性能会提升。。
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}

// 4. Get All sheets name
function main(workbook: ExcelScript.Workbook) {
  let sheets = workbook.getWorksheets();
  let names = sheets.map((sheet) => sheet.getName());
  console.log(names);
  console.log(sheets.length);

  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;
    sheet.setTabColor(colorString);
  }
  
  /*
  // 1. delete the sheet
  for(let sheet of sheets){
	  sheet.delete();
  }
  
  // 2. add new sheet
  let newSheet = workbook.addWorksheet("Index");
  newSheet.activate();
  */
}


// 5. Set Formular into the cells
function main(workbook: ExcelScript.Workbook) {
  let test_worksheet = workbook.getWorksheet("test");
  let a1 = test_worksheet.getRange("A1");
  a1.setValue(2);
  
  let a2 = test_worksheet.getRange("A2");
  a2.setFormula("=(2*A1)");
  console.log(a2.getValue());
}

// 6. getPivotTables
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getBodyAndTotalRange();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console. 解析一下。。
  // 
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}

// 7. 