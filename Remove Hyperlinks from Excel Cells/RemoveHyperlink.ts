// 演示：如何去除表格中的超链接
function main(workbook: ExcelScript.Workbook) {
  // 遍历文件中所有的sheet
  let sheets = workbook.getWorksheets();

  // 遍历所有sheet，去除hyper link
  for(let sheet of sheets){
    console.log("现在去除" + sheet.getName());
    // 获取需要处理的区域
    const target_range = sheet.getUsedRange(true);
    if (!target_range) {
      console.log(`There is no data in the worksheet. `)
      continue;
    }

    // 调取处理出去hyperlink函数
    removeHyperLink(target_range);

    // 将完成的sheet 标记上颜色
    let colorString = `#${Math.random().toString(16).substr(-6)}`;
    sheet.setTabColor(colorString);
  }
  return;
}

// 去除hyper link的函数
function removeHyperLink(target_range: ExcelScript.Range): void {
  const rowCount = target_range.getRowCount();
  const colCount = target_range.getColumnCount();
  let clearedCount = 0;

  for (let i = 0; i < rowCount; i++) {
    for (let j = 0; j < colCount; j++) {
      const cell = target_range.getCell(i, j);
      const hyperlink = cell.getHyperlink();
      if (hyperlink) {
        // 直接从录制的代码里面拷贝过来
        cell.clear(ExcelScript.ClearApplyTo.removeHyperlinks);
        cell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.none);
        cell.getFormat().getFont().setColor('Red');
        clearedCount++;
      }
    }
  }
  console.log(`Done. Clearned hyperlinks in: ${clearedCount} cells`);
  return;
}
