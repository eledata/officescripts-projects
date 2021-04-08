function main(workbook: ExcelScript.Workbook, keys: string): string {
  // 巧用代码录像功能，直接将需要动作代码封装到自己内部的函数里面。
  let table = workbook.getWorksheet('Transactions').getTables()[0];
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  let overallMatch = true;
  table.getColumnByName('Date').getRangeBetweenHeaderAndTotal().setNumberFormatLocal("yyyy-mm-dd;@");
  table.getColumnByName('Capacity').getRangeBetweenHeaderAndTotal().getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  
  let rows = range.getValues();
  let keysObject = JSON.parse(keys) as EventData[];

  for (let i = 0; i < rows.length; i++){
    let row = rows[i];
    let [event, date, location, capacity] = row;
    let match = false;

    // 校验功能
    for (let keyObject of keysObject){
      if (keyObject.event === event) {
        match = true;

        if (keyObject.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        if (keyObject.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        if (keyObject.capacity !== capacity) {
          overallMatch = false;
          range.getCell(i, 3).getFormat()
          .getFill()
          .setColor("FFFF00");
        }   
        break;             
      }

      // 模拟Vlookup功能
    for (let keyObject of keysObject){
      if (keyObject.event === event) {
          range.getCell(i, 3).setValue('testr');
        }   
        break;             
      }
    }

    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");      
    }

  }

  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  }
  console.log("Returning: " + returnString);
  return returnString;
}
  
// EventData 接口
interface EventData {
  event: string
  date: number
  location: string
  capacity: number
}