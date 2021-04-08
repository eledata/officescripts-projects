// 1. Generate the return data to Powerautomate.
// 2. Setting the interface structure for transfer data between Powerautomate and the mapping file..
function main(workbook: ExcelScript.Workbook): EventData[] {
    // 获取行信息
    let table = workbook.getWorksheet('Keys').getTables()[0];
    let range = table.getRangeBetweenHeaderAndTotal();
    let rows = range.getValues();

    let records: EventData[] = [];
    for (let row of rows) {
        let [event, date, location, capacity] = row;
        records.push({
            event: event as string,
            date: date as number, 
            location: location as string,
            capacity: capacity as number
        })
    }
    console.log(JSON.stringify(records))
    return records;
}

// data structure
interface EventData {
    event: string
    date: number
    location: string
    capacity: number
}