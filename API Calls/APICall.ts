async function main(workbook: ExcelScript.Workbook) {
  // 清除sheet内容
  const sheet = workbook.getActiveWorksheet();
  sheet.getRanges().clear(ExcelScript.ClearApplyTo.contents);

  // 增加列名
  let col_name_lst = ['id', 'node_id', 'full_name', 'html_url', 'description', 'license']
  let startCell = sheet.getRange('A1');
  const targetRange = startCell.getResizedRange(0, col_name_lst.length - 1);      
  targetRange.setValues([col_name_lst]);
  targetRange.getFormat().getFill().setColor("BFBFBF");     

  // 获取api数据
  const response = await fetch('https://api.github.com/users/eledata/repos');
  const repos: Repository[] = await response.json();
  
  const rows: (string | boolean | number)[][] = [];
  for (let repo of repos){ 
    rows.push([repo.id, repo.node_id, repo.full_name, repo.html_url, repo.description, repo.license?.name, repo.license?.url])
  }

  // 将整理好的数据放入excel中
  const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
  range.setValues(rows);
  return;
}

interface Repository {
  id: string,
  node_id: string,
  full_name: string,
  html_url: string,
  description: string,
  license?: License 
}

interface License {
  name: string,
  url: string
}
