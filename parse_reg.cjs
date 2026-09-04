const ExcelJS = require('exceljs');
const fs = require('fs');
async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('/tmp/codex-web-uploads/f-qpgEc7/REGISTR-2026.xlsx');
  const sheet = wb.getWorksheet(1);
  const rows = [];
  const headerRow = sheet.getRow(1);
  const headers = [];
  for (let c = 1; c <= 29; c++) {
    headers.push(String(headerRow.getCell(c).value || ''));
  }
  fs.writeFileSync('/tmp/reg_headers.json', JSON.stringify(headers, null, 2));
  
  let count = 0;
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const obj = {};
    let hasAny = false;
    for (let c = 1; c <= 29; c++) {
      const v = row.getCell(c).value;
      if (v !== null && v !== undefined && String(v).trim()) hasAny = true;
      obj['c' + c] = v;
    }
    if (hasAny) {
      rows.push(obj);
      count++;
    }
  }
  fs.writeFileSync('/tmp/reg_data.json', JSON.stringify(rows));
  console.log('Total data rows:', count);
  console.log('Headers:', JSON.stringify(headers));
}
main().catch(e => console.error(e));
