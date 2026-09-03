import ExcelJS from 'exceljs';
import { getCourseName, getCourseHours } from '../data/courses';

function groupRecordsByCourseAndDate(records) {
  const groups = {};
  records.forEach(r => {
    if (!r.fullName && !r.courseCode) return;
    const key = `${r.courseCode}__${r.startDate}__${r.finishDate}`;
    if (!groups[key]) {
      groups[key] = { courseCode: r.courseCode, startDate: r.startDate, finishDate: r.finishDate, students: [] };
    }
    groups[key].students.push(r);
  });
  return Object.values(groups);
}

export function getUniqueCourseGroups(records) {
  return groupRecordsByCourseAndDate(records).map(g => ({
    courseCode: g.courseCode,
    startDate: g.startDate,
    finishDate: g.finishDate,
    studentCount: g.students.length,
    courseName: getCourseName(g.courseCode),
  }));
}

function applyBorder(cell, style) {
  cell.border = {
    left: { style }, right: { style },
    top: { style }, bottom: { style }
  };
}

export async function generateTrainingPlan(filteredRecords, groupNumbers, templateBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  const sheet = workbook.getWorksheet('Plan');

  // Find footer row
  let footerRow = sheet.rowCount;
  while (footerRow > 2 && !sheet.getRow(footerRow).getCell(6).value) footerRow--;
  const footerData = [];
  for (let c = 1; c <= 6; c++) footerData.push(sheet.getRow(footerRow).getCell(c).value);

  // Delete rows 3 to footerRow
  sheet.spliceRows(3, footerRow - 2);

  let currentRow = 3;
  const groups = groupRecordsByCourseAndDate(filteredRecords);

  groups.forEach(group => {
    const { courseCode, startDate, finishDate, students } = group;
    if (students.length === 0) return;

    const courseName = getCourseName(courseCode);
    const courseHours = getCourseHours(courseCode);
    const groupNum = groupNumbers[courseCode] || '';
    const dateRange = `${startDate || ''}-${finishDate || ''}`;

    // --- Header row (labels) ---
    const hdr = sheet.getRow(currentRow);
    hdr.height = 22;
    const hdrLabels = ['Kursun adı', 'Tədrisin ümumi saatı', 'Qrup nömrəsi', 'Başlama və bitmə tarixi', 'Kursu tədris edən müəllimlərin adı və soyadı'];
    [2, 3, 4, 5, 6].forEach((col, i) => {
      const cell = hdr.getCell(col);
      cell.value = hdrLabels[i];
      cell.font = { name: 'Arial', size: 11, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      applyBorder(cell, 'thin');
    });
    currentRow++;

    // --- Values row ---
    const val = sheet.getRow(currentRow);
    val.height = 22;
    const vals = [courseName, courseHours, groupNum, dateRange, ''];
    [2, 3, 4, 5, 6].forEach((col, i) => {
      const cell = val.getCell(col);
      cell.value = vals[i];
      cell.font = { name: 'Arial', size: 11 };
      cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      applyBorder(cell, 'thin');
    });
    currentRow++;

    // --- Student rows ---
    students.forEach((student, idx) => {
      const sRow = sheet.getRow(currentRow);
      sRow.height = 20;

      // A: sequence
      const a = sRow.getCell(1);
      a.value = idx + 1;
      a.font = { name: 'Arial', size: 11, bold: true };
      a.alignment = { horizontal: 'center' };
      applyBorder(a, 'thin');

      // B: full name
      const b = sRow.getCell(2);
      b.value = student.fullName || '';
      b.font = { name: 'Arial', size: 11 };
      b.alignment = { horizontal: 'center' };
      applyBorder(b, 'thin');

      // C: empty
      const c = sRow.getCell(3);
      c.font = { name: 'Arial', size: 11 };
      c.alignment = { horizontal: 'center' };
      applyBorder(c, 'thin');

      // D: rank
      const d = sRow.getCell(4);
      d.value = student.rank || '';
      d.font = { name: 'Arial', size: 11 };
      d.alignment = { horizontal: 'center' };
      applyBorder(d, 'thin');

      // E: empty
      const e = sRow.getCell(5);
      e.font = { name: 'Arial', size: 11 };
      e.alignment = { horizontal: 'center' };
      applyBorder(e, 'thin');

      // F: status
      const f = sRow.getCell(6);
      f.value = 'İlkin';
      f.font = { name: 'Arial', size: 11 };
      f.alignment = { horizontal: 'center' };
      f.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F8F9' } };
      applyBorder(f, 'thin');

      currentRow++;
    });

    // Empty separator
    currentRow++;
  });

  // Footer
  const foot = sheet.getRow(currentRow);
  foot.getCell(4).value = footerData[3] || 'Hörmətlə';
  foot.getCell(6).value = footerData[5] || 'Dənizçilərin xüsusi hazırlıq üzrə təlim şöbəsi';
  [1, 3, 4, 5, 6].forEach(col => {
    const cell = foot.getCell(col);
    cell.font = { name: 'Arial', size: 11, bold: true };
    cell.alignment = { horizontal: 'center' };
    applyBorder(cell, 'thin');
  });
  foot.getCell(4).alignment = { horizontal: 'right' };

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Training_Plan.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
