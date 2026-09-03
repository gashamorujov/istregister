import ExcelJS from 'exceljs';
import { getCourseName, getCourseHours } from '../data/courses';

const courseAbbreviations = {
  SK: 'GST', SH: 'GTY', SI: 'SOXQ', SO: 'GST', SL: 'TZYD',
  SA: 'GST', SP: 'MŞRİO', RS: 'RS', SG: 'G.MX', SW: 'ISPS',
  SV: 'IMO', SQ: 'KMDTYGP', SR: 'G.ELK', SZ: 'SXQ', SF: 'QYM',
  SD: 'ROV', SC: 'SA', SE: 'STCW', ST: 'YS', SX: 'G.OP',
  SN: 'N.MX', SM: 'MAŞ', DQ: 'D.QAZ', SB: 'B.OPER', AS: 'AŞ',
  ER: 'EL.MX', DL: 'DLG', SJ: 'KRAN', SU: 'SXQ',
};

function groupRecordsByCourseAndDate(records) {
  const groups = {};
  records.forEach(r => {
    if (!r.courseCode && !r.startDate && !r.fullName) return;
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

export async function generateTrainingPlan(filteredRecords, teacherSelections, templateBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  const sheet = workbook.getWorksheet('Plan');

  // Find the last row with data (the footer "Hörmətlə" row)
  let footerRow = sheet.rowCount;
  while (footerRow > 2 && !sheet.getRow(footerRow).getCell(6).value) {
    footerRow--;
  }
  // footerRow is the "Hörmətlə" row - we keep it

  // Delete all rows between header (row 2) and footer
  // We keep: row 1 (logo), row 2 (title)
  // We delete: rows 3 to footerRow-1
  // Then we rebuild from row 3
  // Actually, simpler: just clear rows 3..footerRow-1 and write fresh data

  // First, find the footer row value to save
  const footerRowData = [];
  const footerRowRef = sheet.getRow(footerRow);
  for (let c = 1; c <= 6; c++) {
    footerRowData.push(footerRowRef.getCell(c).value);
  }

  // Delete rows 3 through footerRow
  sheet.spliceRows(3, footerRow - 2);

  let currentRow = 3;
  let groupCounter = 1;
  const yearSuffix = String(new Date().getFullYear()).slice(-2);
  const groups = groupRecordsByCourseAndDate(filteredRecords);

  // Helper to style header cells
  function styleHeaderRow(row) {
    ['B', 'C', 'D', 'E', 'F'].forEach(col => {
      const cell = row.getCell(col);
      cell.font = { name: 'Arial', size: 25, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'center' };
      cell.border = {
        left: { style: 'thin' }, right: { style: 'thin' },
        top: { style: 'thin' }, bottom: { style: 'thin' }
      };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    });
  }

  // Helper to style value row (course details)
  function styleValueRow(row) {
    ['B', 'C', 'D', 'E', 'F'].forEach(col => {
      const cell = row.getCell(col);
      cell.font = { name: 'Arial', size: 25, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'center' };
      cell.border = {
        left: { style: 'thin' }, right: { style: 'thin' },
        top: { style: 'thin' }, bottom: { style: 'thin' }
      };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    });
  }

  // Helper to style student row
  function styleStudentRow(row) {
    // A: sequence number
    const a = row.getCell(1);
    a.font = { name: 'Arial', size: 27, bold: true };
    a.alignment = { horizontal: 'center' };
    a.border = { left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    // B: full name
    const b = row.getCell(2);
    b.font = { name: 'Arial', size: 30 };
    b.alignment = { horizontal: 'center' };
    b.border = { right: { style: 'thin' }, bottom: { style: 'thin' } };

    // C: empty header col
    const c = row.getCell(3);
    c.font = { name: 'Arial', size: 25, bold: true };
    c.alignment = { horizontal: 'center' };
    c.border = { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    // D: rank
    const d = row.getCell(4);
    d.font = { name: 'Arial', size: 30 };
    d.alignment = { horizontal: 'center' };
    d.border = { right: { style: 'thin' }, bottom: { style: 'thin' } };

    // E: empty header col
    const e = row.getCell(5);
    e.font = { name: 'Arial', size: 25, bold: true };
    e.alignment = { horizontal: 'center' };
    e.border = { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
    e.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    // F: status (İlkin/Təkrar)
    const f = row.getCell(6);
    f.font = { name: 'Arial', size: 30 };
    f.alignment = { horizontal: 'center' };
    f.border = { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
    f.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F8F9' } };
  }

  groups.forEach(group => {
    const { courseCode, startDate, finishDate, students } = group;
    if (students.length === 0) return;

    const courseName = getCourseName(courseCode);
    const courseHours = getCourseHours(courseCode);
    const teacherName = teacherSelections[courseCode] || '';
    const abbr = courseAbbreviations[courseCode] || courseCode;
    const groupNum = `${String(groupCounter).padStart(3, '0')}/${yearSuffix} ${abbr}`;

    // Header row (labels)
    const hdr = sheet.getRow(currentRow);
    hdr.getCell(2).value = 'Kursun adı';
    hdr.getCell(3).value = 'Tədrisin ümumi saatı';
    hdr.getCell(4).value = 'Qrup nömrəsi';
    hdr.getCell(5).value = 'Başlama və bitmə tarixi';
    hdr.getCell(6).value = 'Kursu tədris edən müəllimlərin adı və soyadı';
    styleHeaderRow(hdr);
    currentRow++;

    // Values row
    const val = sheet.getRow(currentRow);
    val.getCell(2).value = courseName;
    val.getCell(3).value = courseHours;
    val.getCell(4).value = groupNum;
    val.getCell(5).value = `${startDate || ''}-${finishDate || ''}`;
    val.getCell(6).value = teacherName;
    styleValueRow(val);
    currentRow++;

    // Student rows
    students.forEach((student, idx) => {
      const sRow = sheet.getRow(currentRow);
      sRow.getCell(1).value = idx + 1;
      sRow.getCell(2).value = student.fullName;
      sRow.getCell(4).value = student.rank;
      sRow.getCell(6).value = 'İlkin';
      styleStudentRow(sRow);
      currentRow++;
    });

    groupCounter++;

    // Empty separator row
    currentRow++;
  });

  // Add footer row
  const foot = sheet.getRow(currentRow);
  foot.getCell(4).value = footerRowData[3] || 'Hörmətlə';
  foot.getCell(6).value = footerRowData[5] || 'Dənizçilərin xüsusi hazırlıq üzrə təlim şöbəsi';
  ['A', 'C', 'F'].forEach(col => {
    const cell = foot.getCell(col);
    cell.font = { name: 'Arial', size: 25, bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = { left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
  });
  foot.getCell('C').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
  foot.getCell('F').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F8F9' } };

  // Download
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
