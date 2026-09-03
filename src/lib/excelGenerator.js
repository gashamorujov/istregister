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

/* ── Template-matching style constants ── */
const FILL_WHITE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
const FILL_LIGHT = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F8F9' } };
const THIN = 'thin';
const BLACK = { argb: 'FF000000' };

function thinBorder(all = true) {
  const s = all ? THIN : undefined;
  return { left: { style: s, color: BLACK }, right: { style: THIN, color: BLACK }, top: { style: s || THIN, color: BLACK }, bottom: { style: THIN, color: BLACK } };
}

function thinBorderNoTop() {
  return { left: { style: THIN, color: BLACK }, right: { style: THIN, color: BLACK }, top: { style: undefined }, bottom: { style: THIN, color: BLACK } };
}

function thinBorderNoLeftNoTop() {
  return { left: { style: undefined }, right: { style: THIN, color: BLACK }, top: { style: undefined }, bottom: { style: THIN, color: BLACK } };
}

function thinBorderNoLeft() {
  return { left: { style: undefined }, right: { style: THIN, color: BLACK }, top: { style: THIN, color: BLACK }, bottom: { style: THIN, color: BLACK } };
}

function styleHeaderLabel(cell, fill) {
  cell.font = { name: 'Arial', size: 25, bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'center' };
  cell.fill = fill;
  cell.border = fill === FILL_WHITE ? thinBorderNoLeft() : thinBorder();
}

function styleValueCell(cell, fill) {
  cell.font = { name: 'Arial', size: 25, bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
  cell.fill = fill;
  cell.border = fill === FILL_WHITE ? thinBorderNoLeft() : thinBorder();
}

function styleStudentNum(cell) {
  cell.font = { name: 'Arial', size: 27, bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'center' };
  cell.border = thinBorderNoTop();
}

function styleStudentText(cell) {
  cell.font = { name: 'Arial', size: 30 };
  cell.alignment = { horizontal: 'center', vertical: 'center' };
  cell.border = thinBorderNoLeftNoTop();
}

function styleStudentEmpty(cell) {
  cell.font = { name: 'Arial', size: 25, bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'center' };
  cell.fill = FILL_WHITE;
  cell.border = thinBorder();
}

function styleStudentEmptyAlt(cell) {
  cell.font = { name: 'Arial', size: 25, bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'center' };
  cell.fill = FILL_WHITE;
  cell.border = thinBorder();
}

function styleStudentStatus(cell) {
  cell.font = { name: 'Arial', size: 30 };
  cell.alignment = { horizontal: 'center', vertical: 'center' };
  cell.fill = FILL_LIGHT;
  cell.border = thinBorder();
}

function styleStudentFill(cell, fill) {
  cell.font = { name: 'Arial', size: 25, bold: true };
  cell.alignment = { horizontal: 'center', vertical: 'center' };
  cell.fill = fill;
  cell.border = thinBorder();
}

/* ── Main generator ── */
export async function generateTrainingPlan(filteredRecords, entries, templateBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  const sheet = workbook.getWorksheet('Plan');

  // Find existing content rows and delete from row 3 onward
  let lastRow = sheet.rowCount;
  while (lastRow > 2) {
    const row = sheet.getRow(lastRow);
    let hasData = false;
    for (let c = 1; c <= 6; c++) {
      if (row.getCell(c).value !== null && row.getCell(c).value !== undefined) { hasData = true; break; }
    }
    if (hasData) break;
    lastRow--;
  }
  if (lastRow > 2) sheet.spliceRows(3, lastRow - 2);

  // Write course blocks
  const groups = groupRecordsByCourseAndDate(filteredRecords);
  let currentRow = 3;

  groups.forEach(group => {
    const { courseCode, startDate, finishDate, students } = group;
    if (students.length === 0) return;

    const entry = entries[courseCode] || {};
    const courseName = getCourseName(courseCode);
    const courseHours = getCourseHours(courseCode);
    const groupNum = entry.groupNum || '';
    const teacher = entry.teacher || '';
    const dateRange = (startDate && finishDate) ? `${startDate}-${finishDate}` : (startDate || finishDate || '');

    /* ── Header labels row ── */
    const hdr = sheet.getRow(currentRow);
    hdr.height = 42;

    // B: Kursun adı (white fill, no left border)
    const bH = hdr.getCell(2);
    bH.value = 'Kursun adı';
    styleHeaderLabel(bH, FILL_WHITE);

    // C: Tədrisin ümumi saatı (white fill, no left border)
    const cH = hdr.getCell(3);
    cH.value = 'Tədrisin ümumi saatı';
    styleHeaderLabel(cH, FILL_WHITE);

    // D: Qrup nömrəsi (light fill, all borders)
    const dH = hdr.getCell(4);
    dH.value = 'Qrup nömrəsi';
    styleHeaderLabel(dH, FILL_LIGHT);

    // E: Başlama və bitmə tarixi (light fill, all borders)
    const eH = hdr.getCell(5);
    eH.value = 'Başlama və bitmə tarixi';
    styleHeaderLabel(eH, FILL_LIGHT);

    // F: Müəllimlər (light fill, all borders)
    const fH = hdr.getCell(6);
    fH.value = 'Kursu tədris edən müəllimlərin adı və soyadı';
    styleHeaderLabel(fH, FILL_LIGHT);

    currentRow++;

    /* ── Values row (course details) ── */
    const vRow = sheet.getRow(currentRow);
    vRow.height = 94.5;

    const bV = vRow.getCell(2);
    bV.value = courseName;
    styleValueCell(bV, FILL_WHITE);

    const cV = vRow.getCell(3);
    cV.value = courseHours;
    styleValueCell(cV, FILL_WHITE);

    const dV = vRow.getCell(4);
    dV.value = groupNum;
    styleValueCell(dV, FILL_LIGHT);

    const eV = vRow.getCell(5);
    eV.value = dateRange;
    styleValueCell(eV, FILL_WHITE);

    const fV = vRow.getCell(6);
    fV.value = teacher;
    styleValueCell(fV, FILL_WHITE);

    currentRow++;

    /* ── Student rows ── */
    students.forEach((student, idx) => {
      const sRow = sheet.getRow(currentRow);
      sRow.height = 42;

      // A: sequence number
      const aCell = sRow.getCell(1);
      aCell.value = idx + 1;
      styleStudentNum(aCell);

      // B: full name
      const bCell = sRow.getCell(2);
      bCell.value = student.fullName || '';
      styleStudentText(bCell);

      // C: empty (white fill, all borders)
      const cCell = sRow.getCell(3);
      cCell.value = undefined;
      styleStudentEmpty(cCell);

      // D: rank
      const dCell = sRow.getCell(4);
      dCell.value = student.rank || '';
      styleStudentText(dCell);

      // E: empty (white fill, all borders)
      const eCell = sRow.getCell(5);
      eCell.value = undefined;
      styleStudentEmpty(eCell);

      // F: status "İlkin" (light fill)
      const fCell = sRow.getCell(6);
      fCell.value = 'İlkin';
      styleStudentStatus(fCell);

      currentRow++;
    });

    // Empty separator row
    const sepRow = sheet.getRow(currentRow);
    sepRow.height = 42;
    for (let c = 1; c <= 6; c++) sepRow.getCell(c).value = undefined;
    currentRow++;
  });

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
