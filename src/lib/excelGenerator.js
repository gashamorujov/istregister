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

/* ── Colors from 28.07.2026.xlsx ── */
const NO_FILL = { type: 'pattern', pattern: 'none' };
const WHITE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
const LIGHT = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F8F9' } };
const THIN = 'thin';
const BLACK = { argb: 'FF000000' };

function bAll() {
  return { left:{style:THIN,color:BLACK}, right:{style:THIN,color:BLACK}, top:{style:THIN,color:BLACK}, bottom:{style:THIN,color:BLACK} };
}
function bNoLeft() {
  return { left:{style:undefined}, right:{style:THIN,color:BLACK}, top:{style:THIN,color:BLACK}, bottom:{style:THIN,color:BLACK} };
}
function bNoTop() {
  return { left:{style:THIN,color:BLACK}, right:{style:THIN,color:BLACK}, top:{style:undefined}, bottom:{style:THIN,color:BLACK} };
}
function bNoLeftNoTop() {
  return { left:{style:undefined}, right:{style:THIN,color:BLACK}, top:{style:undefined}, bottom:{style:THIN,color:BLACK} };
}

function applyColWidths(sheet) {
  sheet.getColumn(1).width = 14;
  sheet.getColumn(2).width = 211.86;
  sheet.getColumn(3).width = 54.43;
  sheet.getColumn(4).width = 99;
  sheet.getColumn(5).width = 53;
  sheet.getColumn(6).width = 102.14;
  sheet.getColumn(7).width = 8.71;
}

/* ── Main generator ── */
export async function generateTrainingPlan(filteredRecords, entries, templateBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  const sheet = workbook.getWorksheet('Plan');

  // Apply exact column widths from template
  applyColWidths(sheet);

  // Find last data row and clear from row 3 onward
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
    hdr.height = 33;

    const bH = hdr.getCell(2);
    bH.value = 'Kursun adı';
    bH.font = { name: 'Arial', size: 25, bold: true };
    bH.alignment = { horizontal: 'center', vertical: 'bottom' };
    bH.fill = WHITE;
    bH.border = bAll();

    const cH = hdr.getCell(3);
    cH.value = 'Tədrisin ümumi saatı';
    cH.font = { name: 'Arial', size: 25, bold: true };
    cH.alignment = { horizontal: 'center', vertical: 'bottom' };
    cH.fill = WHITE;
    cH.border = bAll();

    const dH = hdr.getCell(4);
    dH.value = 'Qrup nömrəsi';
    dH.font = { name: 'Arial', size: 25, bold: true };
    dH.alignment = { horizontal: 'center', vertical: 'bottom' };
    dH.fill = LIGHT;
    dH.border = bAll();

    const eH = hdr.getCell(5);
    eH.value = 'Başlama və bitmə tarixi';
    eH.font = { name: 'Arial', size: 25, bold: true };
    eH.alignment = { horizontal: 'center', vertical: 'bottom' };
    eH.fill = LIGHT;
    eH.border = bAll();

    const fH = hdr.getCell(6);
    fH.value = 'Kursu tədris edən müəllimlərin adı və soyadı';
    fH.font = { name: 'Arial', size: 25, bold: true };
    fH.alignment = { horizontal: 'center', vertical: 'bottom' };
    fH.fill = LIGHT;
    fH.border = bAll();

    currentRow++;

    /* ── Values row (course details) ── */
    const vRow = sheet.getRow(currentRow);
    vRow.height = 33;

    // B: course name — no fill (matches template pattern:"none")
    const bV = vRow.getCell(2);
    bV.value = courseName;
    bV.font = { name: 'Arial', size: 25, bold: true };
    bV.alignment = { horizontal: 'center', vertical: 'bottom', wrapText: true };
    bV.fill = NO_FILL;
    bV.border = bAll();

    // C: hours — white
    const cV = vRow.getCell(3);
    cV.value = courseHours;
    cV.font = { name: 'Arial', size: 25, bold: true };
    cV.alignment = { horizontal: 'center', vertical: 'bottom' };
    cV.fill = WHITE;
    cV.border = bAll();

    // D: group — light
    const dV = vRow.getCell(4);
    dV.value = groupNum;
    dV.font = { name: 'Arial', size: 25, bold: true };
    dV.alignment = { horizontal: 'center', vertical: 'bottom' };
    dV.fill = LIGHT;
    dV.border = bAll();

    // E: dates — white
    const eV = vRow.getCell(5);
    eV.value = dateRange;
    eV.font = { name: 'Arial', size: 25, bold: true };
    eV.alignment = { horizontal: 'center', vertical: 'bottom' };
    eV.fill = WHITE;
    eV.border = bAll();

    // F: teacher — light
    const fV = vRow.getCell(6);
    fV.value = teacher;
    fV.font = { name: 'Arial', size: 25, bold: true };
    fV.alignment = { horizontal: 'center', vertical: 'bottom' };
    fV.fill = LIGHT;
    fV.border = bAll();

    currentRow++;

    /* ── Student rows ── */
    students.forEach((student, idx) => {
      const sRow = sheet.getRow(currentRow);
      sRow.height = 33;

      // A: seq number — Arial 27pt bold, no fill, no top border
      const a = sRow.getCell(1);
      a.value = idx + 1;
      a.font = { name: 'Arial', size: 27, bold: true };
      a.alignment = { horizontal: 'center', vertical: 'bottom' };
      a.fill = NO_FILL;
      a.border = bNoTop();

      // B: name — Arial 28pt NOT bold, no fill
      const b = sRow.getCell(2);
      b.value = student.fullName || '';
      b.font = { name: 'Arial', size: 28 };
      b.alignment = { horizontal: 'center', vertical: 'bottom' };
      b.fill = NO_FILL;
      b.border = bNoLeftNoTop();

      // C: empty — LIGHT fill always
      const c = sRow.getCell(3);
      c.value = undefined;
      c.font = { name: 'Arial', size: 28 };
      c.alignment = { horizontal: 'center', vertical: 'bottom' };
      c.fill = LIGHT;
      c.border = bAll();

      // D: Vəzifə (Rank) — Arial 30pt NOT bold, LIGHT fill always
      const d = sRow.getCell(4);
      d.value = student.rank || '';
      d.font = { name: 'Arial', size: 30 };
      d.alignment = { horizontal: 'center', vertical: 'bottom' };
      d.fill = LIGHT;
      d.border = bNoLeftNoTop();

      // E: İlkin — Arial 30pt NOT bold, LIGHT fill always
      const e = sRow.getCell(5);
      e.value = 'İlkin';
      e.font = { name: 'Arial', size: 30 };
      e.alignment = { horizontal: 'center', vertical: 'bottom' };
      e.fill = LIGHT;
      e.border = bNoLeftNoTop();

      // F: empty — no fill
      const f = sRow.getCell(6);
      f.value = undefined;
      f.font = { name: 'Arial', size: 28 };
      f.alignment = { horizontal: 'center', vertical: 'bottom' };
      f.fill = NO_FILL;
      f.border = bAll();

      currentRow++;
    });

    // Empty separator row
    const sep = sheet.getRow(currentRow);
    sep.height = 33;
    for (let c = 1; c <= 6; c++) sep.getCell(c).value = undefined;
    currentRow++;
  });

  // Footer
  const foot = sheet.getRow(currentRow);
  foot.height = 33;
  const dF = foot.getCell(4);
  dF.value = 'Hörmətlə ';
  dF.font = { name: 'Arial', size: 25, bold: true };
  dF.alignment = { horizontal: 'center', vertical: 'bottom' };
  dF.fill = WHITE;
  dF.border = bAll();

  const fF = foot.getCell(6);
  fF.value = 'Dənizçilərin xüsusi hazırlıq üzrə təlim şöbəsi';
  fF.font = { name: 'Arial', size: 25, bold: true };
  fF.alignment = { horizontal: 'center', vertical: 'bottom' };
  fF.fill = WHITE;
  fF.border = bAll();

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
