import ExcelJS from 'exceljs';
import { getCourseName, getCourseHours } from '../data/courses';

/* ============================================================================
 * Training Plan export
 * ----------------------------------------------------------------------------
 * This used to load a bundled Training_plan_template.xlsx and mutate it in
 * place (clear rows, re-fill them). That template had accumulated 800+ rows
 * of leftover cell formatting from earlier exports, and several of its styles
 * (a fill color here, a border there) had drifted from what the reference
 * "Training_Plan.xlsx" actually looks like — which is why past exports didn't
 * quite match it.
 *
 * This version builds the workbook from scratch every time, using values
 * measured directly from the reference file (fonts, sizes, bold, fill colors,
 * border sides, column widths, row heights, logo position/size). There is no
 * template file to keep in sync anymore — every visual rule lives in this
 * file, next to the code that uses it. The only bundled asset is the IST logo
 * image itself (src/assets/ist-logo.png), which gets embedded directly.
 * ========================================================================== */

/* ── Date helpers ── */
function addDays(dateStr, days) {
  if (!dateStr) return '';
  const parts = dateStr.split('.');
  if (parts.length !== 3) return '';
  const d = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  d.setDate(d.getDate() + days);
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}.${mm}.${d.getFullYear()}`;
}

/* Group by courseCode + startDate only (ignore finishDate) */
function groupRecords(records) {
  const groups = {};
  records.forEach(r => {
    if (!r.fullName && !r.courseCode) return;
    const key = `${r.courseCode}__${r.startDate}`;
    if (!groups[key]) {
      groups[key] = { courseCode: r.courseCode, startDate: r.startDate, students: [] };
    }
    groups[key].students.push(r);
  });
  return Object.values(groups);
}

export function getUniqueCourseGroups(records) {
  return groupRecords(records).map(g => {
    const hours = getCourseHours(g.courseCode);
    const days = Math.ceil(hours / 8);
    const finishDate = addDays(g.startDate, days - 1);
    return {
      courseCode: g.courseCode,
      startDate: g.startDate,
      finishDate,
      studentCount: g.students.length,
      courseName: getCourseName(g.courseCode),
    };
  });
}

/* ============================================================================
 * Exact visual spec — every number below was measured from the reference
 * Training_Plan.xlsx (openpyxl cell-by-cell inspection + a rendered preview),
 * not eyeballed. Column letters in comments refer to that file's layout.
 * ========================================================================== */

const FONT = 'Arial';
const BLACK = { argb: 'FF000000' };
const THIN = 'thin';

const WHITE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
const LIGHT_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F8F9' } };

// Column A (seq. number): boxed on left/right/bottom. No top side anywhere in the sheet —
// every row's top edge is supplied by the row above's bottom border instead, so two
// adjacent thin lines never get drawn on top of each other.
function borderSeq() {
  return {
    left: { style: THIN, color: BLACK },
    right: { style: THIN, color: BLACK },
    bottom: { style: THIN, color: BLACK },
  };
}
// Fully boxed on all 4 sides. Used by the label/value rows (columns B–F) and, in every row
// of the sheet, by column C — which is why column C reads as a continuous boxed stripe.
function borderBoxed() {
  return {
    left: { style: THIN, color: BLACK },
    right: { style: THIN, color: BLACK },
    top: { style: THIN, color: BLACK },
    bottom: { style: THIN, color: BLACK },
  };
}
// Right + bottom only. Used by columns B, D, E, F on student/footer rows — the left edge
// comes from the neighbouring cell's right border, the top edge from the row above.
function borderRightBottom() {
  return {
    right: { style: THIN, color: BLACK },
    bottom: { style: THIN, color: BLACK },
  };
}

const COLS = { SEQ: 1, NAME: 2, HOURS: 3, RANK: 4, DATE: 5, TEACHER: 6 };

function applyColWidths(sheet) {
  sheet.getColumn(COLS.SEQ).width = 13.98828125;
  sheet.getColumn(COLS.NAME).width = 211.875;
  sheet.getColumn(COLS.HOURS).width = 54.48046875;
  sheet.getColumn(COLS.RANK).width = 99.0078125;
  sheet.getColumn(COLS.DATE).width = 53;
  sheet.getColumn(COLS.TEACHER).width = 102.1015625;
  sheet.getColumn(7).width = 8.7421875;
}

/* ── Row builders, one per visual role ── */

// Column-label row ("Kursun adı", "Tədrisin ümumi saatı", ...): Arial 25 bold, white fill, boxed.
function writeLabelRow(sheet, rowNum) {
  const row = sheet.getRow(rowNum);
  row.height = 33;
  const labels = {
    [COLS.NAME]: 'Kursun adı',
    [COLS.HOURS]: 'Tədrisin ümumi saatı',
    [COLS.RANK]: 'Qrup nömrəsi',
    [COLS.DATE]: 'Başlama və bitmə tarixi',
    [COLS.TEACHER]: 'Kursu tədris edən müəllimlərin adı və soyadı',
  };

  const seq = row.getCell(COLS.SEQ);
  seq.font = { name: FONT, size: 27, bold: true };
  seq.alignment = { horizontal: 'center' };
  seq.border = borderSeq();

  [COLS.NAME, COLS.HOURS, COLS.RANK, COLS.DATE, COLS.TEACHER].forEach(c => {
    const cell = row.getCell(c);
    cell.value = labels[c];
    cell.font = { name: FONT, size: 25, bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.fill = WHITE_FILL;
    cell.border = borderBoxed();
  });
}

// Values row (course name / hours / group / dates / teacher) — same box + font as the label
// row above it, carrying the actual data for one course group.
function writeValuesRow(sheet, rowNum, { courseName, courseHours, groupNum, dateRange, teacher }) {
  const row = sheet.getRow(rowNum);
  row.height = 33;
  const values = {
    [COLS.NAME]: courseName,
    [COLS.HOURS]: courseHours,
    [COLS.RANK]: groupNum,
    [COLS.DATE]: dateRange,
    [COLS.TEACHER]: teacher,
  };

  const seq = row.getCell(COLS.SEQ);
  seq.font = { name: FONT, size: 27, bold: true };
  seq.alignment = { horizontal: 'center' };
  seq.border = borderSeq();

  [COLS.NAME, COLS.HOURS, COLS.RANK, COLS.DATE, COLS.TEACHER].forEach(c => {
    const cell = row.getCell(c);
    cell.value = values[c];
    cell.font = { name: FONT, size: 25, bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.fill = WHITE_FILL;
    cell.border = borderBoxed();
  });
}

// One student line: seq number (27 bold) | name (30, left-aligned) | spacer (30, light+boxed)
// | rank (30) | spacer (30) | status (30).
function writeStudentRow(sheet, rowNum, seqNum, student) {
  const row = sheet.getRow(rowNum);
  row.height = 33;

  const seq = row.getCell(COLS.SEQ);
  seq.value = seqNum;
  seq.font = { name: FONT, size: 27, bold: true };
  seq.alignment = { horizontal: 'center' };
  seq.border = borderSeq();

  const name = row.getCell(COLS.NAME);
  name.value = student.fullName || '';
  name.font = { name: FONT, size: 30 };
  name.alignment = { horizontal: 'left' };
  name.border = borderRightBottom();

  const hours = row.getCell(COLS.HOURS); // decorative spacer column in student rows
  hours.font = { name: FONT, size: 30 };
  hours.alignment = { horizontal: 'center' };
  hours.fill = LIGHT_FILL;
  hours.border = borderBoxed();

  const rank = row.getCell(COLS.RANK);
  rank.value = student.rank || '';
  rank.font = { name: FONT, size: 30, bold: false };
  rank.alignment = { horizontal: 'center' };
  rank.border = borderRightBottom();

  const date = row.getCell(COLS.DATE); // decorative spacer column in student rows
  date.font = { name: FONT, size: 30 };
  date.alignment = { horizontal: 'center' };
  date.border = borderRightBottom();

  const teacher = row.getCell(COLS.TEACHER);
  teacher.value = 'İlkin';
  teacher.font = { name: FONT, size: 30, bold: false };
  teacher.alignment = { horizontal: 'center' };
  teacher.border = borderRightBottom();
}

// Blank line between groups. Left fully empty except column C, which keeps the light
// "spacer" stripe running continuously down the sheet, matching the reference file.
function writeSeparatorRow(sheet, rowNum) {
  const row = sheet.getRow(rowNum);
  row.height = 33;
  row.getCell(COLS.SEQ).border = borderSeq();
  const spacer = row.getCell(COLS.HOURS);
  spacer.fill = LIGHT_FILL;
  spacer.border = borderBoxed();
}

// Closing line: "Hörmətlə" under Rank, department name under Teacher — styled like the
// label/values rows since both are real (bold, boxed, white) content, not blank fields.
function writeFooterRow(sheet, rowNum) {
  const row = sheet.getRow(rowNum);
  row.height = 33;

  const seq = row.getCell(COLS.SEQ);
  seq.font = { name: FONT, size: 27, bold: true };
  seq.alignment = { horizontal: 'center' };
  seq.border = borderSeq();

  const name = row.getCell(COLS.NAME);
  name.font = { name: FONT, size: 30 };
  name.alignment = { horizontal: 'left' };
  name.border = borderRightBottom();

  const hours = row.getCell(COLS.HOURS);
  hours.font = { name: FONT, size: 30 };
  hours.alignment = { horizontal: 'center' };
  hours.fill = LIGHT_FILL;
  hours.border = borderBoxed();

  const rank = row.getCell(COLS.RANK);
  rank.value = 'Hörmətlə ';
  rank.font = { name: FONT, size: 25, bold: true };
  rank.alignment = { horizontal: 'center' };
  rank.fill = WHITE_FILL;
  rank.border = borderBoxed();

  const date = row.getCell(COLS.DATE);
  date.font = { name: FONT, size: 30 };
  date.alignment = { horizontal: 'center' };
  date.border = borderRightBottom();

  const teacher = row.getCell(COLS.TEACHER);
  teacher.value = 'Dənizçilərin xüsusi hazırlıq üzrə təlim şöbəsi';
  teacher.font = { name: FONT, size: 25, bold: true };
  teacher.alignment = { horizontal: 'center' };
  teacher.fill = WHITE_FILL;
  teacher.border = borderBoxed();
}

// Logo placement, measured from the reference file's drawing XML: anchored at column A,
// row 0, offset 704850 EMU (=74px) from the left, sized 5514975 x 2609850 EMU (=579x274px).
// Column A is 13.98828125 characters wide, i.e. 98px at this workbook's default font size,
// so the offset is expressed as a fraction of column A's width (74/98) for ExcelJS's anchor.
function addLogo(workbook, sheet, logoBuffer) {
  if (!logoBuffer) return;
  const imageId = workbook.addImage({ buffer: logoBuffer, extension: 'png' });
  sheet.addImage(imageId, {
    tl: { col: 74 / 98, row: 0 },
    ext: { width: 579, height: 274 },
  });
}

/* ── Main generator ── */
export async function generateTrainingPlan(filteredRecords, entries, logoBuffer) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Plan', {
    pageSetup: {
      orientation: 'landscape',
      paperSize: 9, // A4
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0, footer: 0 },
    },
  });

  applyColWidths(sheet);

  // Row 1: thin spacer the logo overlaps.
  sheet.getRow(1).height = 14.25;

  // Row 2: institution title, merged B2:F2, with the logo floating over its left edge.
  sheet.mergeCells('B2:F2');
  const titleRow = sheet.getRow(2);
  titleRow.height = 181.5;
  const title = sheet.getCell('B2');
  title.value = 'Industrial Support and Training MMC təlim-tədris müəssisəsində tədris edilən xüsusi hazırlıq kurslarına dair həftəlik dərs cədvəli  ';
  title.font = { name: FONT, size: 28, bold: true };
  title.alignment = { horizontal: 'center', vertical: 'middle' };

  addLogo(workbook, sheet, logoBuffer);

  const groups = groupRecords(filteredRecords);
  let currentRow = 3;

  groups.forEach(group => {
    const { courseCode, startDate, students } = group;
    if (students.length === 0) return;

    const entry = entries[courseCode] || {};
    const courseName = getCourseName(courseCode);
    const courseHours = getCourseHours(courseCode);
    const groupNum = entry.groupNum || '';
    const teacher = entry.teacher || '';

    // finishDate is derived, never stored: hours -> 8h/day -> calendar days from startDate.
    const days = Math.ceil(courseHours / 8);
    const finishDate = addDays(startDate, days - 1);
    const dateRange = (startDate && finishDate) ? `${startDate} - ${finishDate}` : (startDate || '');

    writeLabelRow(sheet, currentRow);
    currentRow++;

    writeValuesRow(sheet, currentRow, { courseName, courseHours, groupNum, dateRange, teacher });
    currentRow++;

    students.forEach((student, idx) => {
      writeStudentRow(sheet, currentRow, idx + 1, student);
      currentRow++;
    });

    writeSeparatorRow(sheet, currentRow);
    currentRow++;
  });

  writeFooterRow(sheet, currentRow);

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
