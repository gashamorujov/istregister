import ExcelJS from 'exceljs';
import { getCourseName, getCourseHours } from '../data/courses';

/**
 * Training Plan Excel Generator
 * Generates a Training Plan Excel file based on filtered registry data.
 * Follows the exact structure of the Training_plan.xlsx template.
 */

// Course code abbreviation mapping for group numbers
const courseAbbreviations = {
  SK: 'GST', SH: 'GTY', SI: 'SOXQ', SO: 'GST', SL: 'TZYD',
  SA: 'GST', SP: 'MŞRİO', RS: 'RS', SG: 'G.MX', SW: 'ISPS',
  SV: 'IMO', SQ: 'KMDTYGP', SR: 'G.ELK', SZ: 'SXQ', SF: 'QYM',
  SD: 'ROV', SC: 'SA', SE: 'STCW', ST: 'YS', SX: 'G.OP',
  SN: 'N.MX', SM: 'MAŞ', DQ: 'D.QAZ', SB: 'B.OPER', AS: 'AŞ',
  ER: 'EL.MX', DL: 'DLG', SJ: 'KRAN', SU: 'SXQ',
};

function parseDate(dateStr) {
  if (!dateStr) return null;
  const str = String(dateStr).trim();
  // DD.MM.YYYY
  const m1 = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m1) return new Date(parseInt(m1[3]), parseInt(m1[2]) - 1, parseInt(m1[1]));
  // YYYY-MM-DD
  const m2 = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m2) return new Date(parseInt(m2[1]), parseInt(m2[2]) - 1, parseInt(m2[3]));
  return null;
}

function formatDate(d) {
  if (!d) return '';
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}.${month}.${year}`;
}

function groupRecordsByCourseAndDate(records) {
  const groups = {};
  records.forEach(r => {
    const key = `${r.courseCode}__${r.startDate}__${r.finishDate}`;
    if (!groups[key]) {
      groups[key] = {
        courseCode: r.courseCode,
        startDate: r.startDate,
        finishDate: r.finishDate,
        students: []
      };
    }
    groups[key].students.push(r);
  });
  return Object.values(groups);
}

export async function generateTrainingPlan(filteredRecords, teacherSelections) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Plan');

  // Set column widths
  sheet.getColumn(1).width = 8;   // A - sıra nömrəsi
  sheet.getColumn(2).width = 45;  // B - tam ad
  sheet.getColumn(3).width = 5;   // C - boş
  sheet.getColumn(4).width = 35;  // D - vəzifə
  sheet.getColumn(5).width = 5;   // E - boş
  sheet.getColumn(6).width = 12;  // F - status

  // Row 1: Logo area
  sheet.getRow(1).height = 30;
  sheet.getCell('A1').value = '           ';

  // Row 2: Title
  sheet.mergeCells('B2:F2');
  sheet.getCell('B2').value = 'Industrial Support and Training MMC təlim-tədris müəssisəsində tədris edilən xüsusi hazırlıq kurslarına dair həftəlik dərs cədvəli';
  sheet.getCell('B2').font = { bold: true, size: 11, name: 'Arial' };
  sheet.getCell('B2').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  sheet.getRow(2).height = 40;

  let currentRow = 3;
  let groupCounter = 1;
  const yearSuffix = String(new Date().getFullYear()).slice(-2);

  // Group records by course code and date range
  const groups = groupRecordsByCourseAndDate(filteredRecords);

  groups.forEach(group => {
    const { courseCode, startDate, finishDate, students } = group;
    const courseName = getCourseName(courseCode);
    const courseHours = getCourseHours(courseCode);
    const teacherName = teacherSelections[courseCode] || '';
    const abbr = courseAbbreviations[courseCode] || courseCode;
    const groupNum = `${String(groupCounter).padStart(3, '0')}/${yearSuffix} ${abbr}`;

    // Header row: labels
    const headerRow = sheet.getRow(currentRow);
    headerRow.getCell(2).value = 'Kursun adı';
    headerRow.getCell(3).value = 'Tədrisin ümumi saatı';
    headerRow.getCell(4).value = 'Qrup nömrəsi';
    headerRow.getCell(5).value = 'Başlama və bitmə tarixi';
    headerRow.getCell(6).value = 'Kursu tədris edən müəllimlərin adı və soyadı';
    headerRow.eachCell(cell => {
      cell.font = { bold: true, size: 10, name: 'Arial' };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin' }, bottom: { style: 'thin' },
        left: { style: 'thin' }, right: { style: 'thin' }
      };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    });
    headerRow.height = 20;
    currentRow++;

    // Values row
    const dateRange = `${startDate}-${finishDate}`;
    const valuesRow = sheet.getRow(currentRow);
    valuesRow.getCell(2).value = courseName;
    valuesRow.getCell(3).value = courseHours;
    valuesRow.getCell(4).value = groupNum;
    valuesRow.getCell(5).value = dateRange;
    valuesRow.getCell(6).value = teacherName;
    valuesRow.eachCell(cell => {
      cell.font = { size: 10, name: 'Arial' };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'thin' }, bottom: { style: 'thin' },
        left: { style: 'thin' }, right: { style: 'thin' }
      };
    });
    valuesRow.height = 25;
    currentRow++;

    // Student rows
    students.forEach((student, idx) => {
      const studentRow = sheet.getRow(currentRow);
      studentRow.getCell(1).value = idx + 1;
      studentRow.getCell(2).value = student.fullName;
      studentRow.getCell(4).value = student.rank;
      studentRow.getCell(6).value = 'İlkin';
      studentRow.eachCell(cell => {
        cell.font = { size: 10, name: 'Arial' };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin' }, bottom: { style: 'thin' },
          left: { style: 'thin' }, right: { style: 'thin' }
        };
      });
      currentRow++;
    });

    groupCounter++;

    // Empty row after each group
    sheet.getRow(currentRow).height = 10;
    currentRow++;
  });

  // Generate and download
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

export function getUniqueCourseGroups(records) {
  const groups = groupRecordsByCourseAndDate(records);
  return groups.map(g => ({
    courseCode: g.courseCode,
    startDate: g.startDate,
    finishDate: g.finishDate,
    studentCount: g.students.length,
    courseName: getCourseName(g.courseCode),
  }));
}
