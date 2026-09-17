function getStudentsForSchedule() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Student Profiles');

  if (!sheet) {
    throw new Error('The Student Profiles sheet was not found.');
  }

  const rows = sheet.getDataRange().getValues();

  return rows
    .slice(1)
    .filter(row =>
      row[7] === true ||
      String(row[7]).toLowerCase() === 'true'
    )
    .map(row => ({
      studentId: String(row[0] || ''),
      studentName: String(row[1] || ''),
      studentEmail: String(row[4] || '')
    }))
    .filter(student =>
      student.studentId &&
      student.studentName
    )
    .sort((a, b) =>
      a.studentName.localeCompare(b.studentName)
    );
}