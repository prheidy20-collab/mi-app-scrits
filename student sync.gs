function syncStudentsToProfiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const studentsSheet = ss.getSheetByName('Students');
  const profilesSheet = ss.getSheetByName('Student Profiles');

  if (!studentsSheet) {
    throw new Error('The Students sheet was not found.');
  }

  if (!profilesSheet) {
    throw new Error(
      'The Student Profiles sheet was not found. Run setupStudentProfilesModule first.'
    );
  }

  const students = studentsSheet.getDataRange().getValues();
  const profiles = profilesSheet.getDataRange().getValues();

  const existingProfileIds = new Set(
    profiles
      .slice(1)
      .map(row => String(row[0] || '').trim())
      .filter(Boolean)
  );

  let added = 0;

  students.slice(1).forEach(row => {
    const studentId = String(row[0] || '').trim();
    const fullName = String(row[1] || '').trim();
    const email = String(row[2] || '').trim();

    const active =
      row[5] === true ||
      String(row[5]).toLowerCase() === 'true';

    if (
      !studentId ||
      !fullName ||
      existingProfileIds.has(studentId)
    ) {
      return;
    }

    const now = new Date();

    profilesSheet.appendRow([
      studentId,       // Student ID
      fullName,        // Full Name
      '',              // University Student Number
      '',              // Last 4 SSN
      email,           // Email
      '',              // Phone
      '',              // Home Address
      active,          // Active
      now,             // Created On
      now,             // Updated On
      'Imported automatically from the Students sheet.'
    ]);

    existingProfileIds.add(studentId);
    added++;
  });

  return added + ' student profile(s) created.';
}