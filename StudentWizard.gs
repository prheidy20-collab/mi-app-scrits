function createStudentFromWizard(formData) {
  if (!formData) {
    throw new Error('No student information was received.');
  }

  const fullName = String(formData.fullName || '').trim();
  const studentNumber = String(formData.studentNumber || '').trim();
  const last4 = String(formData.last4 || '').trim();
  const phone = String(formData.phone || '').trim();
  const email = String(formData.email || '').trim();
  const address = String(formData.address || '').trim();
  const pin = String(formData.pin || '').trim();

  if (!fullName || !studentNumber || !last4 || !email || !pin) {
    throw new Error(
      'Complete full name, student number, last 4 SSN, email, and PIN.'
    );
  }

  if (!/^\d{4}$/.test(last4)) {
    throw new Error('Last 4 SSN must contain exactly four numbers.');
  }

  if (!/^\d{4,8}$/.test(pin)) {
    throw new Error('PIN must contain 4 to 8 numbers.');
  }

  const studentId = 'STU-' + studentNumber;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName('Students');
  const profilesSheet = ss.getSheetByName('Student Profiles');

  if (!studentsSheet || !profilesSheet) {
    throw new Error(
      'Students or Student Profiles sheet is missing.'
    );
  }

  const studentRows = studentsSheet.getDataRange().getValues();
  const profileRows = profilesSheet.getDataRange().getValues();

  const duplicateStudent = studentRows
    .slice(1)
    .some(row =>
      String(row[0]) === studentId ||
      String(row[2]).toLowerCase() === email.toLowerCase()
    );

  if (duplicateStudent) {
    throw new Error(
      'A clock account already exists for this student or email.'
    );
  }

  const duplicateProfile = profileRows
    .slice(1)
    .some(row =>
      String(row[0]) === studentId ||
      String(row[2]) === studentNumber ||
      String(row[4]).toLowerCase() === email.toLowerCase()
    );

  if (duplicateProfile) {
    throw new Error(
      'A student profile already exists with this student number or email.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const salt = Utilities.getUuid();
    const pinHash = hashText_(salt + pin);
    const now = new Date();

    studentsSheet.appendRow([
      studentId,
      fullName,
      email,
      salt,
      pinHash,
      true
    ]);

    profilesSheet.appendRow([
      studentId,
      fullName,
      studentNumber,
      last4,
      email,
      phone,
      address,
      true,
      now,
      now,
      'Created through Student Wizard'
    ]);

    return {
      ok: true,
      studentId: studentId,
      message:
        fullName +
        ' was created successfully and is now available for clock-in, schedules, assignments, and reports.'
    };

  } finally {
    lock.releaseLock();
  }
}