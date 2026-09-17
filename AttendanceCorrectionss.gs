/**
 * FIESTA HUB
 * Attendance Corrections
 *
 * Preserves original tablet punches while allowing
 * coordinator-approved times for official attendance.
 */


/**
 * Creates Attendance Corrections sheet if needed.
 */
function getAttendanceCorrectionsSheet_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      'Attendance Corrections'
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        'Attendance Corrections'
      );

    sheet.appendRow([
      'Correction ID',
      'Entry ID',
      'Student ID',
      'Student Name',
      'Work Date',
      'Original In',
      'Original Out',
      'Approved In',
      'Approved Out',
      'Approved Hours',
      'Reason',
      'Corrected By',
      'Corrected On',
      'Active'
    ]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}


/**
 * Returns active attendance corrections
 * for one student/month.
 */
function getAttendanceCorrections_(
  studentId,
  monthText
) {

  const sheet =
    getAttendanceCorrectionsSheet_();

  const corrections = {};

  if (sheet.getLastRow() < 2) {
    return corrections;
  }

  const tz =
    Session.getScriptTimeZone();

  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);

  rows.forEach(function(row) {

    const entryId =
      String(row[1] || '').trim();

    const rowStudentId =
      String(row[2] || '').trim();

    const workDate =
      row[4];

    const active =
      row[13] === true ||
      String(row[13] || '')
        .toLowerCase() === 'true';

    if (
      !active ||
      !entryId ||
      rowStudentId !==
        String(studentId).trim() ||
      !(workDate instanceof Date)
    ) {
      return;
    }

    const rowMonth =
      Utilities.formatDate(
        workDate,
        tz,
        'yyyy-MM'
      );

    if (rowMonth !== monthText) {
      return;
    }

    corrections[entryId] = {
      correctionId:
        String(row[0] || ''),

      entryId:
        entryId,

      studentId:
        rowStudentId,

      studentName:
        String(row[3] || ''),

      workDate:
        workDate,

      originalIn:
        row[5],

      originalOut:
        row[6],

      approvedIn:
        row[7],

      approvedOut:
        row[8],

      approvedHours:
        Number(row[9] || 0),

      reason:
        String(row[10] || '')
    };
  });

  return corrections;
}


/**
 * Saves a coordinator correction for a tablet entry.
 *
 * IMPORTANT:
 * Original Time Entries row is NOT changed.
 */
function saveAttendanceCorrection(formData) {

  if (!formData) {
    throw new Error(
      'Correction information is required.'
    );
  }

  const entryId =
    String(
      formData.entryId || ''
    ).trim();

  const studentId =
    String(
      formData.studentId || ''
    ).trim();

  const approvedInText =
    String(
      formData.approvedIn || ''
    ).trim();

  const approvedOutText =
    String(
      formData.approvedOut || ''
    ).trim();

  const reason =
    String(
      formData.reason || ''
    ).trim();


  if (!entryId || !studentId) {
    throw new Error(
      'Entry ID and Student ID are required.'
    );
  }

  if (
    !approvedInText ||
    !approvedOutText
  ) {
    throw new Error(
      'Approved Entrada and Salida are required.'
    );
  }

  if (!reason) {
    throw new Error(
      'Please enter a reason for the correction.'
    );
  }


  // ----------------------------------
  // FIND ORIGINAL TABLET ENTRY
  // ----------------------------------

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const timeSheet =
    ss.getSheetByName('Time Entries');

  if (
    !timeSheet ||
    timeSheet.getLastRow() < 2
  ) {
    throw new Error(
      'Time Entries sheet was not found.'
    );
  }

  const rows =
    timeSheet
      .getDataRange()
      .getValues();

  let original = null;

  for (
    let i = 1;
    i < rows.length;
    i++
  ) {

    if (
      String(rows[i][0] || '').trim()
        === entryId
    ) {

      original = {
        entryId:
          String(rows[i][0] || ''),

        studentId:
          String(rows[i][1] || ''),

        studentName:
          String(rows[i][2] || ''),

        clockIn:
          rows[i][3],

        clockOut:
          rows[i][4]
      };

      break;
    }
  }


  if (!original) {
    throw new Error(
      'Original tablet entry was not found.'
    );
  }


  if (
    original.studentId !== studentId
  ) {
    throw new Error(
      'Student does not match the original entry.'
    );
  }


  if (
    !(original.clockIn instanceof Date)
  ) {
    throw new Error(
      'Original Clock In is invalid.'
    );
  }


  // ----------------------------------
  // BUILD APPROVED TIMES
  // ----------------------------------

  const year =
    original.clockIn.getFullYear();

  const month =
    original.clockIn.getMonth();

  const day =
    original.clockIn.getDate();


  const approvedIn =
    buildAttendanceCorrectionTime_(
      year,
      month,
      day,
      approvedInText
    );

  const approvedOut =
    buildAttendanceCorrectionTime_(
      year,
      month,
      day,
      approvedOutText
    );


  if (
    approvedOut.getTime() <=
    approvedIn.getTime()
  ) {
    throw new Error(
      'Approved Salida must be after Approved Entrada.'
    );
  }
assertNoAttendanceOverlap_(
  studentId,
  approvedIn,
  approvedOut,
  {
    ignoreEntryId: entryId
  }
);

  const approvedHours =
    Math.round(
      (
        (
          approvedOut.getTime() -
          approvedIn.getTime()
        ) / 3600000
      ) * 100
    ) / 100;


  // ----------------------------------
  // SAVE CORRECTION
  // ----------------------------------

  const sheet =
    getAttendanceCorrectionsSheet_();

  const existing =
    sheet.getDataRange().getValues();


  // Deactivate previous correction
  // for the same Time Entry.
  for (
    let i = 1;
    i < existing.length;
    i++
  ) {

    const existingEntryId =
      String(
        existing[i][1] || ''
      ).trim();

    const active =
      existing[i][13] === true ||
      String(
        existing[i][13] || ''
      ).toLowerCase() === 'true';

    if (
      existingEntryId === entryId &&
      active
    ) {

      sheet
        .getRange(
          i + 1,
          14
        )
        .setValue(false);
    }
  }


  sheet.appendRow([
    Utilities.getUuid(),
    original.entryId,
    original.studentId,
    original.studentName,
    original.clockIn,
    original.clockIn,
    original.clockOut,
    approvedIn,
    approvedOut,
    approvedHours,
    reason,
    getSettings_()
      .COORDINATOR_EMAIL ||
      'Administrator',
    new Date(),
    true
  ]);


SpreadsheetApp.flush();

markPayrollForReview_(
  studentId,
  approvedIn
);

return {
    ok: true,

    message:
      'Attendance correction saved.',

    entryId:
      entryId,

    approvedHours:
      approvedHours
  };
}


/**
 * Converts HTML time such as 15:00
 * into a Date on the original work date.
 */
function buildAttendanceCorrectionTime_(
  year,
  month,
  day,
  timeText
) {

  const parts =
    String(timeText)
      .split(':');

  if (parts.length < 2) {
    throw new Error(
      'Invalid time.'
    );
  }

  const hour =
    Number(parts[0]);

  const minute =
    Number(parts[1]);

  if (
    !Number.isInteger(hour) ||
    !Number.isInteger(minute) ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {
    throw new Error(
      'Invalid time.'
    );
  }

  return new Date(
    year,
    month,
    day,
    hour,
    minute,
    0,
    0
  );
  
}function setupAttendanceCorrections() {
  getAttendanceCorrectionsSheet_();
}

function assertNoAttendanceOverlap_(
  studentId,
  startTime,
  endTime,
  options
) {

  options = options || {};

  const ignoreEntryId =
    String(options.ignoreEntryId || '').trim();

  const ignoreShiftId =
    String(options.ignoreShiftId || '').trim();

  const tz =
    Session.getScriptTimeZone();

  const monthText =
    Utilities.formatDate(
      startTime,
      tz,
      'yyyy-MM'
    );

  const dateKey =
    Utilities.formatDate(
      startTime,
      tz,
      'yyyy-MM-dd'
    );

  const attendance =
    getStudentMonthlyAttendanceData_(
      studentId,
      monthText
    );

  for (
    let i = 0;
    i < attendance.shifts.length;
    i++
  ) {

    const shift =
      attendance.shifts[i];

    if (
      ignoreEntryId &&
      String(shift.entryId || '').trim() ===
        ignoreEntryId
    ) {
      continue;
    }

    if (
      ignoreShiftId &&
      String(shift.shiftId || '').trim() ===
        ignoreShiftId
    ) {
      continue;
    }

    if (
      !(shift.startTime instanceof Date) ||
      !(shift.endTime instanceof Date)
    ) {
      continue;
    }

    const shiftDate =
      Utilities.formatDate(
        shift.startTime,
        tz,
        'yyyy-MM-dd'
      );

    if (shiftDate !== dateKey) {
      continue;
    }

    const overlaps =
      startTime.getTime() <
        shift.endTime.getTime() &&
      endTime.getTime() >
        shift.startTime.getTime();

    if (overlaps) {

      const existingStart =
        Utilities.formatDate(
          shift.startTime,
          tz,
          'h:mm a'
        );

      const existingEnd =
        Utilities.formatDate(
          shift.endTime,
          tz,
          'h:mm a'
        );

      throw new Error(
        'Overlapping hours detected. ' +
        'This student already has attendance from ' +
        existingStart +
        ' to ' +
        existingEnd +
        '. Please choose different times.'
      );
    }
  }
}