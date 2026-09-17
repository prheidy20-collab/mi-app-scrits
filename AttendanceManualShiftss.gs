/**
 * FIESTA HUB
 * Manual Attendance Shifts
 *
 * Used to assign actual date/time details
 * to approved hours that were not captured
 * by a tablet clock-in/out.
 */


function getAttendanceManualShiftsSheet_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      'Attendance Manual Shifts'
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        'Attendance Manual Shifts'
      );

    sheet.appendRow([
      'Shift ID',
      'Student ID',
      'Student Name',
      'Work Date',
      'Start Time',
      'End Time',
      'Hours',
      'Reason',
      'Created By',
      'Created On',
      'Active'
    ]);

    sheet.setFrozenRows(1);

    sheet
      .getRange('D:D')
      .setNumberFormat('m/d/yyyy');

    sheet
      .getRange('E:F')
      .setNumberFormat('h:mm AM/PM');

    sheet
      .getRange('G:G')
      .setNumberFormat('0.00');

    sheet
      .getRange('J:J')
      .setNumberFormat(
        'm/d/yyyy h:mm AM/PM'
      );
  }

  return sheet;
}


function saveAttendanceManualShift(formData) {

  if (!formData) {
    throw new Error(
      'Shift information is required.'
    );
  }

  const studentId =
    String(
      formData.studentId || ''
    ).trim();

  const workDateText =
    String(
      formData.workDate || ''
    ).trim();

  const startText =
    String(
      formData.startTime || ''
    ).trim();

  const endText =
    String(
      formData.endTime || ''
    ).trim();

  const reason =
    String(
      formData.reason || ''
    ).trim();


  if (
    !studentId ||
    !workDateText ||
    !startText ||
    !endText ||
    !reason
  ) {
    throw new Error(
      'Complete all required fields.'
    );
  }


  const student =
    getStudentProfileDetail(
      studentId
    );

  if (!student) {
    throw new Error(
      'Student not found.'
    );
  }


  const workDate =
    parseAttendanceManualDate_(
      workDateText
    );


  const startTime =
    combineAttendanceManualDateTime_(
      workDate,
      startText
    );

  const endTime =
    combineAttendanceManualDateTime_(
      workDate,
      endText
    );


  if (
    endTime.getTime() <=
    startTime.getTime()
  ) {
    throw new Error(
      'Salida must be after Entrada.'
    );
  }

assertNoAttendanceOverlap_(
  studentId,
  startTime,
  endTime
);


  const hours =
    Math.round(
      (
        (
          endTime.getTime() -
          startTime.getTime()
        ) / 3600000
      ) * 100
    ) / 100;


  // ---------------------------------
  // SAFETY CHECK AGAINST APPROVED HOURS
  // ---------------------------------

  const monthText =
    Utilities.formatDate(
      workDate,
      Session.getScriptTimeZone(),
      'yyyy-MM'
    );


  const editorData =
    getAttendanceEditorData(
      studentId,
      monthText
    );


  const availableHours =
    Number(
      editorData.unassignedHours || 0
    );


  if (
    hours >
    availableHours + 0.001
  ) {
    throw new Error(
      'This shift is ' +
      hours.toFixed(2) +
      ' hours, but only ' +
      availableHours.toFixed(2) +
      ' approved hours still need details.'
    );
  }


  const sheet =
    getAttendanceManualShiftsSheet_();


  sheet.appendRow([
    Utilities.getUuid(),
    studentId,
    student.fullName ||
      student.studentName ||
      '',
    workDate,
    startTime,
    endTime,
    hours,
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
  workDate
);

return {
    ok: true,

    message:
      'Attendance shift added.',

    studentId:
      studentId,

    hours:
      hours,

    month:
      monthText
  };
}

function resolveMissingAttendanceAsWorked(
  alertId
) {

  const cleanAlertId =
    String(alertId || '').trim();

  if (!cleanAlertId) {
    throw new Error(
      'Alert ID is required.'
    );
  }


  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);


  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const alertsSheet =
      ss.getSheetByName(
        'Attendance Alerts'
      );

    if (!alertsSheet) {
      throw new Error(
        'Attendance Alerts sheet was not found.'
      );
    }


    const rows =
      alertsSheet
        .getDataRange()
        .getValues();


    let alertRowNumber = -1;
    let alertRow = null;


    for (
      let i = 1;
      i < rows.length;
      i++
    ) {

      const rowAlertId =
        String(
          rows[i][0] || ''
        ).trim();

      if (
        rowAlertId === cleanAlertId
      ) {
        alertRowNumber = i + 1;
        alertRow = rows[i];
        break;
      }
    }


    if (
      alertRowNumber === -1 ||
      !alertRow
    ) {
      throw new Error(
        'Attendance alert was not found.'
      );
    }


    const alertType =
      String(
        alertRow[6] || ''
      )
        .trim()
        .toUpperCase();

    const status =
      String(
        alertRow[11] || ''
      )
        .trim()
        .toUpperCase();


    if (
      alertType !== 'MISSING'
    ) {
      throw new Error(
        'Only MISSING attendance alerts can be marked as worked.'
      );
    }


    if (
      status !== 'OPEN'
    ) {
      throw new Error(
        'This attendance alert has already been reviewed.'
      );
    }


    const studentId =
      String(
        alertRow[1] || ''
      ).trim();

    const studentName =
      String(
        alertRow[2] || ''
      ).trim();

    const scheduledDate =
      alertRow[4];


    if (
      !studentId ||
      !(scheduledDate instanceof Date)
    ) {
      throw new Error(
        'The attendance alert does not contain a valid student or scheduled date.'
      );
    }


    const tz =
      Session.getScriptTimeZone();

    const dateKey =
      Utilities.formatDate(
        scheduledDate,
        tz,
        'yyyy-MM-dd'
      );


    const expectedSchedule =
      getExpectedScheduleForDate_(
        studentId,
        dateKey
      );


    if (
      !expectedSchedule ||
      !(expectedSchedule.start instanceof Date) ||
      !(expectedSchedule.end instanceof Date)
    ) {
      throw new Error(
        'The expected schedule for this date could not be found.'
      );
    }


    const startTime =
      expectedSchedule.start;

    const endTime =
      expectedSchedule.end;


    if (
      endTime.getTime() <=
      startTime.getTime()
    ) {
      throw new Error(
        'The expected schedule has invalid hours.'
      );
    }


    // Prevent duplicate or overlapping attendance.
    assertNoAttendanceOverlap_(
      studentId,
      startTime,
      endTime
    );


    const hours =
      Math.round(
        (
          (
            endTime.getTime() -
            startTime.getTime()
          ) / 3600000
        ) * 100
      ) / 100;


    const manualSheet =
      getAttendanceManualShiftsSheet_();


    const coordinator =
      getSettings_()
        .COORDINATOR_EMAIL ||
      Session
        .getActiveUser()
        .getEmail() ||
      'Administrator';


    manualSheet.appendRow([
      Utilities.getUuid(),
      studentId,
      studentName,
      scheduledDate,
      startTime,
      endTime,
      hours,
      'Worked - forgot to clock in/out',
      coordinator,
      new Date(),
      true
    ]);


    SpreadsheetApp.flush();


    // Resolve the original MISSING alert.
    alertsSheet
      .getRange(
        alertRowNumber,
        12
      )
      .setValue('WORKED');

    alertsSheet
      .getRange(
        alertRowNumber,
        13
      )
      .setValue(
        'Student worked scheduled shift but did not clock in/out.'
      );

    alertsSheet
      .getRange(
        alertRowNumber,
        14
      )
      .setValue(
        coordinator
      );

    alertsSheet
      .getRange(
        alertRowNumber,
        15
      )
      .setValue(
        new Date()
      );


    SpreadsheetApp.flush();


    markPayrollForReview_(
      studentId,
      scheduledDate
    );


    return {
      ok: true,

      message:
        'Attendance was marked as worked and the scheduled hours were added.',

      studentId:
        studentId,

      studentName:
        studentName,

      workDate:
        dateKey,

      hours:
        hours
    };


  } finally {

    lock.releaseLock();

  }
}

function resolveMissingAttendanceAsWorked(
  alertId
) {

  const cleanAlertId =
    String(alertId || '').trim();

  if (!cleanAlertId) {
    throw new Error(
      'Alert ID is required.'
    );
  }


  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);


  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const alertsSheet =
      ss.getSheetByName(
        'Attendance Alerts'
      );

    if (!alertsSheet) {
      throw new Error(
        'Attendance Alerts sheet was not found.'
      );
    }


    const rows =
      alertsSheet
        .getDataRange()
        .getValues();


    let alertRowNumber = -1;
    let alertRow = null;


    for (
      let i = 1;
      i < rows.length;
      i++
    ) {

      const rowAlertId =
        String(
          rows[i][0] || ''
        ).trim();

      if (
        rowAlertId === cleanAlertId
      ) {
        alertRowNumber = i + 1;
        alertRow = rows[i];
        break;
      }
    }


    if (
      alertRowNumber === -1 ||
      !alertRow
    ) {
      throw new Error(
        'Attendance alert was not found.'
      );
    }


    const alertType =
      String(
        alertRow[6] || ''
      )
        .trim()
        .toUpperCase();

    const status =
      String(
        alertRow[11] || ''
      )
        .trim()
        .toUpperCase();


    if (
      alertType !== 'MISSING'
    ) {
      throw new Error(
        'Only MISSING attendance alerts can be marked as worked.'
      );
    }


    if (
      status !== 'OPEN'
    ) {
      throw new Error(
        'This attendance alert has already been reviewed.'
      );
    }


    const studentId =
      String(
        alertRow[1] || ''
      ).trim();

    const studentName =
      String(
        alertRow[2] || ''
      ).trim();

    const scheduledDate =
      alertRow[4];


    if (
      !studentId ||
      !(scheduledDate instanceof Date)
    ) {
      throw new Error(
        'The attendance alert does not contain a valid student or scheduled date.'
      );
    }


    const tz =
      Session.getScriptTimeZone();

    const dateKey =
      Utilities.formatDate(
        scheduledDate,
        tz,
        'yyyy-MM-dd'
      );


    const expectedSchedule =
      getExpectedScheduleForDate_(
        studentId,
        dateKey
      );


    if (
      !expectedSchedule ||
      !(expectedSchedule.start instanceof Date) ||
      !(expectedSchedule.end instanceof Date)
    ) {
      throw new Error(
        'The expected schedule for this date could not be found.'
      );
    }


    const startTime =
      expectedSchedule.start;

    const endTime =
      expectedSchedule.end;


    if (
      endTime.getTime() <=
      startTime.getTime()
    ) {
      throw new Error(
        'The expected schedule has invalid hours.'
      );
    }


    // Prevent duplicate or overlapping attendance.
    assertNoAttendanceOverlap_(
      studentId,
      startTime,
      endTime
    );


    const hours =
      Math.round(
        (
          (
            endTime.getTime() -
            startTime.getTime()
          ) / 3600000
        ) * 100
      ) / 100;


    const manualSheet =
      getAttendanceManualShiftsSheet_();


    const coordinator =
      getSettings_()
        .COORDINATOR_EMAIL ||
      Session
        .getActiveUser()
        .getEmail() ||
      'Administrator';


    manualSheet.appendRow([
      Utilities.getUuid(),
      studentId,
      studentName,
      scheduledDate,
      startTime,
      endTime,
      hours,
      'Worked - forgot to clock in/out',
      coordinator,
      new Date(),
      true
    ]);


    SpreadsheetApp.flush();


    // Resolve the original MISSING alert.
    alertsSheet
      .getRange(
        alertRowNumber,
        12
      )
      .setValue('WORKED');

    alertsSheet
      .getRange(
        alertRowNumber,
        13
      )
      .setValue(
        'Student worked scheduled shift but did not clock in/out.'
      );

    alertsSheet
      .getRange(
        alertRowNumber,
        14
      )
      .setValue(
        coordinator
      );

    alertsSheet
      .getRange(
        alertRowNumber,
        15
      )
      .setValue(
        new Date()
      );


    SpreadsheetApp.flush();


    markPayrollForReview_(
      studentId,
      scheduledDate
    );


    return {
      ok: true,

      message:
        'Attendance was marked as worked and the scheduled hours were added.',

      studentId:
        studentId,

      studentName:
        studentName,

      workDate:
        dateKey,

      hours:
        hours
    };


  } finally {

    lock.releaseLock();

  }
}

function getAttendanceManualShifts_(
  studentId,
  monthText
) {

  const sheet =
    getAttendanceManualShiftsSheet_();

  const results = [];

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {
    return results;
  }


  const tz =
    Session.getScriptTimeZone();


  const rows =
    sheet
      .getDataRange()
      .getValues()
      .slice(1);


  rows.forEach(function(row) {

    const rowStudentId =
      String(
        row[1] || ''
      ).trim();

    const workDate =
      row[3];

    const active =
      row[10] === true ||
      String(
        row[10] || ''
      ).toLowerCase() ===
        'true';


    if (
      !active ||
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


    if (
      rowMonth !== monthText
    ) {
      return;
    }


    results.push({
      shiftId:
        String(row[0] || ''),

      studentId:
        rowStudentId,

      studentName:
        String(row[2] || ''),

      workDate:
        row[3],

      startTime:
        row[4],

      endTime:
        row[5],

      hours:
        Number(row[6] || 0),

      reason:
        String(row[7] || '')
    });
  });


  return results;
}


function parseAttendanceManualDate_(
  dateText
) {

  const parts =
    String(dateText)
      .split('-');

  if (
    parts.length !== 3
  ) {
    throw new Error(
      'Invalid work date.'
    );
  }


  const year =
    Number(parts[0]);

  const month =
    Number(parts[1]) - 1;

  const day =
    Number(parts[2]);


  const date =
    new Date(
      year,
      month,
      day
    );


  if (
    Number.isNaN(
      date.getTime()
    )
  ) {
    throw new Error(
      'Invalid work date.'
    );
  }


  return date;
}


function combineAttendanceManualDateTime_(
  workDate,
  timeText
) {

  const parts =
    String(timeText)
      .split(':');


  if (
    parts.length < 2
  ) {
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
    workDate.getFullYear(),
    workDate.getMonth(),
    workDate.getDate(),
    hour,
    minute,
    0,
    0
  );
}


function setupAttendanceManualShifts() {
  getAttendanceManualShiftsSheet_();
}
/**
 * Updates an existing manual attendance shift.
 */
function updateAttendanceManualShift(formData) {

  if (!formData) {
    throw new Error('Shift information is required.');
  }

  const shiftId =
    String(formData.shiftId || '').trim();

  const workDateText =
    String(formData.workDate || '').trim();

  const startText =
    String(formData.startTime || '').trim();

  const endText =
    String(formData.endTime || '').trim();

  const reason =
    String(formData.reason || '').trim();


  if (
    !shiftId ||
    !workDateText ||
    !startText ||
    !endText ||
    !reason
  ) {
    throw new Error(
      'Complete all required fields.'
    );
  }


  const sheet =
    getAttendanceManualShiftsSheet_();

  const values =
    sheet.getDataRange().getValues();

  let rowNumber = -1;
  let studentId = '';


  for (let i = 1; i < values.length; i++) {

    if (
      String(values[i][0] || '').trim() ===
      shiftId
    ) {
      rowNumber = i + 1;

      studentId =
        String(values[i][1] || '').trim();

      break;
    }
  }


  if (rowNumber === -1) {
    throw new Error(
      'Manual attendance shift was not found.'
    );
  }


  const workDate =
    parseAttendanceManualDate_(
      workDateText
    );

  const startTime =
    combineAttendanceManualDateTime_(
      workDate,
      startText
    );

  const endTime =
    combineAttendanceManualDateTime_(
      workDate,
      endText
    );


  if (
    endTime.getTime() <=
    startTime.getTime()
  ) {
    throw new Error(
      'Salida must be after Entrada.'
    );
  }

assertNoAttendanceOverlap_(
  studentId,
  startTime,
  endTime,
  {
    ignoreShiftId: shiftId
  }
);

  const hours =
    Math.round(
      (
        (
          endTime.getTime() -
          startTime.getTime()
        ) / 3600000
      ) * 100
    ) / 100;


  // Update only the editable fields.
  sheet
    .getRange(rowNumber, 4)
    .setValue(workDate);

  sheet
    .getRange(rowNumber, 5)
    .setValue(startTime);

  sheet
    .getRange(rowNumber, 6)
    .setValue(endTime);

  sheet
    .getRange(rowNumber, 7)
    .setValue(hours);

  sheet
    .getRange(rowNumber, 8)
    .setValue(reason);

  // Update timestamp
  sheet
    .getRange(rowNumber, 10)
    .setValue(new Date());


  SpreadsheetApp.flush();

markPayrollForReview_(
  studentId,
  workDate
);

return {
    ok: true,
    message: 'Manual attendance shift updated.',
    shiftId: shiftId,
    studentId: studentId,
    hours: hours
  };
}


/**
 * Soft-deletes a manual attendance shift.
 *
 * The row remains for audit history,
 * but Active becomes FALSE.
 */
function deleteAttendanceManualShift(shiftId) {

  shiftId =
    String(shiftId || '').trim();


  if (!shiftId) {
    throw new Error(
      'Shift ID is required.'
    );
  }


  const sheet =
    getAttendanceManualShiftsSheet_();

  const values =
    sheet.getDataRange().getValues();


  for (let i = 1; i < values.length; i++) {

    if (
      String(values[i][0] || '').trim() ===
      shiftId
    ) {

      // Column K = Active
      sheet
        .getRange(i + 1, 11)
        .setValue(false);

      SpreadsheetApp.flush();

markPayrollForReview_(
  studentId,
  workDate
);

return {
        ok: true,
        message:
          'Manual attendance shift removed.'
      };
    }
  }


  throw new Error(
    'Manual attendance shift was not found.'
  );
}