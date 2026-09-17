function addPayrollShiftAdjustment(formData) {
  if (!formData) {
    throw new Error('No adjustment information was received.');
  }

  const studentId =
    String(formData.studentId || '').trim();

  const workDateText =
    String(formData.workDate || '').trim();

  const startTimeText =
    String(formData.startTime || '').trim();

  const endTimeText =
    String(formData.endTime || '').trim();

  const reason =
    String(formData.reason || '').trim();

  if (
    !studentId ||
    !workDateText ||
    !startTimeText ||
    !endTimeText ||
    !reason
  ) {
    throw new Error('Complete all required fields.');
  }

  const student =
    getScheduleRequestStudent_(studentId);

  const workDate =
    parseScheduleRequestDate_(workDateText);

  const startDateTime =
    combineScheduleRequestDateTime_(
      workDate,
      startTimeText
    );

  const endDateTime =
    combineScheduleRequestDateTime_(
      workDate,
      endTimeText
    );

  if (endDateTime <= startDateTime) {
    throw new Error(
      'Ending time must be later than starting time.'
    );
  }

  const hours =
    Math.round(
      (
        (
          endDateTime.getTime() -
          startDateTime.getTime()
        ) /
        3600000
      ) * 100
    ) / 100;

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      'Payroll Shift Adjustments'
    );

  if (!sheet) {
    sheet =
      ss.insertSheet(
        'Payroll Shift Adjustments'
      );

    sheet.appendRow([
      'Adjustment ID',
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

  sheet.appendRow([
    Utilities.getUuid(),
    student.studentId,
    student.studentName,
    workDate,
    startDateTime,
    endDateTime,
    hours,
    reason,
    getSettings_().COORDINATOR_EMAIL ||
      'Administrator',
    new Date(),
    true
  ]);

  markPayrollForReview_(
  student.studentId,
  workDate
);
  
  return {
    ok: true,
    studentName:
      student.studentName,
    hours: hours,
    message:
      'Shift adjustment added for ' +
      student.studentName +
      ' (' +
      hours.toFixed(2) +
      ' hours).'
  };
}


function getPayrollShiftAdjustments(
  monthText,
  studentId
) {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        'Payroll Shift Adjustments'
      );

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const tz =
    Session.getScriptTimeZone();

  return sheet
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(row => {
      const active =
        row[10] === true ||
        String(row[10]).toLowerCase() ===
          'true';

      if (!active) {
        return false;
      }

      if (
        String(row[1]) !==
        String(studentId)
      ) {
        return false;
      }

      const workDate =
        row[3];

      if (!(workDate instanceof Date)) {
        return false;
      }

      const rowMonth =
        Utilities.formatDate(
          workDate,
          tz,
          'yyyy-MM'
        );

      return rowMonth === monthText;
    })
    .map(row => ({
      adjustmentId:
        String(row[0]),

      studentId:
        String(row[1]),

      studentName:
        String(row[2]),

      workDate:
        Utilities.formatDate(
          row[3],
          tz,
          'MMM d, yyyy'
        ),

      startTime:
        Utilities.formatDate(
          row[4],
          tz,
          'h:mm a'
        ),

      endTime:
        Utilities.formatDate(
          row[5],
          tz,
          'h:mm a'
        ),

      hours:
        Number(row[6] || 0),

      reason:
        String(row[7] || '')
    }));
}

function markPayrollForReview_(studentId, workDate) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Payroll Approvals');

  if (!sheet || sheet.getLastRow() < 2) {
    return;
  }

  const tz = Session.getScriptTimeZone();

  const adjustmentMonth =
    Utilities.formatDate(
      workDate,
      tz,
      'yyyy-MM'
    );

  const rows =
    sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {

    const rowStudentId =
      String(rows[i][1] || '').trim();

    let rowMonth = '';

    const payrollMonth = rows[i][3];

    if (payrollMonth instanceof Date) {
      rowMonth =
        Utilities.formatDate(
          payrollMonth,
          tz,
          'yyyy-MM'
        );
    } else {
      rowMonth =
        String(payrollMonth || '')
          .trim()
          .substring(0, 7);
    }

    if (
      rowStudentId === String(studentId).trim() &&
      rowMonth === adjustmentMonth
    ) {
      sheet
        .getRange(i + 1, 10)
        .setValue('NEEDS REVIEW');

      sheet
        .getRange(i + 1, 13)
        .setValue(
          'Hours changed after payroll approval. Re-approval required.'
        );
    }
  }
}
  function getPayrollShiftAdjustmentTotals_(monthText) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Payroll Shift Adjustments');

  const totals = {};

  if (!sheet || sheet.getLastRow() < 2) {
    return totals;
  }

  const tz = Session.getScriptTimeZone();

  const rows = sheet
    .getDataRange()
    .getValues()
    .slice(1);

  rows.forEach(row => {
    const studentId =
      String(row[1] || '').trim();

    const workDate =
      row[3];

    const hours =
      Number(row[6]);

    const active =
      row[10] === true ||
      String(row[10]).toLowerCase() === 'true';

    if (
      !active ||
      !studentId ||
      !(workDate instanceof Date) ||
      !Number.isFinite(hours)
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

    totals[studentId] =
      (totals[studentId] || 0) +
      hours;
  });

  return totals;
}